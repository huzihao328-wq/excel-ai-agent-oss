/*
 * run_script：专用工具覆盖不到的计算（复杂解析、模糊匹配、多条件规则、自定义排名、跨表逻辑……）
 * 让模型写一段 JS 在沙箱里算，数据在插件内读取、结果直接落表，不经过对话上下文。
 */
import { ToolResult } from "../types";
import { gridPreview, pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { loadTable, MAX_RESHAPE_CELLS, ReshapeContext, writeSheets } from "./reshapeTools";
import { checkScriptSource, runScript } from "./scriptSandbox";

const DEFAULT_TIMEOUT_S = 15;
const MAX_TIMEOUT_S = 60;
const MAX_INPUTS = 8;
/** 没有 output 时返回给模型的结果：数组最多保留这么多项，其余只报总数 */
const RETURN_ITEMS = 100;
const PREVIEW_ROWS = 12;

type WriteGrid = (sheet: string, startCell: string, grid: unknown[][]) => Promise<ToolResult>;

interface InputSpec {
  name: string;
  sheet: string;
  range?: string;
}

function parseInputs(raw: unknown): InputSpec[] {
  if (raw == null) return [];
  if (!Array.isArray(raw)) throw new Error("inputs 必须是数组，如 [{\"name\":\"orders\",\"sheet\":\"订单\"}]");
  if (raw.length > MAX_INPUTS) throw new Error(`inputs 最多 ${MAX_INPUTS} 项`);
  const seen = new Set<string>();
  return raw.map((item: any, i) => {
    const name = String(item?.name ?? "").trim();
    const sheet = String(item?.sheet ?? "").trim();
    if (!/^[A-Za-z_$][\w$]*$/.test(name)) throw new Error(`inputs[${i}].name 必须是合法的 JS 变量名（字母开头，如 orders、t1）`);
    if (seen.has(name)) throw new Error(`inputs 里的 name「${name}」重复`);
    if (!sheet) throw new Error(`inputs[${i}] 缺少 sheet`);
    seen.add(name);
    const range = typeof item?.range === "string" && item.range.trim() ? item.range.trim() : undefined;
    return { name, sheet, range };
  });
}

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === "object" && !Array.isArray(v) && !(v instanceof Date);
}

function cellValue(v: unknown): unknown {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return Number.isFinite(v) ? v : "";
  if (typeof v === "string" || typeof v === "boolean") return v;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? "" : v.toISOString().slice(0, 10);
  if (typeof v === "bigint") return Number(v);
  throw new Error("单元格只能是文本、数字、布尔值或 null，不能是对象或数组");
}

/** 脚本返回值 → 可写入的二维数组。支持二维数组、对象数组（自动加表头）、一维数组（写成一列）。 */
export function toOutputGrid(value: unknown): unknown[][] {
  if (!Array.isArray(value) || value.length === 0) {
    throw new Error("要写入表格时脚本必须返回非空数组：二维数组（第一行作表头）或对象数组");
  }
  let grid: unknown[][];
  if (value.every((row) => isPlainObject(row))) {
    const headers: string[] = [];
    const seen = new Set<string>();
    for (const row of value as Record<string, unknown>[]) {
      for (const key of Object.keys(row)) if (!seen.has(key)) { seen.add(key); headers.push(key); }
    }
    grid = [headers, ...(value as Record<string, unknown>[]).map((row) => headers.map((h) => row[h]))];
  } else if (value.every((row) => Array.isArray(row))) {
    grid = value as unknown[][];
  } else if (value.every((row) => !Array.isArray(row) && !isPlainObject(row))) {
    grid = value.map((v) => [v]);
  } else {
    throw new Error("返回的数组混合了行数组、对象和单个值，请统一成二维数组或对象数组");
  }
  const cols = grid.reduce((max, row) => Math.max(max, row.length), 0);
  if (!cols) throw new Error("返回的二维数组没有任何列");
  return grid.map((row, r) => Array.from({ length: cols }, (_, c) => {
    try {
      return cellValue(row[c]);
    } catch (e: any) {
      throw new Error(`第 ${r + 1} 行第 ${c + 1} 列：${e.message}`);
    }
  }));
}

/** 给模型看的返回值：大数组截断并注明总数，避免把整表灌进上下文 */
function summarizeValue(value: unknown): { result: unknown; note?: string } {
  if (Array.isArray(value) && value.length > RETURN_ITEMS) {
    return {
      result: value.slice(0, RETURN_ITEMS),
      note: `结果共 ${value.length} 项，这里只返回前 ${RETURN_ITEMS} 项。要完整落表请加 output 参数，不要把结果抄进 write_range。`,
    };
  }
  return { result: value };
}

export async function toolRunScript(args: any, ctx: ReshapeContext, writeGrid: WriteGrid): Promise<ToolResult> {
  const sourceError = checkScriptSource(args.code);
  if (sourceError) return { ok: false, error: sourceError };
  const inputsSpec = parseInputs(args.inputs);
  const output = isPlainObject(args.output) ? (args.output as Record<string, unknown>) : null;
  const newSheet = output && typeof output.new_sheet === "string" ? output.new_sheet.trim() : "";
  const targetSheet = output && typeof output.sheet === "string" ? output.sheet.trim() : "";
  if (output && !newSheet && !targetSheet) {
    return { ok: false, error: "output 需要 new_sheet（写到新建 Sheet）或 sheet + start_cell（写到已有位置）其中之一" };
  }
  if (newSheet && targetSheet) return { ok: false, error: "output 里 new_sheet 与 sheet 只能选一个" };

  const timeoutS = Math.min(MAX_TIMEOUT_S, Math.max(1, Number(args.timeout_seconds) || DEFAULT_TIMEOUT_S));

  const inputs: Record<string, unknown[][]> = {};
  const sources: string[] = [];
  let cells = 0;
  for (const spec of inputsSpec) {
    const table = await loadTable(spec.sheet, spec.range, false, ctx);
    const cols = table.rows[0]?.length ?? 0;
    cells += table.rows.length * cols;
    if (cells > MAX_RESHAPE_CELLS) {
      return { ok: false, error: `输入合计超过 ${MAX_RESHAPE_CELLS} 个单元格，请用 range 只读需要的列` };
    }
    inputs[spec.name] = table.rows;
    sources.push(`${spec.name}=${table.sheet}!${table.address}（${table.rows.length}×${cols}）`);
  }

  const outcome = await runScript({ code: String(args.code), inputs, timeoutMs: timeoutS * 1000 }, ctx.signal);
  if (ctx.signal?.aborted) throw new Error("用户已停止任务，未继续执行操作。");
  const logs = outcome.logs.length ? outcome.logs : undefined;
  if (!outcome.ok) {
    return { ok: false, error: `${outcome.error}${logs ? `\nlog 输出:\n${logs.join("\n")}` : ""}` };
  }

  if (!output) {
    const { result, note } = summarizeValue(outcome.value);
    return {
      ok: true,
      data: { inputs: sources, result, ...(note ? { note } : {}), ...(logs ? { logs } : {}), ms: outcome.ms },
      summary: `脚本已运行（${outcome.ms} ms）`,
    };
  }

  let grid: unknown[][];
  try {
    grid = toOutputGrid(outcome.value);
  } catch (e: any) {
    return { ok: false, error: `${e.message}。脚本本身已运行成功，但结果没有写入。${logs ? `\nlog 输出:\n${logs.join("\n")}` : ""}` };
  }
  const header = output.header !== false;

  if (targetSheet) {
    const startCell = typeof output.start_cell === "string" && output.start_cell.trim() ? output.start_cell.trim() : "A1";
    const written = await writeGrid(targetSheet, startCell, grid);
    if (!written.ok) return written;
    const data = isPlainObject(written.data) ? written.data : {};
    return { ...written, data: { ...data, inputs: sources, preview: grid.slice(0, PREVIEW_ROWS), ...(logs ? { logs } : {}) } };
  }

  const [sheet] = await writeSheets([{ baseName: newSheet, grid, formats: [], header }], ctx);
  const description = `脚本计算 → 新 Sheet「${sheet.name}」（${grid.length} 行）`;
  pushOpLog({ description, kind: "write", sheet: sheet.name, address: sheet.address, after: gridPreview(grid) });
  registerArtifact({
    name: sheet.name,
    kind: "range",
    sheet: sheet.name,
    address: sheet.address,
    brief: typeof args.purpose === "string" && args.purpose.trim() ? args.purpose.trim().slice(0, 60) : `脚本计算结果 ${grid.length} 行`,
  });
  return {
    ok: true,
    data: {
      dest_sheet: sheet.name,
      range: sheet.address,
      rows: grid.length,
      cols: grid[0].length,
      inputs: sources,
      preview: grid.slice(0, PREVIEW_ROWS),
      ...(logs ? { logs } : {}),
      note: "结果是静态值；源数据改动后需重新运行。",
    },
    summary: description,
    undoable: false,
  };
}
