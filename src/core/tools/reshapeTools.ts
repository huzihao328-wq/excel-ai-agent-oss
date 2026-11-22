/* global Excel */
import { AppSettings, ToolResult } from "../types";
import { gridPreview, pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { isInternalSheet } from "./checkpoint";
import { uniqueSheetName } from "./tableTools";
import { indexToColLetter, resolveColSpec } from "./formulaRefs";
import { keepTextAsText, normalizeNumberFormat } from "./safeValues";
import {
  AGG_LABELS,
  Aggregation,
  BAD_DATE_LABEL,
  BLANK_LABEL,
  DATE_PART_FORMATS,
  DatePart,
  diffTables,
  GroupColumn,
  groupSummary,
  MetricColumn,
  reconcileTables,
  splitByKey,
  stackTables,
  unpivotRows,
} from "./reshape";

/** 单次整表处理读入的单元格上限；更大的表请缩小 range 或用透视表 */
export const MAX_RESHAPE_CELLS = 1_000_000;
const READ_BLOCK_CELLS = 40_000;
const WRITE_BLOCK_CELLS = 20_000;
const FORMAT_SAMPLE_ROWS = 20;
const AUTOFIT_ROWS = 500;
const PREVIEW_ROWS = 12;
const DEFAULT_SPLIT_SHEETS = 30;
const MAX_SPLIT_SHEETS = 100;
const CONFIRM_SPLIT_OVER = 10;

export interface ReshapeContext {
  settings: AppSettings;
  signal?: AbortSignal;
  confirm: (title: string, detail: string) => Promise<boolean>;
}

function checkAbort(ctx: ReshapeContext): void {
  if (ctx.signal?.aborted) throw new Error("用户已停止任务，未继续执行操作。");
}

export interface LoadedTable {
  sheet: string;
  address: string;
  columnIndex: number;
  headerRow: unknown[];
  headers: string[];
  rows: unknown[][];
  /** 每列数据区第一个非「常规」数字格式 */
  formats: string[];
}

function pickColumnFormats(grid: unknown[][], cols: number): string[] {
  return Array.from({ length: cols }, (_, c) => {
    for (const row of grid) {
      const fmt = normalizeNumberFormat(row?.[c]);
      if (fmt !== "General") return String(row[c]);
    }
    return "";
  });
}

export async function loadTable(sheetName: string, rangeArg: unknown, hasHeaders: boolean, ctx: ReshapeContext): Promise<LoadedTable> {
  const name = String(sheetName || "").trim();
  if (!name) throw new Error("需要指定 Sheet");
  if (isInternalSheet(name)) throw new Error("不能读取检查点备份表。");
  const rangeText = typeof rangeArg === "string" ? rangeArg.trim() : "";
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) throw new Error(`找不到工作表「${name}」`);
    const used = rangeText
      ? sheet.getRange(rangeText).getIntersectionOrNullObject(sheet.getUsedRange(true))
      : sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
    await context.sync();
    if (used.isNullObject) throw new Error(rangeText ? `${name}!${rangeText} 里没有数据` : `${name} 是空表`);
    const rows = used.rowCount;
    const cols = used.columnCount;
    if (rows * cols > MAX_RESHAPE_CELLS) {
      throw new Error(
        `${name} 有 ${rows} 行 × ${cols} 列，超过单次 ${MAX_RESHAPE_CELLS} 个单元格的上限。请用 range 只选需要的列，或改用 create_pivot_table。`
      );
    }
    const grid: unknown[][] = [];
    const blockRows = Math.max(1, Math.floor(READ_BLOCK_CELLS / cols));
    for (let r = 0; r < rows; r += blockRows) {
      checkAbort(ctx);
      const block = sheet.getRangeByIndexes(used.rowIndex + r, used.columnIndex, Math.min(blockRows, rows - r), cols);
      block.load("values");
      await context.sync();
      for (const row of block.values as unknown[][]) grid.push(row);
    }
    const headerCount = hasHeaders ? 1 : 0;
    const dataRows = rows - headerCount;
    let formats: string[] = Array.from({ length: cols }, () => "");
    if (dataRows > 0) {
      const sample = sheet.getRangeByIndexes(used.rowIndex + headerCount, used.columnIndex, Math.min(FORMAT_SAMPLE_ROWS, dataRows), cols);
      sample.load("numberFormat");
      await context.sync();
      formats = pickColumnFormats(sample.numberFormat as unknown[][], cols);
    }
    const letterName = (c: number) => `列${indexToColLetter(used.columnIndex + c)}`;
    const headerRow = hasHeaders ? grid[0] : [];
    const headers = Array.from({ length: cols }, (_, c) =>
      hasHeaders ? String(headerRow[c] ?? "").trim() || letterName(c) : letterName(c)
    );
    return {
      sheet: name,
      address: used.address.split("!").pop() as string,
      columnIndex: used.columnIndex,
      headerRow,
      headers,
      rows: grid.slice(headerCount),
      formats,
    };
  });
}

function resolveColumn(table: LoadedTable, spec: unknown, label: string): number {
  return resolveColSpec(table.headers, spec, table.columnIndex, label);
}

function specList(value: unknown): unknown[] {
  if (Array.isArray(value)) return value.filter((v) => v != null && v !== "");
  return value == null || value === "" ? [] : [value];
}

export interface OutputSheet {
  baseName: string;
  grid: unknown[][];
  /** 每列数据行的数字格式；表头行不套格式 */
  formats: (string | undefined)[];
  /** 第一行是否为表头（加粗、冻结），默认 true */
  header?: boolean;
}

export interface WrittenSheet {
  name: string;
  address: string;
  rows: number;
}

export async function writeSheets(outputs: OutputSheet[], ctx: ReshapeContext): Promise<WrittenSheet[]> {
  const created: string[] = [];
  try {
    return await Excel.run(async (context) => {
      const written: WrittenSheet[] = [];
      for (const out of outputs) {
        checkAbort(ctx);
        const name = await uniqueSheetName(context, out.baseName);
        const ws = context.workbook.worksheets.add(name);
        created.push(name);
        const rows = out.grid.length;
        const cols = Math.max(1, out.grid.reduce((max, row) => Math.max(max, row.length), 0));
        const headerRows = out.header === false ? 0 : 1;
        const blockRows = Math.max(1, Math.floor(WRITE_BLOCK_CELLS / cols));
        for (let r0 = 0; r0 < rows; r0 += blockRows) {
          checkAbort(ctx);
          const n = Math.min(blockRows, rows - r0);
          const dataStart = Math.max(r0, headerRows);
          const dataRows = r0 + n - dataStart;
          if (dataRows > 0) {
            out.formats.forEach((fmt, c) => {
              if (!fmt || c >= cols) return;
              ws.getRangeByIndexes(dataStart, c, dataRows, 1).numberFormat = Array.from({ length: dataRows }, () => [fmt]);
            });
          }
          const block = out.grid.slice(r0, r0 + n).map((row) =>
            Array.from({ length: cols }, (_, c) => keepTextAsText(row[c] ?? ""))
          );
          ws.getRangeByIndexes(r0, 0, n, cols).values = block as any[][];
          await context.sync();
        }
        if (headerRows) {
          ws.getRangeByIndexes(0, 0, 1, cols).format.font.bold = true;
          if (rows > 1) ws.freezePanes.freezeRows(1);
        }
        ws.getRangeByIndexes(0, 0, Math.min(rows, AUTOFIT_ROWS), cols).format.autofitColumns();
        await context.sync();
        written.push({ name, address: `A1:${indexToColLetter(cols - 1)}${rows}`, rows });
      }
      if (written.length) {
        context.workbook.worksheets.getItem(written[0].name).activate();
        await context.sync();
      }
      return written;
    });
  } catch (e) {
    if (created.length) {
      try {
        await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          const names = new Set(created);
          for (const ws of sheets.items) if (names.has(ws.name)) ws.delete();
          await context.sync();
        });
      } catch {
        /* 尽力清理写了一半的结果表 */
      }
    }
    throw e;
  }
}

function previewRows(grid: unknown[][], extraLast = false): unknown[][] {
  if (grid.length <= PREVIEW_ROWS + 1) return grid;
  const head = grid.slice(0, PREVIEW_ROWS + 1);
  return extraLast ? [...head, grid[grid.length - 1]] : head;
}

// ---------------- 分组汇总 ----------------

const AGGREGATIONS = new Set<Aggregation>(["sum", "count", "count_rows", "average", "min", "max", "count_distinct"]);
const DATE_PARTS = new Set<DatePart>(["year", "quarter", "month", "week", "day"]);

function metricFormat(metric: MetricColumn, table: LoadedTable): string | undefined {
  if (metric.index == null || metric.agg === "count" || metric.agg === "count_rows" || metric.agg === "count_distinct") {
    return undefined;
  }
  const source = table.formats[metric.index];
  if (metric.agg !== "average") return source || undefined;
  return source && (/[.]/.test(source) || /[ymd]/i.test(source)) ? source : "#,##0.00";
}

export async function toolGroupSummary(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const table = await loadTable(args.sheet, args.range, args.has_headers !== false, ctx);
  if (!table.rows.length) return { ok: false, error: `${table.sheet}!${table.address} 没有数据行` };

  const groups: GroupColumn[] = specList(args.group_by).map((spec, i) => {
    const index = resolveColumn(table, spec, `group_by[${i}]`);
    return { index, header: table.headers[index] };
  });
  if (args.date_column != null && args.date_column !== "") {
    const index = resolveColumn(table, args.date_column, "date_column");
    const part: DatePart = DATE_PARTS.has(args.date_part) ? args.date_part : "month";
    const existing = groups.find((g) => g.index === index);
    if (existing) existing.datePart = part;
    else groups.unshift({ index, header: table.headers[index], datePart: part });
  }
  if (!groups.length) return { ok: false, error: "group_summary 需要 group_by（分组列）或 date_column（按日期分段）" };

  const metrics: MetricColumn[] = (Array.isArray(args.metrics) ? args.metrics : []).map((raw: any, i: number) => {
    const agg = String(raw?.agg || "sum") as Aggregation;
    if (!AGGREGATIONS.has(agg)) throw new Error(`metrics[${i}].agg 不支持「${raw?.agg}」`);
    if (agg === "count_rows") return { index: null, agg, header: String(raw?.header || "").trim() || "行数" };
    const index = resolveColumn(table, raw?.column, `metrics[${i}].column`);
    return { index, agg, header: String(raw?.header || "").trim() || `${table.headers[index]}_${AGG_LABELS[agg]}` };
  });
  if (!metrics.length) metrics.push({ index: null, agg: "count_rows", header: "行数" });

  const sort = ["key", "desc", "asc"].includes(args.sort) ? args.sort : "key";
  const result = groupSummary(table.rows, groups, metrics, {
    sort,
    totalRow: args.total_row !== false,
    parseTextNumbers: args.parse_text_numbers === true,
  });
  const formats = [
    ...groups.map((g) => (g.datePart ? DATE_PART_FORMATS[g.datePart] : table.formats[g.index] || undefined)),
    ...metrics.map((m) => metricFormat(m, table)),
  ];
  checkAbort(ctx);
  const [written] = await writeSheets([{ baseName: String(args.dest_sheet || "").trim() || `${table.sheet}_汇总`, grid: result.grid, formats }], ctx);

  const warnings: string[] = [];
  metrics.forEach((m, i) => {
    if (result.blanks[i] > 0 && m.agg !== "count" && m.agg !== "count_distinct") {
      warnings.push(`「${table.headers[m.index!]}」有 ${result.blanks[i]} 行为空，未计入${AGG_LABELS[m.agg]}`);
    }
    if (result.nonNumeric[i] > 0) {
      warnings.push(
        `「${table.headers[m.index!]}」有 ${result.nonNumeric[i]} 个非空单元格不是数字，未计入${args.parse_text_numbers === true ? "" : "（可先 clean_column 转数字，或设 parse_text_numbers=true）"}`
      );
    }
  });
  if (result.unparsedDates > 0) warnings.push(`有 ${result.unparsedDates} 行的日期无法识别，归入「${BAD_DATE_LABEL}」`);
  if (result.blankGroupRows > 0) warnings.push(`有 ${result.blankGroupRows} 行的分组列为空，归入「${BLANK_LABEL}」`);

  const source = `${table.sheet}!${table.address}`;
  pushOpLog({
    description: `分组汇总 ${source} → 新 Sheet「${written.name}」（${result.groupCount} 组）`,
    kind: "write",
    sheet: written.name,
    address: written.address,
    after: gridPreview(result.grid),
  });
  registerArtifact({
    name: written.name,
    kind: "range",
    sheet: written.name,
    address: written.address,
    brief: `按 ${result.grid[0].slice(0, groups.length).join("、")} 汇总 ${table.sheet} 全表（${result.groupCount} 组）`,
  });
  return {
    ok: true,
    data: {
      dest_sheet: written.name,
      range: written.address,
      source,
      source_rows: result.rowsUsed,
      groups: result.groupCount,
      total_row: result.hasTotalRow,
      preview: previewRows(result.grid, result.hasTotalRow),
      warnings,
      note: "结果是静态值，覆盖全表；源数据改动后需重新汇总。",
    },
    summary: `已按 ${groups.map((g) => g.header).join("、")} 汇总 ${result.rowsUsed} 行 → 「${written.name}」（${result.groupCount} 组）`,
    undoable: false,
  };
}

// ---------------- 两表对账 ----------------

function sideLabel(table: LoadedTable, other: LoadedTable, fallback: string): string {
  if (table.sheet === other.sheet) return fallback;
  return table.sheet.length > 12 ? table.sheet.slice(0, 12) : table.sheet;
}

export async function toolReconcileTables(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const hasHeaders = args.has_headers !== false;
  const left = await loadTable(args.left_sheet, args.left_range, hasHeaders, ctx);
  const right = await loadTable(args.right_sheet, args.right_range, hasHeaders, ctx);
  if (left.sheet === right.sheet && left.address === right.address) {
    return { ok: false, error: "左右两边是同一个区域，请指定两张不同的表或区域" };
  }

  const leftKeys = specList(args.keys).map((spec, i) => resolveColumn(left, spec, `keys[${i}]`));
  if (!leftKeys.length) return { ok: false, error: "reconcile_tables 需要 keys（匹配键列）" };
  const rightKeySpecs = specList(args.right_keys).length ? specList(args.right_keys) : leftKeys.map((i) => left.headers[i]);
  if (rightKeySpecs.length !== leftKeys.length) return { ok: false, error: "right_keys 的列数必须与 keys 一致" };
  const rightKeys = rightKeySpecs.map((spec, i) => resolveColumn(right, spec, `right_keys[${i}]`));

  const leftCompare = specList(args.compare_columns).map((spec, i) => resolveColumn(left, spec, `compare_columns[${i}]`));
  const rightCompareSpecs = specList(args.right_compare_columns).length
    ? specList(args.right_compare_columns)
    : leftCompare.map((i) => left.headers[i]);
  if (rightCompareSpecs.length !== leftCompare.length) {
    return { ok: false, error: "right_compare_columns 的列数必须与 compare_columns 一致" };
  }
  const rightCompare = rightCompareSpecs.map((spec, i) => resolveColumn(right, spec, `right_compare_columns[${i}]`));

  const leftLabel = sideLabel(left, right, "左");
  const rightLabel = sideLabel(right, left, "右");
  const keyHeaders = leftKeys.map((i) => left.headers[i]);
  const compareHeaders = leftCompare.map((i) => left.headers[i]);
  const result = reconcileTables(
    { label: leftLabel, rows: left.rows, keys: leftKeys, compare: leftCompare },
    { label: rightLabel, rows: right.rows, keys: rightKeys, compare: rightCompare },
    {
      keyHeaders,
      compareHeaders,
      tolerance: Number(args.tolerance) || 0,
      keyOptions: { ignoreCase: args.ignore_case === true },
      includeMatched: args.include_matched === true,
    }
  );

  const counts = {
    left_rows: result.leftRows,
    right_rows: result.rightRows,
    matched: result.matched,
    mismatched: result.mismatched,
    only_left: result.onlyLeft,
    only_right: result.onlyRight,
  };
  const warnings: string[] = [];
  if (result.leftBlankKeyRows) warnings.push(`${left.sheet} 有 ${result.leftBlankKeyRows} 行键为空，未参与对账`);
  if (result.rightBlankKeyRows) warnings.push(`${right.sheet} 有 ${result.rightBlankKeyRows} 行键为空，未参与对账`);
  if (result.leftDuplicateKeys) warnings.push(`${left.sheet} 有 ${result.leftDuplicateKeys} 个键重复出现，数值已按键求和`);
  if (result.rightDuplicateKeys) warnings.push(`${right.sheet} 有 ${result.rightDuplicateKeys} 个键重复出现，数值已按键求和`);
  const sides = { left: `${left.sheet}!${left.address}`, right: `${right.sheet}!${right.address}`, left_label: leftLabel, right_label: rightLabel };

  if (result.grid.length === 1) {
    return {
      ok: true,
      data: { ...sides, ...counts, columns: result.columns, warnings, dest_sheet: null },
      summary: `两表按 ${keyHeaders.join("+")} 完全一致（${result.matched} 个键），未新建结果表`,
    };
  }

  const formats = result.columnRoles.map((role) => {
    if (role.role === "key") return left.formats[leftKeys[role.index]] || undefined;
    if (role.role === "left" || role.role === "diff") return left.formats[leftCompare[role.index]] || undefined;
    if (role.role === "right") return right.formats[rightCompare[role.index]] || undefined;
    return undefined;
  });
  checkAbort(ctx);
  const [written] = await writeSheets([{ baseName: String(args.dest_sheet || "").trim() || "对账结果", grid: result.grid, formats }], ctx);
  pushOpLog({
    description: `对账 ${sides.left} ↔ ${sides.right} → 新 Sheet「${written.name}」`,
    kind: "write",
    sheet: written.name,
    address: written.address,
    after: gridPreview(result.grid),
  });
  registerArtifact({
    name: written.name,
    kind: "range",
    sheet: written.name,
    address: written.address,
    brief: `${left.sheet} 与 ${right.sheet} 按 ${keyHeaders.join("+")} 对账：不一致 ${result.mismatched}、仅左 ${result.onlyLeft}、仅右 ${result.onlyRight}`,
  });
  return {
    ok: true,
    data: {
      ...sides,
      ...counts,
      dest_sheet: written.name,
      range: written.address,
      columns: result.columns,
      preview: previewRows(result.grid),
      warnings,
      note: "「状态」列可筛选；差额 = 左 − 右，缺失一侧按 0 计；两表总额见 columns。",
    },
    summary: `对账完成：不一致 ${result.mismatched}、仅在${leftLabel} ${result.onlyLeft}、仅在${rightLabel} ${result.onlyRight} → 「${written.name}」`,
    undoable: false,
  };
}

// ---------------- 多表纵向合并 ----------------

export async function toolStackSheets(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const names = [...new Set((Array.isArray(args.sheets) ? args.sheets : []).map((s: unknown) => String(s ?? "").trim()).filter(Boolean))] as string[];
  if (names.length < 2) return { ok: false, error: "stack_sheets 需要至少两张不同的 Sheet" };
  const hasHeaders = args.has_headers !== false;
  const tables: LoadedTable[] = [];
  let cells = 0;
  for (const name of names) {
    const table = await loadTable(name, undefined, hasHeaders, ctx);
    cells += (table.rows.length + 1) * table.headers.length;
    if (cells > MAX_RESHAPE_CELLS) {
      return { ok: false, error: `合并的表合计超过 ${MAX_RESHAPE_CELLS} 个单元格，请分批合并或用 columns 只保留需要的列` };
    }
    tables.push(table);
  }
  const columns = Array.isArray(args.columns) ? args.columns.map((c: unknown) => String(c ?? "").trim()).filter(Boolean) : [];
  const sourceHeader = args.source_column === undefined ? "来源" : String(args.source_column).trim() || null;
  const result = stackTables(
    tables.map((t) => ({ label: t.sheet, headers: t.headers, rows: t.rows, formats: t.formats })),
    { sourceHeader, columns }
  );
  if (columns.length && result.unknownColumns.length === columns.length) {
    return { ok: false, error: `指定的 columns 在这些表里都找不到：${columns.join("、")}` };
  }
  const dataRows = result.grid.length - 1;
  if (dataRows < 1) return { ok: false, error: "这些表都没有数据行" };
  checkAbort(ctx);
  const [written] = await writeSheets([{ baseName: String(args.dest_sheet || "").trim() || "合并结果", grid: result.grid, formats: result.formats }], ctx);

  const warnings: string[] = [];
  for (const s of result.sources) {
    if (s.missing.length) warnings.push(`「${s.label}」缺少列：${s.missing.join("、")}（已留空）`);
    if (s.dropped.length) warnings.push(`「${s.label}」的列未输出：${s.dropped.join("、")}`);
  }
  if (result.unknownColumns.length) warnings.push(`这些列在所有表中都不存在：${result.unknownColumns.join("、")}`);

  pushOpLog({
    description: `合并 ${names.length} 张表（${names.join("、")}）→ 新 Sheet「${written.name}」（${dataRows} 行）`,
    kind: "write",
    sheet: written.name,
    address: written.address,
    after: gridPreview(result.grid),
  });
  registerArtifact({
    name: written.name,
    kind: "range",
    sheet: written.name,
    address: written.address,
    brief: `${names.join("、")} 纵向合并（${dataRows} 行）`,
  });
  return {
    ok: true,
    data: {
      dest_sheet: written.name,
      range: written.address,
      rows: dataRows,
      columns: result.grid[0],
      sources: result.sources.map((s) => ({ sheet: s.label, rows: s.rows, missing: s.missing, dropped: s.dropped })),
      warnings,
      note: "结果是静态值，不带源表公式。",
    },
    summary: `已把 ${names.length} 张表合并为「${written.name}」（${dataRows} 行，${result.grid[0].length} 列）`,
    undoable: false,
  };
}

// ---------------- 按列拆表 ----------------

export async function toolSplitByColumn(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const hasHeaders = args.has_headers !== false;
  const table = await loadTable(args.sheet, args.range, hasHeaders, ctx);
  if (!table.rows.length) return { ok: false, error: `${table.sheet}!${table.address} 没有数据行` };
  const column = resolveColumn(table, args.column, "column");
  const header = table.headers[column];
  const groups = splitByKey(table.rows, column);
  if (groups.length < 2) {
    return { ok: false, error: `「${header}」只有 ${groups.length} 个取值${groups[0] ? `「${groups[0].label}」` : ""}，不需要拆表` };
  }
  const limit = Math.min(MAX_SPLIT_SHEETS, Math.max(1, Math.trunc(Number(args.max_sheets) || DEFAULT_SPLIT_SHEETS)));
  if (groups.length > limit) {
    return {
      ok: false,
      error: `「${header}」有 ${groups.length} 个不同取值，超过 max_sheets=${limit}（上限 ${MAX_SPLIT_SHEETS}）。前几个取值：${groups
        .slice(0, 8)
        .map((g) => g.label)
        .join("、")}。请确认后调大 max_sheets，或换一列 / 先用 group_summary 看分布。`,
    };
  }
  const prefix = String(args.name_prefix ?? "").trim();
  if (groups.length > CONFIRM_SPLIT_OVER && ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `按「${header}」拆成 ${groups.length} 张新 Sheet`,
      `源表 ${table.sheet}!${table.address} 不会改动。将新建：\n${groups
        .slice(0, 12)
        .map((g) => `${prefix}${g.label}（${g.rows.length} 行）`)
        .join("\n")}${groups.length > 12 ? `\n…共 ${groups.length} 张` : ""}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了拆表。请不要重试，向用户确认拆分列或数量。" };
  }

  const written = await writeSheets(
    groups.map((g) => ({
      baseName: `${prefix}${g.label}`,
      grid: hasHeaders ? [table.headerRow, ...g.rows] : g.rows,
      formats: table.formats,
      header: hasHeaders,
    })),
    ctx
  );
  const sheets = written.map((w, i) => ({ sheet: w.name, value: groups[i].label, rows: groups[i].rows.length }));
  pushOpLog({
    description: `按「${header}」把 ${table.sheet} 拆成 ${sheets.length} 张表：${sheets.map((s) => s.sheet).join("、")}`,
    kind: "structure",
    sheet: written[0].name,
  });
  registerArtifact({
    name: `${table.sheet}按${header}拆分`,
    kind: "sheet",
    sheet: written[0].name,
    brief: `按「${header}」拆出 ${sheets.length} 张表：${sheets.slice(0, 6).map((s) => s.sheet).join("、")}${sheets.length > 6 ? "…" : ""}`,
  });
  return {
    ok: true,
    data: {
      source: `${table.sheet}!${table.address}`,
      column: header,
      sheets,
      total_rows: sheets.reduce((s, x) => s + x.rows, 0),
      note: "写入的是值和数字格式，不带公式；源表未改动。",
    },
    summary: `已按「${header}」拆成 ${sheets.length} 张表（共 ${sheets.reduce((s, x) => s + x.rows, 0)} 行）`,
    undoable: false,
  };
}

// ---------------- 宽表转长表 ----------------

const MAX_SHEET_ROWS = 1_048_576;

export async function toolUnpivotColumns(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const table = await loadTable(args.sheet, args.range, true, ctx);
  if (!table.rows.length) return { ok: false, error: `${table.sheet}!${table.address} 没有数据行` };
  const idSpecs = specList(args.id_columns);
  const valueSpecs = specList(args.value_columns);
  if (!idSpecs.length && !valueSpecs.length) {
    return { ok: false, error: "unpivot_columns 需要 id_columns（保持不动的标识列，如 产品、地区）或 value_columns（要转成行的列，如 1月…12月）" };
  }
  const all = table.headers.map((_, i) => i);
  let idCols = idSpecs.map((spec, i) => resolveColumn(table, spec, `id_columns[${i}]`));
  const valueCols = valueSpecs.length
    ? valueSpecs.map((spec, i) => resolveColumn(table, spec, `value_columns[${i}]`))
    : all.filter((i) => !idCols.includes(i));
  if (!idSpecs.length) idCols = all.filter((i) => !valueCols.includes(i));
  const overlap = valueCols.filter((i) => idCols.includes(i));
  if (overlap.length) {
    return { ok: false, error: `这些列同时出现在 id_columns 和 value_columns 里：${overlap.map((i) => table.headers[i]).join("、")}` };
  }
  if (!valueCols.length) return { ok: false, error: "没有要展开的值列：表里除了 id_columns 已经没有其他列" };

  const result = unpivotRows(
    table.rows,
    idCols,
    valueCols.map((i) => ({ index: i, label: isEmptyHeader(table.headerRow[i]) ? table.headers[i] : table.headerRow[i] })),
    { skipBlanks: args.skip_blanks !== false }
  );
  if (!result.rows.length) return { ok: false, error: `「${valueCols.map((i) => table.headers[i]).join("、")}」都是空的，没有可展开的值` };
  if (result.rows.length + 1 > MAX_SHEET_ROWS) {
    return { ok: false, error: `展开后有 ${result.rows.length} 行，超过 Excel 单表 ${MAX_SHEET_ROWS} 行上限。请用 range 或 value_columns 分批处理` };
  }
  const attributeHeader = String(args.attribute_header ?? "").trim() || "属性";
  const valueHeader = String(args.value_header ?? "").trim() || "值";
  const grid: unknown[][] = [[...idCols.map((i) => table.headers[i]), attributeHeader, valueHeader], ...result.rows];
  const valueFormats = [...new Set(valueCols.map((i) => table.formats[i]))];
  const formats = [...idCols.map((i) => table.formats[i] || undefined), undefined, valueFormats.length === 1 ? valueFormats[0] || undefined : undefined];
  checkAbort(ctx);
  const [written] = await writeSheets([{ baseName: String(args.dest_sheet || "").trim() || `${table.sheet}_长表`, grid, formats }], ctx);

  const source = `${table.sheet}!${table.address}`;
  pushOpLog({
    description: `宽表转长表 ${source} → 新 Sheet「${written.name}」（${result.rows.length} 行）`,
    kind: "write",
    sheet: written.name,
    address: written.address,
    after: gridPreview(grid),
  });
  registerArtifact({
    name: written.name,
    kind: "range",
    sheet: written.name,
    address: written.address,
    brief: `${table.sheet} 的 ${valueCols.length} 列转成「${attributeHeader}/${valueHeader}」两列（${result.rows.length} 行）`,
  });
  return {
    ok: true,
    data: {
      dest_sheet: written.name,
      range: written.address,
      source,
      source_rows: result.sourceRows,
      id_columns: idCols.map((i) => table.headers[i]),
      value_columns: valueCols.map((i) => table.headers[i]),
      rows: result.rows.length,
      skipped_blanks: result.skippedBlanks,
      preview: previewRows(grid),
      note: "结果是静态值；可以直接用 group_summary 或 create_pivot_table 按「属性」列汇总。",
    },
    summary: `已把 ${table.sheet} 的 ${valueCols.length} 列转成长表「${written.name}」（${result.rows.length} 行）`,
    undoable: false,
  };
}

function isEmptyHeader(value: unknown): boolean {
  return value == null || String(value).trim() === "";
}

// ---------------- 两个版本逐格比对 ----------------

const MAX_DIFF_ROWS = 20_000;
const MAX_HIGHLIGHT_MARKS = 3_000;
const HIGHLIGHT_BATCH = 500;
const CHANGED_FILL = "#FFEB9C";
const ADDED_FILL = "#C6EFCE";

function firstRowNumber(address: string): number {
  const m = /^\$?[A-Za-z]+\$?(\d+)/.exec(address);
  return m ? Number(m[1]) : 1;
}

export async function toolCompareSheets(args: any, ctx: ReshapeContext): Promise<ToolResult> {
  const hasHeaders = args.has_headers !== false;
  const oldTable = await loadTable(args.old_sheet, args.old_range, hasHeaders, ctx);
  const newTable = await loadTable(args.new_sheet, args.new_range, hasHeaders, ctx);
  if (oldTable.sheet === newTable.sheet && oldTable.address === newTable.address) {
    return { ok: false, error: "新旧两边是同一个区域，请指定两张不同的表或区域" };
  }
  const oldKeys = specList(args.keys).map((spec, i) => resolveColumn(oldTable, spec, `keys[${i}]`));
  const newKeySpecs = specList(args.new_keys).length ? specList(args.new_keys) : oldKeys.map((i) => oldTable.headers[i]);
  if (newKeySpecs.length !== oldKeys.length) return { ok: false, error: "new_keys 的列数必须与 keys 一致" };
  const newKeys = newKeySpecs.map((spec, i) => resolveColumn(newTable, spec, `new_keys[${i}]`));
  const headerRows = hasHeaders ? 1 : 0;
  const side = (table: LoadedTable, keys: number[]) => ({
    headers: table.headers,
    rows: table.rows,
    firstRow: firstRowNumber(table.address) + headerRows,
    firstColumn: table.columnIndex,
    keys,
  });
  const diff = diffTables(side(oldTable, oldKeys), side(newTable, newKeys), {
    matchByHeader: hasHeaders,
    tolerance: Number(args.tolerance) || 0,
    ignoreCase: args.ignore_case === true,
    maxChanges: MAX_DIFF_ROWS,
  });
  const sides = { old: `${oldTable.sheet}!${oldTable.address}`, new: `${newTable.sheet}!${newTable.address}` };
  const counts = {
    changed_cells: diff.changedCells,
    changed_rows: diff.changedRows,
    added_rows: diff.addedRows,
    removed_rows: diff.removedRows,
    unchanged_rows: diff.unchangedRows,
  };
  const warnings: string[] = [];
  if (diff.onlyOldColumns.length) warnings.push(`新版本没有这些列：${diff.onlyOldColumns.join("、")}`);
  if (diff.onlyNewColumns.length) warnings.push(`新版本多了这些列：${diff.onlyNewColumns.join("、")}`);
  if (diff.duplicateKeys) warnings.push(`有 ${diff.duplicateKeys} 个键重复出现，已按出现顺序一一配对`);
  if (diff.blankKeyRows) warnings.push(`有 ${diff.blankKeyRows} 行键为空，未参与比对`);
  if (diff.omitted) warnings.push(`差异太多，清单只列出前 ${MAX_DIFF_ROWS} 条，另有 ${diff.omitted} 条未列出`);
  if (!oldKeys.length && (diff.addedRows || diff.removedRows) && diff.changedRows > 5) {
    warnings.push("按行位置比对：中间插入或删除行会让后面每行都显示为修改，有唯一键（如工号、订单号）时请用 keys 重新比对");
  }

  const total = diff.changes.length + diff.omitted + diff.onlyOldColumns.length + diff.onlyNewColumns.length;
  if (total === 0) {
    return {
      ok: true,
      data: { ...sides, ...counts, compared_columns: diff.comparedColumns, warnings, dest_sheet: null },
      summary: `两个版本内容一致（比对 ${diff.unchangedRows} 行 × ${diff.comparedColumns.length} 列），未新建结果表`,
    };
  }

  const rowHeader = oldKeys.length ? oldKeys.map((i) => oldTable.headers[i]).join("+") : "行";
  const grid: unknown[][] = [
    ["变化", rowHeader, "列", "原值", "新值", "原位置", "新位置"],
    ...diff.onlyOldColumns.map((h) => ["删除列", "", h, "", "", "", ""]),
    ...diff.onlyNewColumns.map((h) => ["新增列", "", h, "", "", "", ""]),
    ...diff.changes.map((c) => [c.kind, c.rowLabel, c.column, c.before, c.after, c.oldAddress, c.newAddress]),
  ];
  checkAbort(ctx);
  const [written] = await writeSheets([{ baseName: String(args.dest_sheet || "").trim() || "版本差异", grid, formats: [] }], ctx);
  pushOpLog({
    description: `比对 ${sides.old} → ${sides.new}，差异清单写到「${written.name}」`,
    kind: "write",
    sheet: written.name,
    address: written.address,
    after: gridPreview(grid),
  });
  registerArtifact({
    name: written.name,
    kind: "range",
    sheet: written.name,
    address: written.address,
    brief: `${oldTable.sheet} → ${newTable.sheet}：改 ${diff.changedCells} 格、增 ${diff.addedRows} 行、删 ${diff.removedRows} 行`,
  });

  let highlighted = 0;
  if (args.highlight === true) {
    const marks = diff.changes.filter((c) => c.mark);
    if (marks.length > MAX_HIGHLIGHT_MARKS) {
      warnings.push(`差异有 ${marks.length} 处，超过 ${MAX_HIGHLIGHT_MARKS} 处未在新版本上标色，请按差异清单筛选查看`);
    } else {
      highlighted = await highlightChanges(newTable.sheet, marks, ctx);
      pushOpLog({ description: `在「${newTable.sheet}」标出 ${highlighted} 处差异（黄=修改，绿=新增行）`, kind: "format", sheet: newTable.sheet });
    }
  }
  return {
    ok: true,
    data: {
      ...sides,
      ...counts,
      dest_sheet: written.name,
      range: written.address,
      compared_columns: diff.comparedColumns,
      highlighted,
      preview: previewRows(grid),
      warnings,
      note: "「变化」列可筛选；原位置/新位置是各自工作表里的单元格或行号。",
    },
    summary: `比对完成：改动 ${diff.changedCells} 格、新增 ${diff.addedRows} 行、删除 ${diff.removedRows} 行 → 「${written.name}」`,
    undoable: false,
  };
}

async function highlightChanges(sheetName: string, marks: { kind: string; mark?: { row: number; column: number; columns: number } }[], ctx: ReshapeContext): Promise<number> {
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(sheetName);
    let n = 0;
    for (const { kind, mark } of marks) {
      if (!mark) continue;
      ws.getRangeByIndexes(mark.row, mark.column, 1, mark.columns).format.fill.color = kind === "新增行" ? ADDED_FILL : CHANGED_FILL;
      n++;
      if (n % HIGHLIGHT_BATCH === 0) {
        checkAbort(ctx);
        await context.sync();
      }
    }
    await context.sync();
    return n;
  });
}
