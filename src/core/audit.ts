/*
 * 数据体检引擎：纯函数，对二维网格做缺失值 / 重复行 / 类型不一致 /
 * 异常值 / 首尾空格五类检测，输出结构化问题清单（供报告 Sheet 与
 * 可勾选修复卡片共用）。
 */
import { AuditIssue } from "./types";

/** 一行数据：row 为工作表内的绝对行号（1 起），cells 为该行的值 */
export interface AuditRowInput {
  row: number;
  cells: any[];
}

export interface AuditStats {
  /** 参与分析的数据行数（不含表头） */
  rows: number;
  columns: number;
  emptyCells: number;
  totalCells: number;
  /** 实际分析行数小于总行数时为真（大表抽样） */
  sampled: boolean;
  sampledRows?: number;
  totalDataRows?: number;
}

export interface AuditResult {
  stats: AuditStats;
  issues: AuditIssue[];
}

const MAX_DETAIL_CELLS = 300;
const MAX_DUP_GROUPS = 50;
const MIN_OUTLIER_VALUES = 8;

function colLetter(index: number, colStart = 0): string {
  let s = "";
  let n = colStart + index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function isEmpty(v: unknown): boolean {
  return v === "" || v === null || v === undefined;
}

type CellKind = "number" | "numericText" | "dateText" | "text";

/** 判断非空单元格的内容类别（numericText = 存成文本的数字，是类型问题的典型来源） */
function cellKind(v: unknown): CellKind {
  if (typeof v === "number") return "number";
  const s = String(v).trim();
  if (!s) return "text";
  if (/^-?[\d,]+(\.\d+)?%?$/.test(s) || /^-?\.\d+$/.test(s)) return "numericText";
  if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(s) || /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(s)) return "dateText";
  return "text";
}

function previewRow(cells: any[], maxCols = 8): string {
  return cells
    .slice(0, maxCols)
    .map((v) => (isEmpty(v) ? "·" : String(v).slice(0, 14)))
    .join(" | ");
}

function ellipsize(list: string[], shown: number): string {
  const head = list.slice(0, shown).join("、");
  return list.length > shown ? `${head} …（共 ${list.length} 处）` : head;
}

function quantile(sorted: number[], q: number): number {
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/** 对表数据做体检。headers 为列名；rows 为数据行（带绝对行号） */
export function auditData(
  headers: string[],
  rows: AuditRowInput[],
  opts?: { sampledRows?: number; totalDataRows?: number; colStart?: number }
): AuditResult {
  const dataRows = rows.length;
  const cols = Math.max(headers.length, ...rows.map((r) => r.cells.length), 0);
  const colStart = Math.max(0, Math.trunc(opts?.colStart ?? 0));
  const letter = (c: number) => colLetter(c, colStart);
  const headerNames = Array.from({ length: cols }, (_, c) => {
    const h = String(headers[c] ?? "").trim();
    return h || `第${letter(c)}列`;
  });

  let emptyCells = 0;
  let totalCells = 0;
  for (const r of rows) {
    for (let c = 0; c < cols; c++) {
      totalCells++;
      if (isEmpty(r.cells[c])) emptyCells++;
    }
  }

  const issues: AuditIssue[] = [];

  // ---------- 1. 每列缺失值 ----------
  for (let c = 0; c < cols; c++) {
    const emptyRows: number[] = [];
    for (const r of rows) {
      if (isEmpty(r.cells[c])) emptyRows.push(r.row);
    }
    if (emptyRows.length === 0) continue;
    const ratio = emptyRows.length / Math.max(1, dataRows);
    const header = headerNames[c];
    if (emptyRows.length === dataRows) {
      issues.push({
        id: `missing:${letter(c)}:all`,
        type: "missing",
        severity: "medium",
        column: header,
        count: emptyRows.length,
        sample: `${header} 列全部为空`,
        suggestion: `「${header}」列没有任何数据，确认是漏填了一整列还是多余的空列（多余可删除该列）。`,
        detail: { column: letter(c), header, empty_rows: emptyRows.slice(0, MAX_DETAIL_CELLS) },
      });
      continue;
    }
    const severity = ratio >= 0.3 ? "high" : ratio >= 0.05 ? "medium" : "low";
    const samples = emptyRows.slice(0, 5).map((r) => `${letter(c)}${r}`);
    issues.push({
      id: `missing:${letter(c)}`,
      type: "missing",
      severity,
      column: header,
      count: emptyRows.length,
      sample: `${ellipsize(samples, 3)} 为空`,
      suggestion: `「${header}」列有 ${emptyRows.length} 个空值（占 ${Math.round(ratio * 100)}%）。逐个核对是漏填还是本就没有值；可补齐、填默认值（如 0 / 未知）或保留并在分析时排除。`,
      detail: { column: letter(c), header, empty_rows: emptyRows.slice(0, MAX_DETAIL_CELLS) },
    });
  }

  // ---------- 2. 完全重复的行 ----------
  if (dataRows >= 2) {
    const groups = new Map<string, number[]>();
    for (const r of rows) {
      const key = JSON.stringify(r.cells.slice(0, cols));
      const list = groups.get(key);
      if (list) list.push(r.row);
      else groups.set(key, [r.row]);
    }
    const dupGroups = [...groups.values()].filter((g) => g.length > 1);
    if (dupGroups.length > 0) {
      const extraRows = dupGroups.reduce((s, g) => s + g.length - 1, 0);
      const first = rows.find((r) => r.row === dupGroups[0][0]);
      issues.push({
        id: "duplicate",
        type: "duplicate",
        severity: extraRows / dataRows >= 0.05 ? "high" : "medium",
        count: extraRows,
        sample: `${dupGroups.length} 组重复（如第 ${dupGroups[0].join("、")} 行）：${first ? previewRow(first.cells) : ""}`,
        suggestion: `存在 ${dupGroups.length} 组内容完全相同的行，多余 ${extraRows} 行。通常保留每组第一条、删除其余；删除前确认不是业务上允许的重复（如多次下单）。`,
        detail: {
          groups: dupGroups.slice(0, MAX_DUP_GROUPS).map((g) => ({
            keep_row: g[0],
            delete_rows: g.slice(1, MAX_DETAIL_CELLS),
            preview: previewRow((rows.find((r) => r.row === g[0])?.cells ?? []).slice(0, cols)),
          })),
        },
      });
    }
  }

  // ---------- 3. 列内类型不一致 ----------
  for (let c = 0; c < cols; c++) {
    const header = headerNames[c];
    const kinds: { kind: CellKind; row: number; value: unknown }[] = [];
    for (const r of rows) {
      const v = r.cells[c];
      if (!isEmpty(v)) kinds.push({ kind: cellKind(v), row: r.row, value: v });
    }
    if (kinds.length < 3) continue;
    const counts: Record<CellKind, number> = { number: 0, numericText: 0, dateText: 0, text: 0 };
    for (const k of kinds) counts[k.kind]++;
    const nonEmpty = kinds.length;
    const strCells = (pred?: (k: CellKind) => boolean) =>
      kinds.filter((k) => k.kind !== "number" && (!pred || pred(k.kind)));

    // 数字列里混入文本（含"待定"这类非数值、千分位/百分号文本）
    if (counts.number >= nonEmpty * 0.6) {
      const bad = strCells();
      if (bad.length > 0) {
        const samples = bad.slice(0, 3).map((b) => `${letter(c)}${b.row}="${String(b.value).slice(0, 10)}"`);
        issues.push({
          id: `type:${letter(c)}:mixed`,
          type: "type",
          severity: bad.length / nonEmpty >= 0.1 ? "high" : "medium",
          column: header,
          count: bad.length,
          sample: ellipsize(samples, 2),
          suggestion: `「${header}」列以数字为主，但混入 ${bad.length} 个文本值（千分位/百分号/全角数字/文字）。统一转成数字（去掉 , % 和空格）后再计算；确属文本的（如"待定"）单独核对。`,
          detail: {
            column: letter(c),
            header,
            mode: "text_in_number",
            cells: bad.slice(0, MAX_DETAIL_CELLS).map((b) => ({ row: b.row, value: b.value })),
          },
        });
      }
    }
    // 数字全以文本形式存储（SUM 等公式会算错）
    else if (counts.numericText >= nonEmpty * 0.9 && counts.numericText >= 3 && counts.number === 0) {
      issues.push({
        id: `type:${letter(c)}:astext`,
        type: "type",
        severity: "medium",
        column: header,
        count: counts.numericText,
        sample: `${header} 列的数字全部以文本存储`,
        suggestion: `「${header}」列的数字被存成了文本，SUM/AVERAGE 等公式会忽略它们。整列转换为数值（可用"分列"或乘 1 的方式）。`,
        detail: {
          column: letter(c),
          header,
          mode: "numbers_as_text",
          cells: kinds.filter((k) => k.kind === "numericText").slice(0, MAX_DETAIL_CELLS).map((b) => ({ row: b.row, value: b.value })),
        },
      });
    }
    // 疑似日期列混入纯数字
    else if (counts.dateText >= nonEmpty * 0.6 && counts.number + counts.numericText > 0) {
      const bad = kinds.filter((k) => k.kind === "number" || k.kind === "numericText");
      const samples = bad.slice(0, 3).map((b) => `${letter(c)}${b.row}=${String(b.value).slice(0, 10)}`);
      issues.push({
        id: `type:${letter(c)}:datemixed`,
        type: "type",
        severity: "medium",
        column: header,
        count: bad.length,
        sample: ellipsize(samples, 2),
        suggestion: `「${header}」列看起来是日期，但混入 ${bad.length} 个纯数字。可能是日期被写成序列号或年份，逐一核对并统一为日期格式。`,
        detail: {
          column: letter(c),
          header,
          mode: "number_in_date",
          cells: bad.slice(0, MAX_DETAIL_CELLS).map((b) => ({ row: b.row, value: b.value })),
        },
      });
    }
  }

  // ---------- 4. 数值列离群值（IQR） ----------
  for (let c = 0; c < cols; c++) {
    const nums: { row: number; value: number }[] = [];
    for (const r of rows) {
      const v = r.cells[c];
      if (typeof v === "number" && Number.isFinite(v)) nums.push({ row: r.row, value: v });
    }
    if (nums.length < MIN_OUTLIER_VALUES) continue;
    const nonEmpty = rows.filter((r) => !isEmpty(r.cells[c])).length;
    if (nums.length < nonEmpty * 0.8) continue; // 不是足够"纯"的数值列
    const sorted = nums.map((n) => n.value).sort((a, b) => a - b);
    const q1 = quantile(sorted, 0.25);
    const q3 = quantile(sorted, 0.75);
    const iqr = q3 - q1;
    if (iqr <= 0) continue;
    const lo = q1 - 1.5 * iqr;
    const hi = q3 + 1.5 * iqr;
    const out = nums.filter((n) => n.value < lo || n.value > hi);
    if (out.length === 0) continue;
    const header = headerNames[c];
    const samples = out.slice(0, 3).map((o) => `${letter(c)}${o.row}=${o.value}`);
    issues.push({
      id: `outlier:${letter(c)}`,
      type: "outlier",
      severity: out.length / nums.length >= 0.05 ? "medium" : "low",
      column: header,
      count: out.length,
      sample: ellipsize(samples, 2),
      suggestion: `「${header}」列有 ${out.length} 个离群值（超出四分位距正常范围 ${Math.round(lo * 100) / 100} ~ ${Math.round(hi * 100) / 100}）。核对是录入错误（多打/少打零、单位不符）还是真实极值；错误的改正，真实的保留。`,
      detail: {
        column: letter(c),
        header,
        bounds: [Math.round(lo * 100) / 100, Math.round(hi * 100) / 100],
        cells: out.slice(0, MAX_DETAIL_CELLS).map((o) => ({ row: o.row, value: o.value })),
      },
    });
  }

  // ---------- 5. 首尾空格 / 全角空格 ----------
  {
    const bad: { row: number; column: number; header: string; value: string }[] = [];
    for (const r of rows) {
      for (let c = 0; c < cols; c++) {
        const v = r.cells[c];
        if (typeof v === "string" && v !== v.trim()) {
          bad.push({ row: r.row, column: c, header: headerNames[c], value: v.slice(0, 20) });
        }
      }
    }
    if (bad.length > 0) {
      const samples = bad.slice(0, 3).map((b) => `${letter(b.column)}${b.row}="${b.value}…"`);
      issues.push({
        id: "whitespace",
        type: "whitespace",
        severity: "low",
        count: bad.length,
        sample: ellipsize(samples, 2),
        suggestion: `${bad.length} 个单元格含首尾空格（含全角空格），会导致 VLOOKUP/筛选匹配不到。批量去除首尾空格即可。`,
        detail: { cells: bad.slice(0, MAX_DETAIL_CELLS) },
      });
    }
  }

  const stats: AuditStats = {
    rows: dataRows,
    columns: cols,
    emptyCells,
    totalCells,
    sampled: !!opts?.sampledRows && opts.sampledRows < (opts?.totalDataRows ?? opts.sampledRows),
    ...(opts?.sampledRows != null ? { sampledRows: opts.sampledRows } : {}),
    ...(opts?.totalDataRows != null ? { totalDataRows: opts.totalDataRows } : {}),
  };
  return { stats, issues };
}
