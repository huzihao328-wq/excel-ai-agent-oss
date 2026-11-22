/**
 * 整表分组汇总、对账、叠表、拆表、逆透视、版本比对的纯计算：输入输出都是二维值网格，
 * 由 reshapeTools.ts 负责从 Excel 分块读全表、写回新 Sheet。
 */
import { excelSerialToIso, parseLooseDate, parseLooseNumber, toHalfWidth } from "./cleanValues";
import { isEmptyCell } from "./safeValues";
import { indexToColLetter } from "./formulaRefs";

export const BLANK_LABEL = "(空白)";
export const BAD_DATE_LABEL = "(不是日期)";

export type DatePart = "year" | "quarter" | "month" | "week" | "day";
export type Aggregation = "sum" | "count" | "count_rows" | "average" | "min" | "max" | "count_distinct";

export const AGG_LABELS: Record<Aggregation, string> = {
  sum: "合计",
  count: "计数",
  count_rows: "行数",
  average: "平均",
  min: "最小",
  max: "最大",
  count_distinct: "去重计数",
};

export const DATE_PART_LABELS: Record<DatePart, string> = {
  year: "年",
  quarter: "季度",
  month: "月",
  week: "周",
  day: "日",
};

/** 分段后的单元格格式；季度和周写成文本 */
export const DATE_PART_FORMATS: Record<DatePart, string | undefined> = {
  year: "0",
  quarter: undefined,
  month: "yyyy-mm",
  week: undefined,
  day: "yyyy-mm-dd",
};

export interface KeyOptions {
  /** 去首尾空格并把连续空白压成一个，默认 true */
  trim?: boolean;
  /** 全角字母、数字、符号转半角，默认 true */
  halfWidth?: boolean;
  /** 忽略大小写，默认 false */
  ignoreCase?: boolean;
}

export { toHalfWidth };

export function normalizeKey(value: unknown, options: KeyOptions = {}): string {
  if (isEmptyCell(value)) return "";
  let text = String(value);
  if (options.halfWidth !== false) text = toHalfWidth(text);
  if (options.trim !== false) text = text.replace(/\s+/g, " ").trim();
  if (options.ignoreCase) text = text.toLowerCase();
  return text;
}

function displayCell(value: unknown): unknown {
  return typeof value === "string" ? value.replace(/\s+/g, " ").trim() : value;
}

function isBlankRow(row: unknown[]): boolean {
  return row.every(isEmptyCell);
}

/** 去掉浮点累加噪声（0.1+0.2），不改变正常精度的金额 */
export function tidyNumber(n: number): number {
  return Math.abs(n) < 1e15 ? Number(n.toFixed(10)) : n;
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

const DAY_MS = 86400000;
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

function excelSerial(year: number, month: number, day: number): number {
  return (Date.UTC(year, month - 1, day) - EXCEL_EPOCH) / DAY_MS;
}

function isoWeek(year: number, month: number, day: number): { year: number; week: number } {
  const target = new Date(Date.UTC(year, month - 1, day));
  const weekday = (target.getUTCDay() + 6) % 7;
  target.setUTCDate(target.getUTCDate() - weekday + 3);
  const weekYear = target.getUTCFullYear();
  const firstThursday = new Date(Date.UTC(weekYear, 0, 4));
  const week =
    1 + Math.round(((target.getTime() - firstThursday.getTime()) / DAY_MS - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
  return { year: weekYear, week };
}

export interface DateBucket {
  /** 分组与排序用的键，如 2024-03 */
  key: string;
  /** 写回单元格的值：年为数字，月/日为 Excel 日期序列，季度/周为文本 */
  value: string | number;
}

/** 日期单元格（Excel 序列或常见日期文本）归到年/季度/月/周/日 */
export function dateBucket(value: unknown, part: DatePart): DateBucket | null {
  const iso = typeof value === "number" ? excelSerialToIso(value) : parseLooseDate(value);
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  switch (part) {
    case "year":
      return { key: String(y), value: y };
    case "quarter": {
      const label = `${y}-Q${Math.ceil(m / 3)}`;
      return { key: label, value: label };
    }
    case "month":
      return { key: `${y}-${pad2(m)}`, value: excelSerial(y, m, 1) };
    case "week": {
      const w = isoWeek(y, m, d);
      const label = `${w.year}-W${pad2(w.week)}`;
      return { key: label, value: label };
    }
    default:
      return { key: iso, value: excelSerial(y, m, d) };
  }
}

/** 数字在前、文本按中文自然顺序、空值最后 */
export function compareCells(a: unknown, b: unknown): number {
  const aEmpty = a == null || a === "";
  const bEmpty = b == null || b === "";
  if (aEmpty || bEmpty) return aEmpty === bEmpty ? 0 : aEmpty ? 1 : -1;
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "number") return -1;
  if (typeof b === "number") return 1;
  return String(a).localeCompare(String(b), "zh-Hans-CN", { numeric: true });
}

// ---------------- 分组汇总 ----------------

export interface GroupColumn {
  index: number;
  header: string;
  datePart?: DatePart;
}

export interface MetricColumn {
  /** count_rows 不需要列 */
  index: number | null;
  agg: Aggregation;
  header: string;
}

export interface GroupSummaryOptions {
  sort?: "key" | "desc" | "asc";
  totalRow?: boolean;
  parseTextNumbers?: boolean;
  keyOptions?: KeyOptions;
}

export interface GroupSummaryResult {
  grid: unknown[][];
  groupCount: number;
  rowsUsed: number;
  blankRowsSkipped: number;
  /** 分组列为空、归入「(空白)」的行数 */
  blankGroupRows: number;
  /** 日期分段列里无法识别为日期的行数 */
  unparsedDates: number;
  /** 每个指标里非空但不是数字、没有计入的单元格数 */
  nonNumeric: number[];
  /** 每个指标列为空的行数（count_rows 恒为 0） */
  blanks: number[];
  hasTotalRow: boolean;
}

interface Accumulator {
  rows: number;
  blanks: number;
  nonEmpty: number;
  numbers: number;
  sum: number;
  min: number;
  max: number;
  nonNumeric: number;
  distinct: Set<string>;
}

function newAccumulator(): Accumulator {
  return { rows: 0, blanks: 0, nonEmpty: 0, numbers: 0, sum: 0, min: Infinity, max: -Infinity, nonNumeric: 0, distinct: new Set() };
}

function metricNumber(value: unknown, parseText: boolean): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (parseText && typeof value === "string") return parseLooseNumber(value);
  return null;
}

function accumulate(acc: Accumulator, metric: MetricColumn, row: unknown[], options: GroupSummaryOptions): void {
  acc.rows++;
  if (metric.index == null) return;
  const value = row[metric.index];
  if (isEmptyCell(value)) {
    acc.blanks++;
    return;
  }
  acc.nonEmpty++;
  if (metric.agg === "count_distinct") {
    acc.distinct.add(normalizeKey(value, options.keyOptions));
    return;
  }
  if (metric.agg === "count") return;
  const n = metricNumber(value, options.parseTextNumbers === true);
  if (n == null) {
    acc.nonNumeric++;
    return;
  }
  acc.numbers++;
  acc.sum += n;
  if (n < acc.min) acc.min = n;
  if (n > acc.max) acc.max = n;
}

function finishMetric(acc: Accumulator, agg: Aggregation): number | "" {
  switch (agg) {
    case "sum":
      return tidyNumber(acc.sum);
    case "count":
      return acc.nonEmpty;
    case "count_rows":
      return acc.rows;
    case "average":
      return acc.numbers ? tidyNumber(acc.sum / acc.numbers) : "";
    case "min":
      return acc.numbers ? acc.min : "";
    case "max":
      return acc.numbers ? acc.max : "";
    default:
      return acc.distinct.size;
  }
}

interface GroupEntry {
  cells: unknown[];
  sortKeys: (unknown | null)[];
  accs: Accumulator[];
}

function compareGroupKeys(a: GroupEntry, b: GroupEntry): number {
  for (let i = 0; i < a.sortKeys.length; i++) {
    const diff = compareCells(a.sortKeys[i], b.sortKeys[i]);
    if (diff !== 0) return diff;
  }
  return 0;
}

export function groupSummary(
  rows: unknown[][],
  groups: GroupColumn[],
  metrics: MetricColumn[],
  options: GroupSummaryOptions = {}
): GroupSummaryResult {
  if (!groups.length) throw new Error("至少需要一个分组列");
  if (!metrics.length) throw new Error("至少需要一个指标");
  const map = new Map<string, GroupEntry>();
  const totals = metrics.map(() => newAccumulator());
  let rowsUsed = 0;
  let blankRowsSkipped = 0;
  let blankGroupRows = 0;
  let unparsedDates = 0;

  for (const row of rows) {
    if (isBlankRow(row)) {
      blankRowsSkipped++;
      continue;
    }
    rowsUsed++;
    const cells: unknown[] = [];
    const sortKeys: (unknown | null)[] = [];
    const keyParts: string[] = [];
    let blankGroup = false;
    for (const group of groups) {
      const raw = row[group.index];
      if (isEmptyCell(raw)) {
        blankGroup = true;
        cells.push(BLANK_LABEL);
        sortKeys.push(null);
        keyParts.push("\u0000blank");
        continue;
      }
      if (group.datePart) {
        const bucket = dateBucket(raw, group.datePart);
        if (!bucket) {
          unparsedDates++;
          cells.push(BAD_DATE_LABEL);
          sortKeys.push(null);
          keyParts.push("\u0000bad-date");
          continue;
        }
        cells.push(bucket.value);
        sortKeys.push(bucket.key);
        keyParts.push(bucket.key);
        continue;
      }
      cells.push(displayCell(raw));
      sortKeys.push(typeof raw === "number" ? raw : normalizeKey(raw, options.keyOptions));
      keyParts.push(normalizeKey(raw, options.keyOptions));
    }
    if (blankGroup) blankGroupRows++;
    const key = keyParts.join("\u0001");
    let entry = map.get(key);
    if (!entry) {
      entry = { cells, sortKeys, accs: metrics.map(() => newAccumulator()) };
      map.set(key, entry);
    }
    metrics.forEach((metric, i) => {
      accumulate(entry!.accs[i], metric, row, options);
      accumulate(totals[i], metric, row, options);
    });
  }

  const entries = [...map.values()];
  const sort = options.sort || "key";
  if (sort === "key") {
    entries.sort(compareGroupKeys);
  } else {
    const direction = sort === "asc" ? 1 : -1;
    entries.sort((a, b) => {
      const av = finishMetric(a.accs[0], metrics[0].agg);
      const bv = finishMetric(b.accs[0], metrics[0].agg);
      if (av === "" || bv === "") return av === bv ? compareGroupKeys(a, b) : av === "" ? 1 : -1;
      return (av - bv) * direction || compareGroupKeys(a, b);
    });
  }

  const header = [
    ...groups.map((g) => (g.datePart ? `${g.header}(${DATE_PART_LABELS[g.datePart]})` : g.header)),
    ...metrics.map((m) => m.header),
  ];
  const grid: unknown[][] = [header];
  for (const entry of entries) {
    grid.push([...entry.cells, ...entry.accs.map((acc, i) => finishMetric(acc, metrics[i].agg))]);
  }
  const hasTotalRow = options.totalRow !== false && entries.length > 1;
  if (hasTotalRow) {
    const label = groups.map((_, i) => (i === 0 ? "合计" : ""));
    grid.push([...label, ...totals.map((acc, i) => finishMetric(acc, metrics[i].agg))]);
  }

  return {
    grid,
    groupCount: entries.length,
    rowsUsed,
    blankRowsSkipped,
    blankGroupRows,
    unparsedDates,
    nonNumeric: totals.map((acc) => acc.nonNumeric),
    blanks: totals.map((acc) => acc.blanks),
    hasTotalRow,
  };
}

// ---------------- 两表对账 ----------------

export type ReconcileStatus = "不一致" | "仅左表" | "仅右表" | "一致";

export interface ReconcileSide {
  label: string;
  rows: unknown[][];
  keys: number[];
  compare: number[];
}

export interface ReconcileOptions {
  keyHeaders: string[];
  compareHeaders: string[];
  /** 数值允许的绝对误差 */
  tolerance?: number;
  keyOptions?: KeyOptions;
  includeMatched?: boolean;
}

type Aggregated = { kind: "empty" } | { kind: "number"; value: number } | { kind: "text"; value: string };

interface SideEntry {
  keyCells: unknown[];
  count: number;
  values: unknown[][];
}

interface CollectedSide {
  map: Map<string, SideEntry>;
  rows: number;
  blankKeyRows: number;
  duplicateKeys: number;
}

function collectSide(side: ReconcileSide, keyOptions: KeyOptions | undefined): CollectedSide {
  const map = new Map<string, SideEntry>();
  let rows = 0;
  let blankKeyRows = 0;
  for (const row of side.rows) {
    if (isBlankRow(row)) continue;
    rows++;
    const parts = side.keys.map((i) => normalizeKey(row[i], keyOptions));
    if (parts.every((p) => p === "")) {
      blankKeyRows++;
      continue;
    }
    const key = parts.join("\u0001");
    let entry = map.get(key);
    if (!entry) {
      entry = { keyCells: side.keys.map((i) => displayCell(row[i])), count: 0, values: side.compare.map(() => []) };
      map.set(key, entry);
    }
    entry.count++;
    side.compare.forEach((col, j) => entry!.values[j].push(row[col]));
  }
  let duplicateKeys = 0;
  for (const entry of map.values()) if (entry.count > 1) duplicateKeys++;
  return { map, rows, blankKeyRows, duplicateKeys };
}

/** 重复键：全是数字（含「1,200」这类文本数字）时求和，否则去重后用「；」连接 */
function aggregateValues(values: unknown[]): Aggregated {
  const present = values.filter((v) => !isEmptyCell(v));
  if (!present.length) return { kind: "empty" };
  const numbers = present.map(parseLooseNumber);
  if (numbers.every((n) => n != null)) {
    return { kind: "number", value: tidyNumber((numbers as number[]).reduce((s, n) => s + n, 0)) };
  }
  const texts = [...new Set(present.map((v) => String(displayCell(v))))];
  return { kind: "text", value: texts.join("；") };
}

function aggregatedCell(value: Aggregated): unknown {
  return value.kind === "empty" ? "" : value.value;
}

function valuesMatch(left: Aggregated, right: Aggregated, tolerance: number, keyOptions: KeyOptions | undefined): boolean {
  if (left.kind === "empty" || right.kind === "empty") return left.kind === right.kind;
  if (left.kind === "number" && right.kind === "number") {
    return Math.abs(left.value - right.value) <= tolerance;
  }
  return normalizeKey(left.value, keyOptions) === normalizeKey(right.value, keyOptions);
}

export interface ReconcileColumnStats {
  header: string;
  numeric: boolean;
  mismatched: number;
  left_total?: number;
  right_total?: number;
  difference?: number;
}

export interface ReconcileResult {
  grid: unknown[][];
  matched: number;
  mismatched: number;
  onlyLeft: number;
  onlyRight: number;
  leftRows: number;
  rightRows: number;
  leftBlankKeyRows: number;
  rightBlankKeyRows: number;
  leftDuplicateKeys: number;
  rightDuplicateKeys: number;
  columns: ReconcileColumnStats[];
  /** 输出列对应的来源，供调用方套数字格式 */
  columnRoles: ({ role: "status" } | { role: "key"; index: number } | { role: "left" | "right" | "diff"; index: number } | { role: "count" })[];
}

export function reconcileTables(left: ReconcileSide, right: ReconcileSide, options: ReconcileOptions): ReconcileResult {
  if (left.keys.length !== right.keys.length || !left.keys.length) throw new Error("左右两表的键列数量必须一致且至少一列");
  if (left.compare.length !== right.compare.length) throw new Error("左右两表的比对列数量必须一致");
  const tolerance = Math.max(0, Number(options.tolerance) || 0);
  const keyOptions = options.keyOptions;
  const L = collectSide(left, keyOptions);
  const R = collectSide(right, keyOptions);
  const compareCount = left.compare.length;

  type Row = { status: ReconcileStatus; keyCells: unknown[]; l: Aggregated[]; r: Aggregated[]; diffs: string[]; lc: number; rc: number };
  const out: Row[] = [];
  const empty: Aggregated[] = Array.from({ length: compareCount }, () => ({ kind: "empty" }));

  for (const [key, entry] of L.map) {
    const lValues = entry.values.map(aggregateValues);
    const other = R.map.get(key);
    if (!other) {
      out.push({ status: "仅左表", keyCells: entry.keyCells, l: lValues, r: empty, diffs: [], lc: entry.count, rc: 0 });
      continue;
    }
    const rValues = other.values.map(aggregateValues);
    const diffs = options.compareHeaders.filter((_, j) => !valuesMatch(lValues[j], rValues[j], tolerance, keyOptions));
    out.push({
      status: diffs.length ? "不一致" : "一致",
      keyCells: entry.keyCells,
      l: lValues,
      r: rValues,
      diffs,
      lc: entry.count,
      rc: other.count,
    });
  }
  for (const [key, entry] of R.map) {
    if (L.map.has(key)) continue;
    out.push({ status: "仅右表", keyCells: entry.keyCells, l: empty, r: entry.values.map(aggregateValues), diffs: [], lc: 0, rc: entry.count });
  }

  const numeric = Array.from({ length: compareCount }, (_, j) => {
    let sawNumber = false;
    for (const row of out) {
      for (const v of [row.l[j], row.r[j]]) {
        if (v.kind === "text") return false;
        if (v.kind === "number") sawNumber = true;
      }
    }
    return sawNumber;
  });

  const order: ReconcileStatus[] = ["不一致", "仅左表", "仅右表", "一致"];
  const visible = out
    .filter((row) => options.includeMatched || row.status !== "一致")
    .sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status));
  const showCounts = L.duplicateKeys > 0 || R.duplicateKeys > 0;

  const header: unknown[] = ["状态", "差异列", ...options.keyHeaders];
  const columnRoles: ReconcileResult["columnRoles"] = [{ role: "status" }, { role: "status" }];
  options.keyHeaders.forEach((_, i) => columnRoles.push({ role: "key", index: i }));
  options.compareHeaders.forEach((h, j) => {
    header.push(`${h}_${left.label}`, `${h}_${right.label}`);
    columnRoles.push({ role: "left", index: j }, { role: "right", index: j });
    if (numeric[j]) {
      header.push(`${h}_差额`);
      columnRoles.push({ role: "diff", index: j });
    }
  });
  if (showCounts) {
    header.push(`${left.label}行数`, `${right.label}行数`);
    columnRoles.push({ role: "count" }, { role: "count" });
  }

  const grid: unknown[][] = [header];
  for (const row of visible) {
    const cells: unknown[] = [row.status, row.diffs.join("、"), ...row.keyCells];
    for (let j = 0; j < compareCount; j++) {
      cells.push(aggregatedCell(row.l[j]), aggregatedCell(row.r[j]));
      if (numeric[j]) {
        const lv = row.l[j].kind === "number" ? (row.l[j] as { value: number }).value : 0;
        const rv = row.r[j].kind === "number" ? (row.r[j] as { value: number }).value : 0;
        cells.push(tidyNumber(lv - rv));
      }
    }
    if (showCounts) cells.push(row.lc, row.rc);
    grid.push(cells);
  }

  const columns: ReconcileColumnStats[] = options.compareHeaders.map((h, j) => {
    const stats: ReconcileColumnStats = { header: h, numeric: numeric[j], mismatched: out.filter((row) => row.diffs.includes(h)).length };
    if (numeric[j]) {
      const total = (side: "l" | "r") =>
        tidyNumber(out.reduce((s, row) => s + (row[side][j].kind === "number" ? (row[side][j] as { value: number }).value : 0), 0));
      stats.left_total = total("l");
      stats.right_total = total("r");
      stats.difference = tidyNumber(stats.left_total - stats.right_total);
    }
    return stats;
  });

  return {
    grid,
    matched: out.filter((row) => row.status === "一致").length,
    mismatched: out.filter((row) => row.status === "不一致").length,
    onlyLeft: out.filter((row) => row.status === "仅左表").length,
    onlyRight: out.filter((row) => row.status === "仅右表").length,
    leftRows: L.rows,
    rightRows: R.rows,
    leftBlankKeyRows: L.blankKeyRows,
    rightBlankKeyRows: R.blankKeyRows,
    leftDuplicateKeys: L.duplicateKeys,
    rightDuplicateKeys: R.duplicateKeys,
    columns,
    columnRoles,
  };
}

// ---------------- 多表纵向合并 ----------------

export interface StackSource {
  label: string;
  headers: string[];
  rows: unknown[][];
  formats?: string[];
}

export interface StackOptions {
  /** 来源列表头；null 表示不加来源列 */
  sourceHeader?: string | null;
  /** 只输出这些列并按此顺序 */
  columns?: string[];
}

export interface StackSourceReport {
  label: string;
  rows: number;
  /** 输出里有、这张表没有的列（留空） */
  missing: string[];
  /** 这张表有、但不在指定 columns 里而被丢弃的列 */
  dropped: string[];
}

export interface StackResult {
  grid: unknown[][];
  formats: (string | undefined)[];
  sources: StackSourceReport[];
  /** 指定 columns 里在任何表都找不到的列 */
  unknownColumns: string[];
}

const headerKey = (h: string) => normalizeKey(h, { ignoreCase: true });

/** 同一张表里重名的表头依次加 _2、_3，避免两列写进同一个输出列 */
function dedupeHeaders(headers: string[]): string[] {
  const seen = new Map<string, number>();
  return headers.map((h) => {
    const key = headerKey(h);
    const n = (seen.get(key) || 0) + 1;
    seen.set(key, n);
    return n === 1 ? h : `${h}_${n}`;
  });
}

export function stackTables(sources: StackSource[], options: StackOptions = {}): StackResult {
  const effective = sources.map((s) => dedupeHeaders(s.headers.map((h) => String(h ?? "").trim())));
  const outHeaders: string[] = [];
  const outIndex = new Map<string, number>();
  const addHeader = (h: string) => {
    const key = headerKey(h);
    if (!key || outIndex.has(key)) return;
    outIndex.set(key, outHeaders.length);
    outHeaders.push(h);
  };
  const restricted = Array.isArray(options.columns) && options.columns.length > 0;
  if (restricted) options.columns!.forEach((h) => addHeader(String(h).trim()));
  else effective.forEach((headers) => headers.forEach(addHeader));

  let sourceHeader = options.sourceHeader === undefined ? "来源" : options.sourceHeader;
  if (sourceHeader && outIndex.has(headerKey(sourceHeader))) sourceHeader = `${sourceHeader}_Sheet`;
  const lead = sourceHeader ? 1 : 0;
  const width = lead + outHeaders.length;
  const grid: unknown[][] = [[...(sourceHeader ? [sourceHeader] : []), ...outHeaders]];
  const formats: (string | undefined)[] = Array.from({ length: width }, () => undefined);
  const seenInAny = new Set<number>();

  const reports = sources.map((source, s) => {
    const mapping = effective[s].map((h) => outIndex.get(headerKey(h)) ?? -1);
    mapping.forEach((out, c) => {
      if (out < 0) return;
      seenInAny.add(out);
      const fmt = source.formats?.[c];
      if (formats[lead + out] === undefined && fmt) formats[lead + out] = fmt;
    });
    let rows = 0;
    for (const row of source.rows) {
      if (isBlankRow(row)) continue;
      const cells: unknown[] = Array.from({ length: width }, () => "");
      if (sourceHeader) cells[0] = source.label;
      mapping.forEach((out, c) => {
        if (out >= 0) cells[lead + out] = row[c] ?? "";
      });
      grid.push(cells);
      rows++;
    }
    const present = new Set(mapping.filter((m) => m >= 0));
    return {
      label: source.label,
      rows,
      missing: outHeaders.filter((_, i) => !present.has(i)),
      dropped: effective[s].filter((_, c) => mapping[c] < 0 && headerKey(effective[s][c])),
    };
  });

  return {
    grid,
    formats,
    sources: reports,
    unknownColumns: restricted ? outHeaders.filter((_, i) => !seenInAny.has(i)) : [],
  };
}

// ---------------- 按列拆表 ----------------

export interface SplitGroup {
  label: string;
  rows: unknown[][];
}

export function splitByKey(rows: unknown[][], column: number, keyOptions?: KeyOptions): SplitGroup[] {
  const map = new Map<string, { label: string; sortKey: unknown; rows: unknown[][] }>();
  for (const row of rows) {
    if (isBlankRow(row)) continue;
    const raw = row[column];
    const key = normalizeKey(raw, keyOptions);
    let group = map.get(key);
    if (!group) {
      const label = key === "" ? BLANK_LABEL : String(displayCell(raw));
      group = { label, sortKey: key === "" ? null : typeof raw === "number" ? raw : key, rows: [] };
      map.set(key, group);
    }
    group.rows.push(row);
  }
  return [...map.values()]
    .sort((a, b) => compareCells(a.sortKey, b.sortKey))
    .map(({ label, rows: groupRows }) => ({ label, rows: groupRows }));
}

// ---------------- 宽表转长表（逆透视） ----------------

export interface UnpivotValueColumn {
  index: number;
  /** 写进「属性」列的值，通常是原表头（可以是日期、数字） */
  label: unknown;
}

export interface UnpivotResult {
  rows: unknown[][];
  sourceRows: number;
  skippedBlanks: number;
}

/** 每个源行 × 每个值列展开成一行：[...标识列, 值列的表头, 值]。整行为空的源行跳过。 */
export function unpivotRows(
  rows: unknown[][],
  idColumns: number[],
  valueColumns: UnpivotValueColumn[],
  options: { skipBlanks?: boolean } = {}
): UnpivotResult {
  const skipBlanks = options.skipBlanks !== false;
  const out: unknown[][] = [];
  let sourceRows = 0;
  let skippedBlanks = 0;
  for (const row of rows) {
    if (isBlankRow(row)) continue;
    sourceRows++;
    const ids = idColumns.map((i) => row[i] ?? "");
    for (const col of valueColumns) {
      const value = row[col.index];
      if (isEmptyCell(value) && skipBlanks) {
        skippedBlanks++;
        continue;
      }
      out.push([...ids, col.label, value ?? ""]);
    }
  }
  return { rows: out, sourceRows, skippedBlanks };
}

// ---------------- 两个版本逐格比对 ----------------

export type DiffKind = "修改" | "新增行" | "删除行";

export interface DiffSide {
  headers: string[];
  rows: unknown[][];
  /** 第一条数据行在工作表里的行号（1 基） */
  firstRow: number;
  /** 区域第一列在工作表里的 0 基列号 */
  firstColumn: number;
  /** 键列（区域内 0 基列号）；为空时按行位置对齐 */
  keys: number[];
}

export interface DiffOptions {
  /** true：按表头名对齐列（列顺序变了也能比）；false：按位置 */
  matchByHeader: boolean;
  tolerance?: number;
  ignoreCase?: boolean;
  /** 最多记录多少条差异，超出只计数 */
  maxChanges?: number;
}

export interface DiffChange {
  kind: DiffKind;
  rowLabel: string;
  column: string;
  before: unknown;
  after: unknown;
  oldAddress: string;
  newAddress: string;
  /** 新版本里要标色的位置（工作表 0 基行列）；删除行没有 */
  mark?: { row: number; column: number; columns: number };
}

export interface DiffResult {
  changes: DiffChange[];
  /** 超过 maxChanges 未记录的差异条数 */
  omitted: number;
  changedCells: number;
  changedRows: number;
  unchangedRows: number;
  addedRows: number;
  removedRows: number;
  comparedColumns: string[];
  onlyOldColumns: string[];
  onlyNewColumns: string[];
  duplicateKeys: number;
  blankKeyRows: number;
}

const ROW_SUMMARY_CHARS = 120;

function cellText(value: unknown): string {
  return isEmptyCell(value) ? "" : String(displayCell(value));
}

function rowSummary(row: unknown[]): string {
  const text = row.map(cellText).filter(Boolean).join(" | ");
  return text.length > ROW_SUMMARY_CHARS ? `${text.slice(0, ROW_SUMMARY_CHARS)}…` : text;
}

/** 看起来一样就算一样：数字按误差比，其余去首尾空白后比文本（100 与 "100" 相同） */
export function sameCell(a: unknown, b: unknown, tolerance = 0, ignoreCase = false): boolean {
  if (typeof a === "number" && typeof b === "number") {
    return Math.abs(a - b) <= tolerance;
  }
  const options = { halfWidth: false, ignoreCase };
  return normalizeKey(a, options) === normalizeKey(b, options);
}

function pairColumns(
  oldHeaders: string[],
  newHeaders: string[],
  byHeader: boolean
): { pairs: [number, number][]; onlyOld: number[]; onlyNew: number[] } {
  const pairs: [number, number][] = [];
  const usedNew = new Set<number>();
  const onlyOld: number[] = [];
  oldHeaders.forEach((header, i) => {
    let j = -1;
    if (byHeader) {
      const key = normalizeKey(header);
      j = newHeaders.findIndex((h, k) => !usedNew.has(k) && normalizeKey(h) === key);
    } else if (i < newHeaders.length) {
      j = i;
    }
    if (j < 0) onlyOld.push(i);
    else {
      usedNew.add(j);
      pairs.push([i, j]);
    }
  });
  const onlyNew = newHeaders.map((_, k) => k).filter((k) => !usedNew.has(k));
  return { pairs, onlyOld, onlyNew };
}

export function diffTables(oldSide: DiffSide, newSide: DiffSide, options: DiffOptions): DiffResult {
  const keyed = oldSide.keys.length > 0;
  if (keyed && oldSide.keys.length !== newSide.keys.length) throw new Error("新旧两边的键列数量必须一致");
  const tolerance = Math.max(0, Number(options.tolerance) || 0);
  const maxChanges = options.maxChanges ?? 20000;
  const { pairs, onlyOld, onlyNew } = pairColumns(oldSide.headers, newSide.headers, options.matchByHeader);
  const width = Math.max(1, newSide.headers.length);
  const result: DiffResult = {
    changes: [],
    omitted: 0,
    changedCells: 0,
    changedRows: 0,
    unchangedRows: 0,
    addedRows: 0,
    removedRows: 0,
    comparedColumns: pairs.map(([, j]) => newSide.headers[j]),
    onlyOldColumns: onlyOld.map((i) => oldSide.headers[i]),
    onlyNewColumns: onlyNew.map((j) => newSide.headers[j]),
    duplicateKeys: 0,
    blankKeyRows: 0,
  };
  const record = (change: DiffChange) => {
    if (result.changes.length < maxChanges) result.changes.push(change);
    else result.omitted++;
  };
  const oldAddr = (r: number, c: number) => `${indexToColLetter(oldSide.firstColumn + c)}${oldSide.firstRow + r}`;
  const newAddr = (r: number, c: number) => `${indexToColLetter(newSide.firstColumn + c)}${newSide.firstRow + r}`;
  const keyOf = (side: DiffSide, row: unknown[]) =>
    side.keys.map((i) => normalizeKey(row[i], { ignoreCase: options.ignoreCase })).join("\u0001");
  const hasKey = (key: string) => key.replace(/\u0001/g, "") !== "";
  const labelOf = (side: DiffSide, row: unknown[], r: number) =>
    keyed ? side.keys.map((i) => cellText(row[i])).join("+") : `第 ${side.firstRow + r} 行`;

  const compareRow = (oi: number, ni: number) => {
    const oldRow = oldSide.rows[oi] || [];
    const newRow = newSide.rows[ni] || [];
    let changed = 0;
    for (const [oc, nc] of pairs) {
      if (sameCell(oldRow[oc], newRow[nc], tolerance, options.ignoreCase)) continue;
      changed++;
      record({
        kind: "修改",
        rowLabel: labelOf(newSide, newRow, ni),
        column: newSide.headers[nc],
        before: oldRow[oc] ?? "",
        after: newRow[nc] ?? "",
        oldAddress: oldAddr(oi, oc),
        newAddress: newAddr(ni, nc),
        mark: { row: newSide.firstRow - 1 + ni, column: newSide.firstColumn + nc, columns: 1 },
      });
    }
    result.changedCells += changed;
    if (changed) result.changedRows++;
    else result.unchangedRows++;
  };
  const removed = (oi: number) => {
    const row = oldSide.rows[oi];
    const r = oldSide.firstRow + oi;
    result.removedRows++;
    record({ kind: "删除行", rowLabel: labelOf(oldSide, row, oi), column: "(整行)", before: rowSummary(row), after: "", oldAddress: `${r}:${r}`, newAddress: "" });
  };
  const added = (ni: number) => {
    const row = newSide.rows[ni];
    const r = newSide.firstRow + ni;
    result.addedRows++;
    record({
      kind: "新增行",
      rowLabel: labelOf(newSide, row, ni),
      column: "(整行)",
      before: "",
      after: rowSummary(row),
      oldAddress: "",
      newAddress: `${r}:${r}`,
      mark: { row: r - 1, column: newSide.firstColumn, columns: width },
    });
  };

  if (!keyed) {
    const n = Math.max(oldSide.rows.length, newSide.rows.length);
    for (let i = 0; i < n; i++) {
      const hasOld = i < oldSide.rows.length && !isBlankRow(oldSide.rows[i]);
      const hasNew = i < newSide.rows.length && !isBlankRow(newSide.rows[i]);
      if (hasOld && i >= newSide.rows.length) removed(i);
      else if (hasNew && i >= oldSide.rows.length) added(i);
      else if (hasOld || hasNew) compareRow(i, i);
    }
    return result;
  }

  // 按键对齐：同一个键出现多次时按出现顺序一一配对
  const queues = new Map<string, number[]>();
  const oldCounts = new Map<string, number>();
  const newCounts = new Map<string, number>();
  newSide.rows.forEach((row, ni) => {
    if (isBlankRow(row)) return;
    const key = keyOf(newSide, row);
    if (!hasKey(key)) {
      result.blankKeyRows++;
      return;
    }
    if (!queues.has(key)) queues.set(key, []);
    queues.get(key)!.push(ni);
    newCounts.set(key, (newCounts.get(key) || 0) + 1);
  });
  const paired = new Set<number>();
  oldSide.rows.forEach((row, oi) => {
    if (isBlankRow(row)) return;
    const key = keyOf(oldSide, row);
    if (!hasKey(key)) {
      result.blankKeyRows++;
      return;
    }
    oldCounts.set(key, (oldCounts.get(key) || 0) + 1);
    const ni = queues.get(key)?.shift();
    if (ni == null) removed(oi);
    else {
      paired.add(ni);
      compareRow(oi, ni);
    }
  });
  newSide.rows.forEach((row, ni) => {
    if (!paired.has(ni) && !isBlankRow(row) && hasKey(keyOf(newSide, row))) added(ni);
  });
  const duplicated = new Set<string>();
  for (const counts of [oldCounts, newCounts]) for (const [key, n] of counts) if (n > 1) duplicated.add(key);
  result.duplicateKeys = duplicated.size;
  return result;
}
