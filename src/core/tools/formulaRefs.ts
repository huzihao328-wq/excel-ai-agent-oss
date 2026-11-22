/** Excel 公式引用与分列辅助：纯函数，便于不启动 Excel 也能核对。 */

export function colLetterToIndex(letter: string): number {
  let n = 0;
  for (const ch of letter.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToColLetter(index: number): string {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 把表头名、列字母或 0 基列号解析成区域内的相对列号。 */
export function resolveColSpec(headers: string[], spec: unknown, colStart: number, label: string): number {
  if (spec == null || spec === "") throw new Error(`${label} 不能为空`);
  if (typeof spec === "number") {
    const n = Math.trunc(spec);
    if (n < 0 || n >= headers.length) throw new Error(`${label} 列号 ${n} 超出范围 0..${headers.length - 1}`);
    return n;
  }
  const s = String(spec).trim();
  const byHeader = headers.findIndex((h) => String(h).trim() === s);
  if (byHeader >= 0) return byHeader;
  if (/^[A-Za-z]{1,3}$/.test(s)) {
    const rel = colLetterToIndex(s.toUpperCase()) - colStart;
    if (rel < 0 || rel >= headers.length) throw new Error(`${label} 列 ${s} 不在当前表区域内`);
    return rel;
  }
  if (/^\d+$/.test(s)) {
    const n = parseInt(s, 10);
    if (n < 0 || n >= headers.length) throw new Error(`${label} 列号 ${n} 超出范围`);
    return n;
  }
  const available = headers.map((h, i) => h || indexToColLetter(colStart + i)).filter(Boolean);
  throw new Error(`找不到${label}「${s}」。可用列: ${available.join("、")}`);
}

/** 工作表名写入公式时的引用：中文、空格、引号、以及看起来像单元格/R1C1 的名字需要加单引号。 */
export function quoteSheetName(name: string): string {
  const s = String(name || "").trim();
  if (!s) throw new Error("Sheet 名称不能为空");
  const looksLikeA1 = /^[A-Za-z]{1,3}\d+$/.test(s);
  const looksLikeR1C1 = /^[Rr]\d+[Cc]\d+$/.test(s);
  if (/^[A-Za-z_][A-Za-z0-9.]*$/.test(s) && !looksLikeA1 && !looksLikeR1C1) return s;
  return `'${s.replace(/'/g, "''")}'`;
}

export function columnA1Range(sheet: string, colLetter: string, startRow: number, endRow: number): string {
  const col = String(colLetter).toUpperCase();
  return `${quoteSheetName(sheet)}!$${col}$${startRow}:$${col}$${endRow}`;
}

export function excelStringLiteral(value: unknown): string {
  if (value == null) return '""';
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return `"${String(value).replace(/"/g, '""')}"`;
}

export function xlookupFormula(lookupCell: string, keyRange: string, valueRange: string, ifNotFoundLiteral = '""'): string {
  return `=XLOOKUP(${lookupCell},${keyRange},${valueRange},${ifNotFoundLiteral})`;
}

export function vlookupFormula(
  lookupCell: string,
  tableRange: string,
  colIndex: number,
  exact: boolean,
  ifNotFoundLiteral = '""'
): string {
  return `=IFERROR(VLOOKUP(${lookupCell},${tableRange},${colIndex},${exact ? "FALSE" : "TRUE"}),${ifNotFoundLiteral})`;
}

export function textJoinFormula(separator: string, cells: string[], ignoreEmpty = true): string {
  return `=TEXTJOIN(${excelStringLiteral(separator)},${ignoreEmpty ? "TRUE" : "FALSE"},${cells.join(",")})`;
}

export interface LookupReturnCol {
  letter: string;
  header: string;
}

export interface LookupColumnPlan {
  destKeyLetter: string;
  destFirstDataRow: number;
  destLastDataRow: number;
  lookupSheet: string;
  lookupKeyLetter: string;
  lookupFirstDataRow: number;
  lookupLastDataRow: number;
  returns: LookupReturnCol[];
  ifNotFound?: unknown;
  engine: "xlookup" | "vlookup";
  includeHeaders?: boolean;
}

export function buildLookupFormulaGrid(plan: LookupColumnPlan): string[][] {
  const ifNotFound = excelStringLiteral(plan.ifNotFound ?? "");
  const headers = plan.returns.map((r) => r.header);
  const rows: string[][] = plan.includeHeaders === false ? [] : [headers];
  const keyRange = columnA1Range(
    plan.lookupSheet,
    plan.lookupKeyLetter,
    plan.lookupFirstDataRow,
    plan.lookupLastDataRow
  );
  const keyIdx = colLetterToIndex(plan.lookupKeyLetter);
  for (let r = plan.destFirstDataRow; r <= plan.destLastDataRow; r++) {
    const lookupCell = `${plan.destKeyLetter}${r}`;
    const cells = plan.returns.map((ret) => {
      if (plan.engine === "vlookup") {
        const retIdx = colLetterToIndex(ret.letter);
        if (retIdx < keyIdx) {
          throw new Error(
            `VLOOKUP 要求返回列在匹配列右侧，列 ${ret.letter} 在 ${plan.lookupKeyLetter} 左侧。请改用 engine=xlookup`
          );
        }
        const tableRange = `${quoteSheetName(plan.lookupSheet)}!$${plan.lookupKeyLetter}$${plan.lookupFirstDataRow}:$${ret.letter}$${plan.lookupLastDataRow}`;
        return vlookupFormula(lookupCell, tableRange, retIdx - keyIdx + 1, true, ifNotFound);
      }
      const valueRange = columnA1Range(plan.lookupSheet, ret.letter, plan.lookupFirstDataRow, plan.lookupLastDataRow);
      return xlookupFormula(lookupCell, keyRange, valueRange, ifNotFound);
    });
    rows.push(cells);
  }
  return rows;
}

const DELIMITER_ALIASES: Record<string, string> = {
  comma: ",",
  ",": ",",
  "，": "，",
  chinese_comma: "，",
  comma_cn: "，",
  semicolon: ";",
  ";": ";",
  "；": "；",
  semicolon_cn: "；",
  tab: "\t",
  "\\t": "\t",
  space: " ",
  whitespace: "whitespace",
  at: "@",
  "@": "@",
  pipe: "|",
  "|": "|",
};

export function resolveDelimiter(raw: unknown): string {
  if (raw == null || String(raw) === "") {
    throw new Error("split_column 需要 delimiter，例如 ','、'@'、space、tab");
  }
  const original = String(raw);
  if (original === " ") return " ";
  const key = original.trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(DELIMITER_ALIASES, key)) return DELIMITER_ALIASES[key];
  return original;
}

/** 按分隔符拆一格。maxParts 有值时多余段并回最后一列。不自动补空列。 */
export function splitCellText(value: unknown, delimiter: string, maxParts?: number, trimParts = true): string[] {
  const s = value == null ? "" : String(value);
  let parts: string[];
  if (delimiter === "whitespace") {
    const t = s.trim();
    parts = t ? t.split(/\s+/) : [""];
  } else {
    parts = s.split(delimiter);
  }
  if (trimParts && delimiter !== "whitespace") parts = parts.map((p) => p.trim());
  if (maxParts != null && maxParts > 0 && parts.length > maxParts) {
    const sep = delimiter === "whitespace" ? " " : delimiter;
    return [...parts.slice(0, maxParts - 1), parts.slice(maxParts - 1).join(sep)];
  }
  return parts;
}

export function padSplitParts(parts: string[], width: number): string[] {
  const out = parts.slice(0, width);
  while (out.length < width) out.push("");
  return out;
}
