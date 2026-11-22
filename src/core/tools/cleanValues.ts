/** 清洗单元格：去空格、宽松数字/日期解析。纯函数，便于不启动 Excel 也能核对。 */

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

export function toIsoDate(year: number, month: number, day: number): string | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
  if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (dt.getUTCFullYear() !== year || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== day) return null;
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

/** Excel 1900 日期系统：序列 25569 = 1970-01-01。 */
export function excelSerialToIso(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) return null;
  const utc = Date.UTC(1970, 0, 1) + Math.round((serial - 25569) * 86400000);
  const d = new Date(utc);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
}

export function parseLooseNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value instanceof Date) return null;
  let s = String(value ?? "").trim();
  if (!s) return null;
  const percent = /%$/.test(s);
  s = s
    .replace(/[%]/g, "")
    .replace(/,/g, "")
    .replace(/[¥￥$€£]/g, "")
    .replace(/元/g, "")
    .trim();
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  const n = Number(s);
  if (!Number.isFinite(n)) return null;
  return percent ? n / 100 : n;
}

export function extractYear(value: unknown): number | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getFullYear();
  const s = String(value ?? "").trim();
  const m = /(\d{4})/.exec(s);
  if (!m) return null;
  const y = parseInt(m[1], 10);
  return y >= 1900 && y <= 2100 ? y : null;
}

export function inferDefaultYear(values: unknown[], fallback: number): number {
  const counts = new Map<number, number>();
  for (const v of values) {
    const y = extractYear(v);
    if (y == null) continue;
    counts.set(y, (counts.get(y) || 0) + 1);
  }
  let best: number | null = null;
  let bestN = 0;
  for (const [y, n] of counts) {
    if (n > bestN) {
      best = y;
      bestN = n;
    }
  }
  return best ?? fallback;
}

export function parseLooseDate(value: unknown, options?: { defaultYear?: number }): string | null {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toIsoDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    // 已经是 Excel 日期序列，保留数字，避免写成文本后无法参与日期运算。
    return null;
  }
  const s = String(value).trim();
  if (!s) return null;

  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/.exec(s);
  if (m) return toIsoDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));

  m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/.exec(s);
  if (m) return toIsoDate(parseInt(m[3], 10), parseInt(m[1], 10), parseInt(m[2], 10));

  m = /^(\d{1,2})月(\d{1,2})日/.exec(s);
  if (m) {
    const year = options?.defaultYear;
    if (year == null) return null;
    return toIsoDate(year, parseInt(m[1], 10), parseInt(m[2], 10));
  }

  m = /^(\d{4})年(\d{1,2})月(\d{1,2})/.exec(s);
  if (m) return toIsoDate(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));

  return null;
}

export function trimCell(value: unknown): unknown {
  if (typeof value !== "string") return value;
  return value.replace(/\s+/g, " ").trim();
}

/** 全角字母、数字、符号转半角，全角空格转普通空格 */
export function toHalfWidth(text: string): string {
  return text
    .replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/　/g, " ");
}

const RMB_DIGITS = "零壹贰叁肆伍陆柒捌玖";
const RMB_UNITS = ["", "拾", "佰", "仟"];
const RMB_SECTIONS = ["", "万", "亿", "万亿"];

function rmbSection(section: number): string {
  let out = "";
  let zero = false;
  for (let p = 3; p >= 0; p--) {
    const d = Math.floor(section / 10 ** p) % 10;
    if (d === 0) {
      if (out) zero = true;
      continue;
    }
    if (zero) out += "零";
    zero = false;
    out += RMB_DIGITS[d] + RMB_UNITS[p];
  }
  return out;
}

/** 人民币金额大写，如 1234.56 →「壹仟贰佰叁拾肆元伍角陆分」；按分四舍五入，超过万亿级返回 null。 */
export function toRmbUpper(amount: number): string | null {
  if (!Number.isFinite(amount)) return null;
  const cents = Math.round(Math.abs(amount) * 100);
  if (cents >= 1e17) return null;
  const yuan = Math.floor(cents / 100);
  const jiao = Math.floor(cents / 10) % 10;
  const fen = cents % 10;
  const sections: number[] = [];
  for (let x = yuan; x > 0; x = Math.floor(x / 10000)) sections.push(x % 10000);
  let text = "";
  let gap = false;
  for (let i = sections.length - 1; i >= 0; i--) {
    const section = sections[i];
    if (section === 0) {
      if (text) gap = true;
      continue;
    }
    if (text && (gap || section < 1000)) text += "零";
    gap = false;
    text += rmbSection(section) + RMB_SECTIONS[i];
  }
  if (text) text += "元";
  if (!jiao && !fen) text = (text || "零元") + "整";
  else {
    if (jiao) text += `${RMB_DIGITS[jiao]}角`;
    else if (text) text += "零";
    if (fen) text += `${RMB_DIGITS[fen]}分`;
  }
  return amount < 0 && cents > 0 ? `负${text}` : text;
}

export type CleanOp = "trim" | "to_number" | "to_date" | "to_halfwidth" | "remove_spaces" | "upper" | "lower" | "rmb_upper";

const CLEAN_OP_ALIASES: Record<string, CleanOp> = {
  trim: "trim",
  strip: "trim",
  collapse_space: "trim",
  to_number: "to_number",
  number: "to_number",
  parse_number: "to_number",
  to_date: "to_date",
  date: "to_date",
  parse_date: "to_date",
  to_halfwidth: "to_halfwidth",
  halfwidth: "to_halfwidth",
  to_half_width: "to_halfwidth",
  full_to_half: "to_halfwidth",
  remove_spaces: "remove_spaces",
  remove_all_spaces: "remove_spaces",
  no_spaces: "remove_spaces",
  upper: "upper",
  to_upper: "upper",
  uppercase: "upper",
  lower: "lower",
  to_lower: "lower",
  lowercase: "lower",
  rmb_upper: "rmb_upper",
  rmb: "rmb_upper",
  amount_upper: "rmb_upper",
  chinese_amount: "rmb_upper",
};

export function normalizeCleanOps(raw: unknown): CleanOp[] {
  const list = Array.isArray(raw) ? raw : raw != null && raw !== "" ? [raw] : ["trim"];
  const out: CleanOp[] = [];
  for (const item of list) {
    const op = CLEAN_OP_ALIASES[String(item).trim().toLowerCase().replace(/-/g, "_")];
    if (!op) {
      throw new Error(`不支持的清洗操作「${item}」，可用 trim / to_number / to_date / to_halfwidth / remove_spaces / upper / lower / rmb_upper`);
    }
    out.push(op);
  }
  return out.length ? out : ["trim"];
}

export function applyCleanOps(
  value: unknown,
  ops: CleanOp[],
  options?: { defaultYear?: number }
): { value: unknown; changed: boolean; unparsed: boolean } {
  let cur: unknown = value;
  let unparsed = false;
  for (const op of ops) {
    if (op === "trim") {
      cur = trimCell(cur);
    } else if (op === "to_number") {
      if (cur == null || cur === "") continue;
      const n = parseLooseNumber(cur);
      if (n == null) unparsed = true;
      else cur = n;
    } else if (op === "to_date") {
      if (cur == null || cur === "") continue;
      if (typeof cur === "number" && Number.isFinite(cur)) continue;
      const iso = parseLooseDate(cur, options);
      if (!iso) unparsed = true;
      else cur = iso;
    } else if (op === "rmb_upper") {
      if (cur == null || cur === "") continue;
      const n = parseLooseNumber(cur);
      const text = n == null ? null : toRmbUpper(n);
      if (text == null) unparsed = true;
      else cur = text;
    } else if (typeof cur === "string") {
      if (op === "to_halfwidth") cur = toHalfWidth(cur);
      else if (op === "remove_spaces") cur = cur.replace(/\s+/g, "");
      else if (op === "upper") cur = cur.toUpperCase();
      else if (op === "lower") cur = cur.toLowerCase();
    }
  }
  const changed = JSON.stringify(cur) !== JSON.stringify(value);
  return { value: cur, changed, unparsed };
}

export function guessColumnNumberFormat(header: string): string | null {
  const h = String(header || "").replace(/\s+/g, "");
  if (!h) return null;
  if (/日期|date|时间|time/i.test(h)) return "yyyy-mm-dd";
  if (/率|%|百分比|占比/i.test(h)) return "0.0%";
  if (/金额|单价|成本|合计|小计|缴费|收入|支出|价格|毛利|费用|销售额/i.test(h)) return "#,##0.00";
  if (/数量|件数|qty/i.test(h)) return "#,##0";
  return null;
}

export function isRowBlank(cells: unknown[]): boolean {
  return cells.every((v) => v === "" || v === null || v === undefined);
}
