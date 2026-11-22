/**
 * 从单元格文本里提取手机号、邮箱、身份证号、数字、日期等，以及由身份证号推出出生日期/性别/年龄。
 * 纯函数，便于不启动 Excel 也能核对。只用前瞻、不用后顾断言，兼容较老的 Office 内核。
 */
import { excelSerialToIso, parseLooseNumber, toHalfWidth, toIsoDate } from "./cleanValues";

export type ExtractTarget =
  | "mobile"
  | "email"
  | "id_card"
  | "number"
  | "date"
  | "url"
  | "chinese"
  | "english"
  | "birth_date"
  | "gender"
  | "age"
  | "id_check";

export const EXTRACT_TARGET_LABELS: Record<ExtractTarget, string> = {
  mobile: "手机号",
  email: "邮箱",
  id_card: "身份证号",
  number: "数字",
  date: "日期",
  url: "网址",
  chinese: "中文",
  english: "英文",
  birth_date: "出生日期",
  gender: "性别",
  age: "年龄",
  id_check: "身份证校验",
};

/** 结果是日期的目标：写入后套日期格式 */
export const DATE_TARGETS = new Set<ExtractTarget>(["date", "birth_date"]);
/** 由身份证号推出的目标 */
const ID_TARGETS = new Set<ExtractTarget>(["id_card", "birth_date", "gender", "age", "id_check"]);

const TARGET_ALIASES: Record<string, ExtractTarget> = {
  mobile: "mobile",
  phone: "mobile",
  cellphone: "mobile",
  手机: "mobile",
  手机号: "mobile",
  email: "email",
  mail: "email",
  邮箱: "email",
  id_card: "id_card",
  idcard: "id_card",
  id: "id_card",
  身份证: "id_card",
  身份证号: "id_card",
  number: "number",
  amount: "number",
  数字: "number",
  金额: "number",
  date: "date",
  日期: "date",
  url: "url",
  link: "url",
  网址: "url",
  chinese: "chinese",
  中文: "chinese",
  english: "english",
  英文: "english",
  birth_date: "birth_date",
  birthday: "birth_date",
  birth: "birth_date",
  出生日期: "birth_date",
  gender: "gender",
  sex: "gender",
  性别: "gender",
  age: "age",
  年龄: "age",
  id_check: "id_check",
  id_valid: "id_check",
  check: "id_check",
  校验: "id_check",
};

export function normalizeExtractTargets(raw: unknown): ExtractTarget[] {
  const list = Array.isArray(raw) ? raw : raw != null && raw !== "" ? [raw] : [];
  const out: ExtractTarget[] = [];
  for (const item of list) {
    const key = String(item ?? "").trim().toLowerCase().replace(/-/g, "_");
    const target = TARGET_ALIASES[key];
    if (!target) {
      throw new Error(`不支持的提取目标「${item}」，可用 ${Object.keys(EXTRACT_TARGET_LABELS).join(" / ")}。自定义正则请改用 run_script`);
    }
    if (!out.includes(target)) out.push(target);
  }
  if (!out.length) throw new Error("extract_text 需要 targets（要提取什么），如 [\"mobile\"] 或 [\"birth_date\",\"gender\",\"age\"]");
  return out;
}

// ---------------- 身份证 ----------------

const ID_WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2];
const ID_CHECK_CODES = "10X98765432";

export interface ChineseId {
  id: string;
  /** 出生日期 yyyy-mm-dd；日期不合法时为 null */
  birth: string | null;
  gender: "男" | "女";
  /** 校验结论：「正确」或具体问题 */
  check: string;
  valid: boolean;
}

export function idCheckDigit(first17: string): string {
  let sum = 0;
  for (let i = 0; i < 17; i++) sum += Number(first17[i]) * ID_WEIGHTS[i];
  return ID_CHECK_CODES[sum % 11];
}

/** 解析 18 位或 15 位（老一代，19xx 年出生）身份证号；格式不对返回 null */
export function parseChineseId(raw: string, today: Date = new Date()): ChineseId | null {
  const id = raw.trim().toUpperCase();
  let birthDigits: string;
  let genderDigit: number;
  if (/^\d{17}[\dX]$/.test(id)) {
    birthDigits = id.slice(6, 14);
    genderDigit = Number(id[16]);
  } else if (/^\d{15}$/.test(id)) {
    birthDigits = `19${id.slice(6, 12)}`;
    genderDigit = Number(id[14]);
  } else {
    return null;
  }
  const birth = toIsoDate(Number(birthDigits.slice(0, 4)), Number(birthDigits.slice(4, 6)), Number(birthDigits.slice(6, 8)));
  const future = birth != null && birth > isoDay(today);
  const gender = genderDigit % 2 === 1 ? "男" : "女";
  let check = "正确";
  if (!birth || future) check = "出生日期无效";
  else if (id.length === 18 && idCheckDigit(id) !== id[17]) check = "校验位错误";
  return { id, birth: birth && !future ? birth : null, gender, check, valid: check === "正确" };
}

function isoDay(d: Date): string {
  return toIsoDate(d.getFullYear(), d.getMonth() + 1, d.getDate()) || "";
}

/** 周岁：按今天是否已过生日 */
export function ageOn(birthIso: string, today: Date = new Date()): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(birthIso);
  if (!m) return null;
  const y = Number(m[1]);
  const md = Number(m[2]) * 100 + Number(m[3]);
  const todayMd = (today.getMonth() + 1) * 100 + today.getDate();
  const age = today.getFullYear() - y - (todayMd < md ? 1 : 0);
  return age >= 0 ? age : null;
}

// ---------------- 文本匹配 ----------------

function allMatches(re: RegExp, text: string, group = 0): string[] {
  const out: string[] = [];
  const g = new RegExp(re.source, re.flags.includes("g") ? re.flags : `${re.flags}g`);
  let m: RegExpExecArray | null;
  while ((m = g.exec(text))) {
    if (m[group]) out.push(m[group]);
    if (m[0] === "") g.lastIndex++;
  }
  return out;
}

const MOBILE_RE = /(^|\D)(?:\+?86[\s-]?)?(1[3-9]\d(?:[\s-]?\d{4}){2})(?!\d)/;
const EMAIL_RE = /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/;
const ID_RE = /(^|[^\dXx])(\d{17}[\dXx]|\d{15})(?![\dXx])/;
const URL_RE = /(?:https?:\/\/|www\.)[^\s，。、；;,"'<>（）()【】一-鿿]+/i;
const CHINESE_RE = /[㐀-鿿]+/;
const ENGLISH_RE = /[A-Za-z]+(?:[ '.-][A-Za-z]+)*/;
const NUMBER_RE = /(^|[^\d.])(-?)(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(%?)/;
const DATE_RES: { re: RegExp; order: "ymd" }[] = [
  { re: /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?/, order: "ymd" },
  { re: /(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?!\d)/, order: "ymd" },
  { re: /(^|\D)(\d{4})(\d{2})(\d{2})(?!\d)/, order: "ymd" },
];

function extractNumbers(text: string): number[] {
  const out: number[] = [];
  const g = new RegExp(NUMBER_RE.source, "g");
  let m: RegExpExecArray | null;
  while ((m = g.exec(text))) {
    // 负号只在开头或空白、括号、冒号后才算，避免把「X-200」「2024-03」的连字符当负号
    const sign = m[2] && (m[1] === "" || /[\s(（:：=]/.test(m[1])) ? -1 : 1;
    const n = parseLooseNumber(m[3]);
    if (n != null) out.push((sign * n) / (m[4] ? 100 : 1));
  }
  return out;
}

function extractDates(text: string): string[] {
  const found: { at: number; iso: string }[] = [];
  for (const { re } of DATE_RES) {
    const g = new RegExp(re.source, "g");
    let m: RegExpExecArray | null;
    while ((m = g.exec(text))) {
      const parts = m.length === 5 ? [m[2], m[3], m[4]] : [m[1], m[2], m[3]];
      const iso = toIsoDate(Number(parts[0]), Number(parts[1]), Number(parts[2]));
      if (iso && !found.some((f) => f.iso === iso)) found.push({ at: m.index, iso });
    }
  }
  return found.sort((a, b) => a.at - b.at).map((f) => f.iso);
}

export interface ExtractOptions {
  today?: Date;
  /** true 时取出全部匹配并用 separator 连接；默认只取第一个 */
  all?: boolean;
  separator?: string;
}

export interface ExtractOutcome {
  value: unknown;
  /** 原值非空但没找到 */
  missing: boolean;
  /** 身份证号被存成了数字，末几位已丢失 */
  precisionLost?: boolean;
  /** 找到了身份证号但校验未通过 */
  invalidId?: boolean;
}

function pick(list: string[], options: ExtractOptions): string {
  if (!list.length) return "";
  return options.all ? [...new Set(list)].join(options.separator ?? "、") : list[0];
}

function idFromCell(value: unknown, today: Date): { id: ChineseId | null; precisionLost: boolean } {
  if (typeof value === "number") {
    // Excel 数字只有 15 位有效数字：18 位身份证存成数字后，出生日期还在，第 17 位（性别）和校验位已不可信
    if (!Number.isFinite(value) || value < 1e14) return { id: null, precisionLost: false };
    const digits = value.toFixed(0);
    if (digits.length === 15) return { id: parseChineseId(digits, today), precisionLost: false };
    if (digits.length !== 18) return { id: null, precisionLost: false };
    const birth = toIsoDate(Number(digits.slice(6, 10)), Number(digits.slice(10, 12)), Number(digits.slice(12, 14)));
    return { id: birth ? { id: digits, birth, gender: "男", check: "存成数字，后几位已丢失", valid: false } : null, precisionLost: true };
  }
  const text = toHalfWidth(String(value ?? ""));
  const m = new RegExp(ID_RE.source).exec(text);
  return { id: m ? parseChineseId(m[2], today) : null, precisionLost: false };
}

export function extractValue(value: unknown, target: ExtractTarget, options: ExtractOptions = {}): ExtractOutcome {
  if (value == null || value === "") return { value: "", missing: false };
  const today = options.today ?? new Date();
  if (ID_TARGETS.has(target)) {
    const { id, precisionLost } = idFromCell(value, today);
    if (!id) return { value: "", missing: true };
    const unsure = precisionLost && target !== "birth_date" && target !== "age";
    if (unsure) return { value: target === "id_check" ? id.check : "", missing: false, precisionLost: true };
    const outcome: ExtractOutcome = { value: "", missing: false, precisionLost, invalidId: !precisionLost && !id.valid };
    if (target === "id_card") outcome.value = id.id;
    else if (target === "id_check") outcome.value = id.check;
    else if (target === "gender") outcome.value = id.gender;
    else if (target === "birth_date") outcome.value = id.birth ?? "";
    else if (target === "age") outcome.value = id.birth ? ageOn(id.birth, today) ?? "" : "";
    return outcome;
  }
  if (target === "number" && typeof value === "number") return { value, missing: false };
  if (target === "date" && typeof value === "number") {
    // 已经是 Excel 日期序列：转成 yyyy-mm-dd 写回，Excel 会按日期识别
    const iso = excelSerialToIso(value);
    return iso ? { value: iso, missing: false } : { value: "", missing: true };
  }
  const text = toHalfWidth(String(value));
  let found: string[] = [];
  if (target === "mobile") found = allMatches(MOBILE_RE, text, 2).map((s) => s.replace(/[\s-]/g, ""));
  else if (target === "email") found = allMatches(EMAIL_RE, text);
  else if (target === "url") found = allMatches(URL_RE, text).map((s) => s.replace(/[.。]+$/, ""));
  else if (target === "chinese") found = allMatches(CHINESE_RE, String(value));
  else if (target === "english") found = allMatches(ENGLISH_RE, text);
  else if (target === "date") found = extractDates(text);
  else if (target === "number") {
    const numbers = extractNumbers(text);
    if (!numbers.length) return { value: "", missing: true };
    return { value: options.all ? numbers.join(options.separator ?? "、") : numbers[0], missing: false };
  }
  return found.length ? { value: pick(found, options), missing: false } : { value: "", missing: true };
}
