/**
 * 模型常见的参数小毛病在这里就地修好，而不是退回去让它再花一轮：
 * - 顶层 JSON 的轻微格式问题（代码块包裹、双重编码、尾逗号、单引号、Python 字面量、裸键名、字符串里的原始换行）；
 * - 字段名写成 camelCase 或常见同义词（sheetName、address、data…）；
 * - 布尔/数字写成字符串、数组写成 JSON 字符串、单个值没包成数组、枚举大小写不一致、可选字段填 null。
 * 截断的 JSON 永远不会被「补全」：那样会把半截数据写进表格。
 */

export type ParsedToolArguments =
  | { ok: true; value: Record<string, unknown>; repaired: boolean }
  | { ok: false; error: string };

const isRecord = (v: unknown): v is Record<string, unknown> => !!v && typeof v === "object" && !Array.isArray(v);

function tryParse(text: string): { value: unknown } | null {
  try {
    return { value: JSON.parse(text) };
  } catch {
    return null;
  }
}

/** 从第一个 { 开始找到与之配对的 }；字符串里的括号不算。找不到（被截断）时返回 null。 */
function firstBalancedObject(text: string): { json: string; rest: string } | null {
  const start = text.indexOf("{");
  if (start < 0) return null;
  let depth = 0;
  let quote = "";
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (quote) {
      if (ch === "\\") i++;
      else if (ch === quote) quote = "";
      continue;
    }
    if (ch === '"' || ch === "'") quote = ch;
    else if (ch === "{" || ch === "[") depth++;
    else if (ch === "}" || ch === "]") {
      depth--;
      if (depth === 0) return { json: text.slice(start, i + 1), rest: text.slice(i + 1) };
    }
  }
  return null;
}

const BARE_WORDS: Record<string, string> = { True: "true", False: "false", None: "null", NaN: "null", undefined: "null" };

/** 把「宽松 JSON」（单引号、裸键名、尾逗号、Python 字面量、字符串内原始换行）转成标准 JSON。 */
function relaxedToStrict(src: string): string {
  let out = "";
  let lastSignificant = "";
  for (let i = 0; i < src.length; ) {
    const ch = src[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      let body = "";
      while (j < src.length && src[j] !== ch) {
        const c = src[j];
        if (c === "\\" && j + 1 < src.length) {
          const next = src[j + 1];
          body += ch === "'" && next === "'" ? "'" : c + next;
          j += 2;
          continue;
        }
        if (c === '"') body += '\\"';
        else if (c === "\n") body += "\\n";
        else if (c === "\r") body += "\\r";
        else if (c === "\t") body += "\\t";
        else body += c;
        j++;
      }
      if (j >= src.length) return src; // 字符串没闭合：交给调用方按失败处理
      out += `"${body}"`;
      lastSignificant = '"';
      i = j + 1;
      continue;
    }
    if (/[A-Za-z_$]/.test(ch)) {
      let j = i;
      while (j < src.length && /[\w$]/.test(src[j])) j++;
      const word = src.slice(i, j);
      let k = j;
      while (k < src.length && /\s/.test(src[k])) k++;
      if ((lastSignificant === "{" || lastSignificant === ",") && src[k] === ":") {
        out += JSON.stringify(word);
        lastSignificant = '"';
      } else {
        out += BARE_WORDS[word] ?? word;
        lastSignificant = "w";
      }
      i = j;
      continue;
    }
    if (ch === ",") {
      let k = i + 1;
      while (k < src.length && /\s/.test(src[k])) k++;
      if (src[k] === "}" || src[k] === "]") {
        i++;
        continue;
      }
    }
    if (!/\s/.test(ch)) lastSignificant = ch;
    out += ch;
    i++;
  }
  return out;
}

/** 解析模型给出的 arguments 字符串；能安全修好的格式问题直接修好。 */
export function parseToolArgumentsJson(text: string | null | undefined): ParsedToolArguments {
  const raw = (text ?? "").trim();
  if (!raw) return { ok: true, value: {}, repaired: false };
  const direct = tryParse(raw);
  if (direct) {
    if (isRecord(direct.value)) return { ok: true, value: direct.value, repaired: false };
    // 双重编码："{\"sheet\":\"S\"}"
    if (typeof direct.value === "string") {
      const inner = parseToolArgumentsJson(direct.value);
      if (inner.ok) return { ...inner, repaired: true };
    }
    return { ok: false, error: `工具参数必须是 JSON 对象，收到: ${raw.slice(0, 120)}` };
  }
  let candidate = raw;
  const fence = /^```[\w-]*\s*([\s\S]*?)\s*```$/.exec(candidate);
  if (fence) candidate = fence[1].trim();
  const balanced = firstBalancedObject(candidate);
  if (!balanced) return { ok: false, error: `工具参数不是完整的 JSON 对象（可能被截断）: ${raw.slice(0, 200)}` };
  if (balanced.rest.trim().startsWith("{")) {
    return { ok: false, error: "工具参数里有多个 JSON 对象。每个 tool_call 只能有一个参数对象；多次操作请拆成多个 tool_calls。" };
  }
  for (const attempt of [balanced.json, relaxedToStrict(balanced.json)]) {
    const parsed = tryParse(attempt);
    if (parsed && isRecord(parsed.value)) return { ok: true, value: parsed.value, repaired: true };
  }
  return { ok: false, error: `工具参数不是合法 JSON: ${raw.slice(0, 200)}` };
}

/** 模型常用的同义字段（已转成 snake_case）→ 候选的正式字段，按顺序取第一个工具里声明了、且还没填的。 */
const KEY_ALIASES: Record<string, string[]> = {
  sheet_name: ["sheet", "name"],
  sheetname: ["sheet", "name"],
  worksheet: ["sheet", "name"],
  worksheet_name: ["sheet", "name"],
  tab: ["sheet"],
  sheet: ["name"],
  address: ["range"],
  range_address: ["range"],
  cell_range: ["range"],
  area: ["range"],
  cell: ["start_cell", "range"],
  start: ["start_cell"],
  start_range: ["start_cell"],
  top_left: ["start_cell"],
  anchor: ["start_cell"],
  data: ["values"],
  rows: ["values"],
  value: ["values"],
  formula: ["formulas"],
  new_sheet_name: ["new_name"],
  rename_to: ["new_name"],
  target_sheet: ["dest_sheet"],
  destination_sheet: ["dest_sheet"],
  to_sheet: ["dest_sheet"],
  from_sheet: ["source_sheet"],
  src_sheet: ["source_sheet"],
  src_range: ["source_range"],
  from_range: ["source_range"],
  target_cell: ["dest_start_cell"],
  dest_cell: ["dest_start_cell"],
  destination_cell: ["dest_start_cell"],
  destination: ["destination_range", "dest_start_cell"],
  search: ["find", "text"],
  find_text: ["find", "text"],
  query: ["text"],
  replace_with: ["replace"],
  replacement: ["replace"],
  column_name: ["column", "key_column"],
  col: ["column"],
  sort_by: ["key_column"],
  sort_column: ["key_column"],
  include_formula: ["include_formulas"],
  with_formulas: ["include_formulas"],
  limit: ["max_rows"],
};

function toSnakeCase(key: string): string {
  return key.replace(/([a-z0-9])([A-Z])/g, "$1_$2").replace(/[\s-]+/g, "_").toLowerCase();
}

function renameAliasedKeys(record: Record<string, unknown>, properties: Record<string, unknown>): void {
  const declared = (key: string) => Object.prototype.hasOwnProperty.call(properties, key);
  for (const key of Object.keys(record)) {
    if (declared(key)) continue;
    const snake = toSnakeCase(key);
    const target = [snake, ...(KEY_ALIASES[snake] ?? [])].find((c) => c !== key && declared(c) && record[c] === undefined);
    if (!target) continue;
    record[target] = record[key];
    delete record[key];
  }
}

const acceptsScalar = (schema: Record<string, any> | undefined) =>
  !schema || (schema.type !== "object" && schema.type !== "array");

/** 按声明递归校验工具参数，兼容模型把布尔值和数字写成字符串。 */
export function normalizeToolArguments(value: unknown, schema: Record<string, any>, path = "参数"): any {
  if (Array.isArray(schema.anyOf)) {
    for (const alternative of schema.anyOf) {
      try { return normalizeToolArguments(value, alternative, path); } catch { /* 尝试下一个声明的形状 */ }
    }
    throw new Error(`${path}不符合工具要求的参数类型或字段，请对照工具定义重试`);
  }
  let normalized = value;
  if (typeof normalized === "string") {
    const text = normalized.trim();
    if (schema.type === "boolean" && /^(true|false)$/i.test(text)) normalized = text.toLowerCase() === "true";
    if ((schema.type === "number" || schema.type === "integer") && text && Number.isFinite(Number(text))) {
      normalized = Number(text);
    }
    // 数组/对象被写成了 JSON 字符串
    if ((schema.type === "array" && text.startsWith("[")) || (schema.type === "object" && text.startsWith("{"))) {
      const parsed = tryParse(text);
      if (parsed && (schema.type === "array" ? Array.isArray(parsed.value) : isRecord(parsed.value))) normalized = parsed.value;
    }
  }
  if (schema.type === "string" && typeof normalized === "number" && Number.isFinite(normalized)) normalized = String(normalized);
  const fail = (detail: string): never => { throw new Error(`${path}${detail}`); };
  if (schema.type === "object") {
    if (!normalized || typeof normalized !== "object" || Array.isArray(normalized)) fail("必须是 JSON 对象");
    const record = { ...(normalized as Record<string, unknown>) };
    const properties: Record<string, unknown> = schema.properties || {};
    renameAliasedKeys(record, properties);
    const required: string[] = schema.required || [];
    // 可选字段填 null 等同于没填
    for (const [key, child] of Object.entries(properties)) {
      if (record[key] === null && !required.includes(key) && (child as Record<string, any>)?.type !== "null") delete record[key];
    }
    if (schema.additionalProperties === false) {
      const unknown = Object.keys(record).filter((key) => !Object.prototype.hasOwnProperty.call(properties, key));
      if (unknown.length) fail(`包含未声明字段 ${unknown.join("、")}；允许字段：${Object.keys(properties).join("、") || "无"}`);
    }
    for (const key of required) {
      if (record[key] == null || (typeof record[key] === "string" && !(record[key] as string).trim())) fail(`缺少必填字段 ${key}`);
    }
    for (const [key, child] of Object.entries(properties)) {
      if (record[key] !== undefined) record[key] = normalizeToolArguments(record[key], child as Record<string, any>, `${path}.${key}`);
    }
    normalized = record;
  } else if (schema.type === "array") {
    const items: Record<string, any> | undefined = schema.items;
    if (!Array.isArray(normalized)) {
      const scalar = normalized != null && typeof normalized !== "object";
      const looksLikeJson = typeof normalized === "string" && /^\s*[[{]/.test(normalized);
      // 单个值没包成数组：[值]；二维参数（如 values）包成 [[值]]
      if (scalar && !looksLikeJson && items?.type === "array" && acceptsScalar(items.items)) normalized = [[normalized]];
      else if (scalar && !looksLikeJson && acceptsScalar(items)) normalized = [normalized];
      else fail("必须是数组");
    }
    const list = normalized as unknown[];
    if (items?.type === "array" && list.length && list.every((item) => item === null || typeof item !== "object")) {
      fail("必须是二维数组（行的数组），例如 [[\"名称\",1],[\"苹果\",3]]；单列数据写成 [[1],[2],[3]]");
    }
    if (schema.minItems != null && list.length < schema.minItems) fail(`至少需要 ${schema.minItems} 项`);
    if (schema.maxItems != null && list.length > schema.maxItems) fail(`最多允许 ${schema.maxItems} 项`);
    normalized = list.map((item, i) => normalizeToolArguments(item, items || {}, `${path}[${i}]`));
  } else if (schema.type === "string" && typeof normalized !== "string") fail("必须是字符串");
  else if (schema.type === "boolean" && typeof normalized !== "boolean") fail("必须是 true 或 false");
  else if (schema.type === "null" && normalized !== null) fail("必须是 null");
  else if (schema.type === "number" || schema.type === "integer") {
    if (typeof normalized !== "number" || !Number.isFinite(normalized)) fail("必须是有限数字");
    if (schema.type === "integer" && !Number.isInteger(normalized)) fail("必须是整数");
  }
  if (schema.enum && !schema.enum.includes(normalized)) {
    // 枚举只差大小写 / 首尾空格时按声明值纠正（center → Center）
    const text = typeof normalized === "string" ? normalized.trim().toLowerCase() : null;
    const match = text === null ? undefined : schema.enum.find((option: unknown) => typeof option === "string" && option.toLowerCase() === text);
    if (match === undefined) fail(`必须为 ${schema.enum.join(" / ")}`);
    normalized = match;
  }
  if (typeof normalized === "number") {
    if (schema.minimum != null && normalized < schema.minimum) fail(`不能小于 ${schema.minimum}`);
    if (schema.maximum != null && normalized > schema.maximum) fail(`不能大于 ${schema.maximum}`);
  }
  if (typeof normalized === "string") {
    if (schema.minLength != null && normalized.length < schema.minLength) fail(`至少需要 ${schema.minLength} 个字符`);
    if (schema.maxLength != null && normalized.length > schema.maxLength) fail(`最多允许 ${schema.maxLength} 个字符`);
  }
  return normalized;
}
