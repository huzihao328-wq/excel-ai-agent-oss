import type { ToolDefinition } from "./types";

/**
 * 工具参数 schema 的跨服务商适配。
 * 我们的定义里有「表头名 / 列字母 / 0 基列号」这类不限类型的参数（没有 type）以及 `items: {}` 的任意元素数组，
 * OpenAI / DeepSeek / Qwen 等都接受，但 Gemini 的 OpenAI 兼容层要求每个 schema 节点都有 type，否则整次请求 400。
 * 这里只对已知严格的服务商做改写，其余原样发送以免多花 token。
 */

const ANY_SCALAR = { anyOf: [{ type: "string" }, { type: "number" }, { type: "boolean" }] };

export function endpointNeedsStrictSchema(baseUrl: string): boolean {
  try {
    const host = new URL(baseUrl).hostname.toLowerCase();
    return host === "googleapis.com" || host.endsWith(".googleapis.com");
  } catch {
    return false;
  }
}

function hasTypeInfo(node: Record<string, unknown>): boolean {
  return "type" in node || "anyOf" in node || "oneOf" in node || "allOf" in node || "enum" in node || "$ref" in node;
}

function strictNode(node: unknown, isProperty: boolean): unknown {
  if (!node || typeof node !== "object" || Array.isArray(node)) return node;
  const rec = { ...(node as Record<string, unknown>) };
  for (const key of ["anyOf", "oneOf", "allOf"]) {
    if (Array.isArray(rec[key])) rec[key] = (rec[key] as unknown[]).map((child) => strictNode(child, true));
  }
  if (rec.properties && typeof rec.properties === "object") {
    const props: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(rec.properties as Record<string, unknown>)) props[k] = strictNode(v, true);
    rec.properties = props;
  }
  if ("items" in rec) {
    const items = rec.items;
    const empty = !items || (typeof items === "object" && Object.keys(items as object).length === 0);
    rec.items = empty ? { ...ANY_SCALAR } : strictNode(items, true);
    if (!("type" in rec)) rec.type = "array";
  }
  if (isProperty && !hasTypeInfo(rec)) {
    Object.assign(rec, ANY_SCALAR);
  }
  return rec;
}

/** 取第一句（到句号/分号为止），超过 limit 字则截断。 */
function firstSentence(text: string, limit: number): string {
  const t = text.trim();
  const end = t.search(/[。；;]|\.(\s|$)/);
  const sentence = end >= 0 ? t.slice(0, end + 1) : t;
  return sentence.length > limit ? `${sentence.slice(0, limit)}…` : sentence;
}

function compactSchemaNode(node: unknown): unknown {
  if (!node || typeof node !== "object" || Array.isArray(node)) return node;
  const rec = { ...(node as Record<string, unknown>) };
  if (typeof rec.description === "string") {
    // 参数说明只留第一个分句（如「表头名、列字母或 0 基列号」），示例和默认值说明去掉。
    const clause = rec.description.trim().split(/[，,。；;（(]/)[0].trim();
    if (clause) rec.description = clause.length > 24 ? `${clause.slice(0, 24)}…` : clause;
    else delete rec.description;
  }
  if (rec.properties && typeof rec.properties === "object") {
    rec.properties = Object.fromEntries(
      Object.entries(rec.properties as Record<string, unknown>).map(([k, v]) => [k, compactSchemaNode(v)])
    );
  }
  if (rec.items) rec.items = compactSchemaNode(rec.items);
  for (const key of ["anyOf", "oneOf", "allOf"]) {
    if (Array.isArray(rec[key])) rec[key] = (rec[key] as unknown[]).map(compactSchemaNode);
  }
  return rec;
}

/**
 * 小窗口模型用的精简工具定义：工具描述只留第一句、参数说明截短；名称、类型、枚举、必填都不变，
 * 所以参数校验和执行完全一样，只是少占上下文。
 */
export function compactToolDefinitions(tools: ToolDefinition[]): ToolDefinition[] {
  return tools.map((tool) => ({
    ...tool,
    function: {
      ...tool.function,
      description: firstSentence(tool.function.description, 60),
      parameters: compactSchemaNode(tool.function.parameters) as Record<string, unknown>,
    },
  }));
}

export function adaptToolsForEndpoint(tools: ToolDefinition[], baseUrl: string): ToolDefinition[] {
  if (!tools.length || !endpointNeedsStrictSchema(baseUrl)) return tools;
  return tools.map((tool) => ({
    ...tool,
    function: {
      ...tool.function,
      parameters: strictNode(tool.function.parameters, false) as Record<string, unknown>,
    },
  }));
}
