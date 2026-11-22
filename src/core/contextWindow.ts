/** 从 API 返回或模型名解析上下文窗口（token 数）。 */
import type { ChatMessage } from "./types";

const CONTEXT_KEY_SCORE: Array<{ re: RegExp; score: number }> = [
  { re: /context[_\s-]*(window|length|size)|max[_\s-]*model[_\s-]*len|max[_\s-]*context|n_ctx|num_ctx/i, score: 3 },
  { re: /max[_\s-]*input[_\s-]*(token|len)|max[_\s-]*position/i, score: 2 },
  { re: /max[_\s-]*seq/i, score: 1 },
];

export function parseTokenCount(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    const n = Math.round(value);
    return n >= 1024 && n <= 16_000_000 ? n : null;
  }
  if (typeof value !== "string") return null;
  const s = value.trim().toLowerCase().replace(/,/g, "").replace(/tokens?/g, "").trim();
  const m = /^(\d+(?:\.\d+)?)\s*([km])?$/.exec(s);
  if (!m) return null;
  let n = Number(m[1]);
  if (m[2] === "k") n *= 1024;
  if (m[2] === "m") n *= 1024 * 1024;
  n = Math.round(n);
  return n >= 1024 && n <= 16_000_000 ? n : null;
}

function keyScore(key: string): number {
  const lower = key.toLowerCase();
  if (/\.context_length$/.test(lower) || /context_length$/.test(lower)) return 3;
  for (const { re, score } of CONTEXT_KEY_SCORE) {
    if (re.test(key)) return score;
  }
  return 0;
}

/** 在一个模型对象（及其嵌套 meta / model_info）里找上下文窗口。 */
export function pickContextTokens(value: unknown, depth = 0): number | null {
  if (depth > 6 || value == null) return null;
  if (typeof value === "number" || typeof value === "string") return depth === 0 ? parseTokenCount(value) : null;
  // 列表可能包含其他模型，必须由调用方先匹配模型 ID。
  if (Array.isArray(value)) return null;
  if (typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  let bestScore = 0;
  let bestVal: number | null = null;
  for (const [key, raw] of Object.entries(record)) {
    const score = keyScore(key);
    if (score > 0) {
      const n = parseTokenCount(raw);
      if (n && score === 1 && n < 8192) continue;
      if (n && (score > bestScore || (score === bestScore && n < (bestVal || Infinity)))) {
        bestScore = score;
        bestVal = n;
      }
    }
    if (score === 0 && raw && typeof raw === "object" && !/output|completion|vision|audio|text_config|generation/i.test(key)) {
      const nested = pickContextTokens(raw, depth + 1);
      if (nested && bestScore < 2) {
        bestScore = 2;
        bestVal = nested;
      } else if (nested && bestScore >= 2 && nested > (bestVal || 0) && depth < 2) {
        /* keep explicit field */
      }
    }
  }
  return bestVal;
}

export function formatContextWindow(tokens: number | null | undefined): string {
  if (!tokens || tokens < 1) return "";
  if (tokens >= 1_000_000 && tokens % 1_000_000 === 0) return `${tokens / 1_000_000}M`;
  if (tokens >= 1024 * 1024 && tokens % (1024 * 1024) === 0) return `${tokens / (1024 * 1024)}M`;
  if (tokens % 1024 === 0 && tokens >= 1024) return `${tokens / 1024}K`;
  if (tokens % 1000 === 0 && tokens >= 1000) return `${tokens / 1000}K`;
  if (tokens >= 10_000) return `${Math.round(tokens / 1000)}K`;
  return String(tokens);
}

export function parseContextWindowInput(raw: string): number {
  const n = parseTokenCount(raw.trim()) ?? parseTokenCount(Number(raw.trim()));
  return n || 0;
}

/** 接口没返回时按模型名估算（偏保守）。 */
export function guessContextWindow(model: string): number | null {
  const id = model.trim().toLowerCase();
  if (!id) return null;
  const explicit = /(?:^|[-_:/])(\d+(?:\.\d+)?[km])(?:$|[-_:/])/.exec(id);
  if (explicit) {
    const namedWindow = parseTokenCount(explicit[1]);
    if (namedWindow) return namedWindow;
  }
  const rules: Array<[RegExp, number]> = [
    // 模型系列名不能证明网关实际开放了百万窗口；这里只给保守估算。
    [/gpt-4\.1|gpt-4\.5|gpt-5/, 128_000],
    [/gpt-4o|chatgpt-4o|o1-|o3-|o4-|^o[134](?:$|-)/, 128_000],
    [/gpt-oss/, 131_072],
    [/gpt-4-turbo|gpt-4-1106|gpt-4-0125/, 128_000],
    [/gpt-3\.5-turbo-16k/, 16_384],
    [/gpt-3\.5/, 16_384],
    [/claude/, 200_000],
    [/gemini-1\.5|gemini-2|gemini-3|gemini-flash|gemini-pro/, 1_000_000],
    [/gemini/, 32_768],
    [/qwen-long|qwen2\.5-72b|qwen2\.5-32b/, 131_072],
    [/qwen3|qwen-plus|qwen-max|qwen-turbo|qwen-flash|qwen2\.5|qwq/, 131_072],
    [/qwen/, 32_768],
    // DeepSeek 官方接口自 V3.1 起为 128K；更早的部署超限时会从报错里学到真实窗口。
    [/deepseek/, 131_072],
    [/moonshot|kimi/, 128_000],
    [/glm-4|glm4|glm-5|glm-z1/, 128_000],
    [/doubao/, 128_000],
    [/minimax|abab/, 128_000],
    [/llama-?4|llama4/, 131_072],
    [/llama-?3\.[13]|llama3\.[13]/, 131_072],
    [/llama-?3\.2|llama3\.2/, 131_072],
    [/mistral-large|mistral-medium|mistral-small|mixtral|codestral|devstral|magistral/, 128_000],
    [/grok/, 131_072],
    [/yi-|yi:/, 32_768],
  ];
  for (const [re, n] of rules) {
    if (re.test(id)) return n;
  }
  return null;
}

/* ---------------- 请求前的历史压缩 ---------------- */

/** 粗略估算 token：CJK 等非 ASCII 约 1 字 1 token，ASCII 约 3.5 字符 1 token。偏保守。 */
export function estimateTokens(text: string): number {
  if (!text) return 0;
  let ascii = 0;
  let other = 0;
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) < 128) ascii++;
    else other++;
  }
  return Math.ceil(ascii / 3.5 + other * 1.1);
}

function messageTokens(m: ChatMessage): number {
  let n = 8;
  if (typeof m.content === "string") n += estimateTokens(m.content);
  else if (Array.isArray(m.content)) {
    for (const part of m.content) {
      if (part.type === "text") n += estimateTokens(part.text);
      else n += 1200; // 一张图按固定成本估
    }
  }
  for (const tc of m.tool_calls || []) n += estimateTokens(tc.function.name) + estimateTokens(tc.function.arguments) + 12;
  return n;
}

export function estimateMessagesTokens(messages: ChatMessage[]): number {
  let n = 0;
  for (const m of messages) n += messageTokens(m);
  return n;
}

const STUB_HINT = "（较早的工具结果已省略；需要时重新调用工具获取最新数据）";

function stubToolResult(m: ChatMessage): ChatMessage {
  const text = typeof m.content === "string" ? m.content : "";
  let ok: boolean | undefined;
  try {
    const parsed = JSON.parse(text) as { ok?: unknown };
    if (typeof parsed?.ok === "boolean") ok = parsed.ok;
  } catch {
    /* 非 JSON 结果 */
  }
  const head = text.replace(/\s+/g, " ").slice(0, 160);
  return {
    role: "tool",
    tool_call_id: m.tool_call_id,
    content: JSON.stringify({ ...(ok === undefined ? {} : { ok }), omitted: true, hint: STUB_HINT, preview: head }),
  };
}

/**
 * 旧的 tool_calls 参数里常有整块写入数据（write_range 的 values 等）。
 * 折叠成形状说明并保持合法 JSON：有的服务商会解析历史里的参数，坏 JSON 会让整次请求 400。
 */
function compactCallArguments(args: string, maxChars: number): string {
  if (args.length <= maxChars) return args;
  const omitted = JSON.stringify({ omitted: "参数过长，已从历史中省略" });
  let parsed: unknown;
  try {
    parsed = JSON.parse(args);
  } catch {
    return omitted;
  }
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) return omitted;
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(parsed as Record<string, unknown>)) {
    const text = JSON.stringify(value) ?? "";
    if (text.length <= 160) out[key] = value;
    else if (Array.isArray(value)) {
      const cols = Array.isArray(value[0]) ? (value[0] as unknown[]).length : 0;
      out[key] = cols ? `<已省略 ${value.length}×${cols} 的数据>` : `<已省略 ${value.length} 项>`;
    } else if (typeof value === "string") out[key] = `${value.slice(0, 120)}…<已省略 ${value.length - 120} 字>`;
    else out[key] = "<已省略>";
  }
  return JSON.stringify(out);
}

function compactOldMessage(m: ChatMessage, longChars: number): ChatMessage {
  if (m.role === "tool") {
    const text = typeof m.content === "string" ? m.content : "";
    return text.length > longChars ? stubToolResult(m) : m;
  }
  if (m.role !== "assistant" || !m.tool_calls?.length) return m;
  let changed = false;
  const tool_calls = m.tool_calls.map((tc) => {
    const args = compactCallArguments(tc.function.arguments, longChars);
    if (args === tc.function.arguments) return tc;
    changed = true;
    return { ...tc, function: { ...tc.function, arguments: args } };
  });
  return changed ? { ...m, tool_calls } : m;
}

/** 找到当前任务的用户指令；执行中的补充不能让原任务被当作旧对话丢掉。 */
function lastUserIndex(history: ChatMessage[]): number {
  let latest = -1;
  for (let i = history.length - 1; i >= 0; i--) {
    const message = history[i];
    if (message.role !== "user") continue;
    if (latest < 0) latest = i;
    if (typeof message.content === "string" && message.content.startsWith("[任务进行中用户补充]")) continue;
    return i;
  }
  return latest;
}

export interface CompactOptions {
  /** 整个请求可用的 token 预算（已扣除 system prompt、工具定义与输出预留） */
  budgetTokens: number;
  /** 超过这个字符数的旧工具结果会被折叠成短摘要 */
  longToolResultChars?: number;
  /** 至少保留最近多少条消息不动 */
  keepRecent?: number;
}

/**
 * 生成用于发请求的历史副本（不改原数组）：
 * 1. 早于当前任务的长工具结果折叠成摘要；
 * 2. 仍超预算时按「user 轮」整块丢弃最旧的对话，并插一句说明；
 * 3. 保证 assistant.tool_calls 与 tool 结果成对出现，不产生悬空引用。
 */
export function compactHistoryForRequest(history: ChatMessage[], opts: CompactOptions): ChatMessage[] {
  const budget = Math.max(0, opts.budgetTokens);
  const longChars = opts.longToolResultChars ?? 1500;
  const keepRecent = opts.keepRecent ?? 6;
  if (estimateMessagesTokens(history) <= budget) return history;

  const turnStart = lastUserIndex(history);
  const protectFrom = Math.max(0, Math.min(turnStart < 0 ? history.length : turnStart, history.length - keepRecent));

  // 第一步：折叠旧工具结果和旧调用里的大块参数
  let out: ChatMessage[] = history.map((m, i) => (i >= protectFrom ? m : compactOldMessage(m, longChars)));
  if (estimateMessagesTokens(out) <= budget) return out;

  // 第二步：当前任务内部的旧长结果也折叠（保留最近 keepRecent 条）
  out = out.map((m, i) => (i >= history.length - keepRecent ? m : compactOldMessage(m, longChars)));
  if (estimateMessagesTokens(out) <= budget) return out;

  // 第三步：从最旧的 user 轮开始整块丢弃，直到进预算或只剩当前任务
  const userIdx: number[] = [];
  out.forEach((m, i) => {
    if (m.role === "user") userIdx.push(i);
  });
  let dropUntil = 0;
  for (let k = 1; k < userIdx.length; k++) {
    const candidate = userIdx[k];
    // keepRecent 只保护结果细节，不能让一个很大的旧任务永远无法整轮移除。
    if (candidate > turnStart) break;
    const rest = out.slice(candidate);
    dropUntil = candidate;
    if (estimateMessagesTokens(rest) + 40 <= budget) break;
  }
  if (dropUntil > 0) {
    const dropped = out.slice(0, dropUntil).filter((m) => m.role === "user").length;
    out = [
      { role: "user", content: `[系统提示] 更早的 ${dropped} 轮对话因上下文过长已省略，如需其中信息请重新查看表格。` },
      { role: "assistant", content: "好的。" },
      ...out.slice(dropUntil),
    ];
  }
  if (estimateMessagesTokens(out) <= budget) return out;

  // 第四步：仍超预算，把当前任务里除「最近一组」以外的长工具结果也折叠，只保留模型正要处理的那批
  let lastCallIdx = -1;
  for (let i = out.length - 1; i >= 0; i--) {
    if (out[i].role === "assistant" && out[i].tool_calls?.length) {
      lastCallIdx = i;
      break;
    }
  }
  if (lastCallIdx > 0) {
    out = out.map((m, i) => (i >= lastCallIdx ? m : compactOldMessage(m, longChars)));
  }
  return out;
}

export interface ContextWindowInfo {
  tokens: number;
  source: "catalog" | "detail" | "ollama-running" | "ollama-config" | "estimate" | "error";
}

export function describeContextWindow(info: ContextWindowInfo): string {
  const source = {
    catalog: "接口返回", detail: "模型元数据", "ollama-running": "Ollama 当前运行配置",
    "ollama-config": "Ollama num_ctx 配置", estimate: "按模型名估算，未验证", error: "接口报错确认",
  }[info.source];
  return `${formatContextWindow(info.tokens)}（${source}）`;
}

const cache = new Map<string, { info: ContextWindowInfo; expires: number }>();

function cacheKey(baseUrl: string, model: string, apiKey = ""): string {
  // 仅进程内缓存；不同凭证可能对应不同网关配额，不能串用。
  return JSON.stringify([baseUrl.trim().replace(/\/+$/, ""), model.trim(), apiKey.trim()]);
}

export function getCachedContextInfo(baseUrl: string, model: string, apiKey = ""): ContextWindowInfo | undefined {
  if (!model.trim()) return undefined;
  const key = cacheKey(baseUrl, model, apiKey);
  const entry = cache.get(key);
  if (!entry || entry.expires <= Date.now()) { cache.delete(key); return undefined; }
  return entry.info;
}

export function getCachedContextWindow(baseUrl: string, model: string, apiKey = ""): number | undefined {
  return getCachedContextInfo(baseUrl, model, apiKey)?.tokens;
}

export function forgetContextWindow(baseUrl: string, model: string, apiKey = ""): void {
  cache.delete(cacheKey(baseUrl, model, apiKey));
}

export function rememberContextWindow(baseUrl: string, model: string, tokens: number, source: ContextWindowInfo["source"] = "catalog", apiKey = ""): ContextWindowInfo {
  const info = { tokens, source };
  if (model.trim() && parseTokenCount(tokens)) {
    if (cache.size >= 200) cache.delete(cache.keys().next().value!);
    cache.set(cacheKey(baseUrl, model, apiKey), { info, expires: Date.now() + (source === "estimate" ? 60_000 : 300_000) });
  }
  return info;
}

export function rememberContextWindows(baseUrl: string, items: Array<{ id: string; contextWindow?: number }>, apiKey = ""): void {
  for (const item of items) {
    const prev = getCachedContextInfo(baseUrl, item.id, apiKey);
    if (item.contextWindow && (!prev || prev.source === "catalog" || prev.source === "estimate")) {
      rememberContextWindow(baseUrl, item.id, item.contextWindow, "catalog", apiKey);
    }
  }
}

/** 只接受明确说出的上下文总上限，不能把请求 token 数、输出上限或 413 体积当成窗口。 */
export function contextLimitFromError(detail: string): number | null {
  const match = /(?:maximum context length|context (?:window|length)(?: limit)?|max(?:imum)?_model_len)\s*(?:is|of|:|=)?\s*([\d,]+)\s*(?:tokens|\b)/i.exec(detail);
  return match ? parseTokenCount(match[1]) : null;
}
