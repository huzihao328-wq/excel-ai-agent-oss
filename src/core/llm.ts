import { AppSettings, ChatMessage, ToolCall, ToolDefinition } from "./types";
import { extractModelCatalog } from "./modelList";
import { isLoopbackEndpoint, modelEndpointError } from "./settings";
import { withLlmRetries } from "./llmRetry";
import { readChatResponse } from "./chatResponse";
import { extractTextToolCalls } from "./toolCallText";
import { adaptToolsForEndpoint } from "./toolSchema";
import {
  effectiveMaxOutputTokens,
  getOutputLimit,
  markStreamUsage,
  outputLimitFromError,
  rememberOutputLimit,
  shouldRequestStreamUsage,
  streamUsageConfirmed,
  TokenUsage,
} from "./modelCapabilities";
import {
  ContextWindowInfo,
  describeContextWindow,
  getCachedContextInfo,
  guessContextWindow,
  pickContextTokens,
  rememberContextWindow,
  rememberContextWindows,
  parseTokenCount,
  contextLimitFromError,
  forgetContextWindow,
} from "./contextWindow";

export class ModelRequestError extends Error {
  constructor(
    message: string,
    public status: number,
    public contextLimit: number | null,
    public contextOverflow: boolean,
    /** 接口报告的单次输出上限（max_tokens 超限时） */
    public outputLimit: number | null = null
  ) {
    super(message);
    this.name = "ModelRequestError";
  }
}

export interface StreamCallbacks {
  onTextDelta?: (delta: string, fullText: string) => void;
  onAttempt?: (attempt: number, maxAttempts: number) => void;
  onRetryWait?: (attempt: number, maxAttempts: number, waitMs: number) => void;
}

export interface AssistantTurn {
  content: string;
  toolCalls: ToolCall[];
  /** 服务端给出的结束原因（stop / tool_calls / length …），拿不到时为空 */
  finishReason: string;
  /** 输出被 max_tokens 截断 */
  truncated: boolean;
  /** 工具调用是从正文里解析出来的（模型没走原生 tool_calls） */
  toolCallsFromText: boolean;
  /** 接口返回的 token 用量（有些接口不返回） */
  usage?: TokenUsage;
}

/**
 * 组装请求体：只放各家 OpenAI 兼容接口都接受的字段，避免多余参数被某些网关 400。
 * stream_options 是唯一例外：它让接口在流末尾报告真实 token 用量，被拒绝时调用方会去掉它重发并记住。
 */
export function buildChatRequestBody(
  settings: Pick<AppSettings, "model" | "temperature" | "maxOutputTokens">,
  messages: ChatMessage[],
  tools: ToolDefinition[],
  stream: boolean,
  extras: { includeUsage?: boolean } = {}
): Record<string, unknown> {
  const body: Record<string, unknown> = {
    model: settings.model,
    temperature: settings.temperature,
    messages,
    stream,
  };
  if (tools.length) body.tools = tools;
  if (settings.maxOutputTokens > 0) body.max_tokens = Math.round(settings.maxOutputTokens);
  if (stream && extras.includeUsage) body.stream_options = { include_usage: true };
  return body;
}

function requestHeaders(apiKey: string): Record<string, string> {
  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (apiKey.trim()) headers.Authorization = `Bearer ${apiKey.trim()}`;
  return headers;
}

/** 把服务商晦涩的 4xx 报错翻译成用户能处理的建议；认不出来就返回空串。 */
export function classifyRequestError(status: number, detail: string, sentTools: boolean): string {
  const d = detail.toLowerCase();
  if (status === 413 || /context[_ ]length|maximum context|too many tokens|token limit|exceeds? the (maximum|limit)|input is too long|prompt is too long|max_tokens.*(exceed|greater)/i.test(d)) {
    return "提示：对话上下文可能已超过模型窗口。请新建对话，或在设置里填写正确的「上下文窗口」以便自动压缩历史。";
  }
  if (sentTools && (status === 400 || status === 404 || status === 422) && /tool|function/.test(d) && /not support|unsupported|does not support|invalid|unknown parameter|unrecognized/.test(d)) {
    return "提示：当前模型或接口可能不支持工具调用（function calling）。请换用支持工具调用的模型，如 deepseek-chat、qwen-plus、glm-4-air、gpt-4o。";
  }
  if (status === 404 && /model/.test(d)) {
    return "提示：模型名不存在或无权限。请在设置里点「刷新」从服务商的模型列表中选择。";
  }
  return "";
}

function modelUrl(baseUrl: string, path: string): string {
  const error = modelEndpointError(baseUrl);
  if (error) throw new Error(error);
  const base = baseUrl.trim().replace(/\/+$/, "");
  if (base === "https://opencode.ai/zen/go/v1") return "/api/opencode-go" + path;
  return base + path;
}

/**
 * 调用 OpenAI 兼容的 /chat/completions 接口（流式），
 * 返回完整的 assistant 回合（文本 + 工具调用）。
 * 对限流、5xx、断网会自动重试；已经开始往界面推文字后不再重试，避免重复输出。
 */
export async function streamChatCompletion(
  settings: AppSettings,
  messages: ChatMessage[],
  tools: ToolDefinition[],
  callbacks: StreamCallbacks,
  signal: AbortSignal
): Promise<AssistantTurn> {
  let streamed = false;
  const wrapped: StreamCallbacks = {
    onTextDelta: (delta, fullText) => {
      streamed = true;
      callbacks.onTextDelta?.(delta, fullText);
    },
  };
  return withLlmRetries(() => streamChatCompletionOnce(settings, messages, tools, wrapped, signal), {
    signal,
    alreadyStarted: () => streamed,
    onAttempt: callbacks.onAttempt,
    onRetryWait: callbacks.onRetryWait,
  });
}

async function streamChatCompletionOnce(
  settings: AppSettings,
  messages: ChatMessage[],
  tools: ToolDefinition[],
  callbacks: StreamCallbacks,
  signal: AbortSignal
): Promise<AssistantTurn> {
  const url = modelUrl(settings.baseUrl, "/chat/completions");
  const adaptedTools = adaptToolsForEndpoint(tools, settings.baseUrl);
  let includeUsage = shouldRequestStreamUsage(settings.baseUrl);
  let usageFallback = false;
  let outputLimitRetried = false;
  // 以下两种自愈重发都发生在 HTTP 层、模型还没输出任何内容，不会重复执行工具。
  for (;;) {
    const maxOutputTokens = effectiveMaxOutputTokens(settings);
    const resp = await fetch(url, {
      method: "POST",
      headers: requestHeaders(settings.apiKey),
      body: JSON.stringify(buildChatRequestBody({ ...settings, maxOutputTokens }, messages, adaptedTools, true, { includeUsage })),
      signal,
    });

    if (!resp.ok) {
      let detail = "";
      try {
        detail = await resp.text();
      } catch {
        /* ignore */
      }
      const error = requestFailure(resp.status, detail, settings, messages, tools.length > 0, maxOutputTokens);
      // max_tokens 超过模型上限：已记住上限，用上限立即重发一次。
      if (error instanceof ModelRequestError && error.outputLimit && !outputLimitRetried) {
        outputLimitRetried = true;
        continue;
      }
      // 带 stream_options 的请求被拒：去掉它重发一次；重发成功才记为该接口不支持。
      if (
        includeUsage && !streamUsageConfirmed(settings.baseUrl) && (resp.status === 400 || resp.status === 422) &&
        error instanceof ModelRequestError && !error.contextOverflow
      ) {
        includeUsage = false;
        usageFallback = true;
        if (/stream_options|include_usage/i.test(detail)) markStreamUsage(settings.baseUrl, false);
        continue;
      }
      throw error;
    }
    if (usageFallback) markStreamUsage(settings.baseUrl, false);
    // 接口接受了 stream_options（不管是否真的返回用量），之后的 4xx 就与它无关，不再回退重发。
    if (includeUsage) markStreamUsage(settings.baseUrl, true);
    const parsed = await readChatResponse(resp, (delta, full) => callbacks.onTextDelta?.(delta, full), signal);
    const { toolCalls, finishReason } = parsed;
    let visible = parsed.content;
    let toolCallsFromText = false;
    if (!toolCalls.length && visible && tools.length) {
      // 模型把工具调用写进了正文（<tool_call>、```json、裸 JSON 等），解析出来照常执行
      const extracted = extractTextToolCalls(visible, tools.map((t) => t.function.name));
      if (extracted.toolCalls.length) {
        toolCalls.push(...extracted.toolCalls);
        visible = extracted.content;
        toolCallsFromText = true;
      }
    }

    return {
      content: visible,
      toolCalls,
      finishReason,
      truncated: finishReason === "length" || finishReason === "max_tokens",
      toolCallsFromText,
      ...(parsed.usage ? { usage: parsed.usage } : {}),
    };
  }
}

/** 把失败响应翻译成带诊断信息的错误，并顺手记住从报错里学到的窗口 / 输出上限。 */
function requestFailure(
  status: number,
  detail: string,
  settings: AppSettings,
  messages: ChatMessage[],
  sentTools: boolean,
  sentMaxTokens: number
): Error {
  // 只在本次请求确实带了图片、且错误内容与视觉能力相关时，才提示换视觉模型，
  // 避免把「不支持工具调用」等其他错误误报成图片问题。
  const sentImages = messages.some(
    (m) => Array.isArray(m.content) && m.content.some((p) => p.type === "image_url")
  );
  if (sentImages && /image|vision|multimodal/i.test(detail)) {
    return new Error(
      "当前模型可能不支持图片。请在设置里换成支持视觉的模型（如 gpt-4o、qwen-vl-plus、gemini-1.5-pro）。"
    );
  }
  const hint = classifyRequestError(status, detail, sentTools);
  const contextLimit = contextLimitFromError(detail);
  const contextOverflow = (status === 400 || status === 413 || status === 422) &&
    /context[_ ]length|maximum context|context window|too many tokens|input is too long|prompt is too long/i.test(detail);
  if (contextLimit && contextOverflow) rememberContextWindow(settings.baseUrl, settings.model, contextLimit, "error", settings.apiKey);
  // 只有我们发出的 max_tokens 确实超过报告的上限时才算输出上限问题。
  const reportedOutput = status >= 400 && status < 500 && sentMaxTokens > 0 ? outputLimitFromError(detail) : null;
  const outputLimit = reportedOutput && reportedOutput < sentMaxTokens ? reportedOutput : null;
  if (outputLimit) rememberOutputLimit(settings.baseUrl, settings.model, outputLimit);
  return new ModelRequestError(
    `模型请求失败 (HTTP ${status}): ${detail.slice(0, 500)}${hint ? `\n\n${hint}` : ""}`,
    status, contextLimit, contextOverflow, outputLimit
  );
}

const MODELS_TIMEOUT_MS = 12000;
const modelsCache = new Map<string, string[]>();

function modelsCacheKey(settings: Pick<AppSettings, "baseUrl" | "apiKey">): string {
  return `${settings.baseUrl.trim()}\n${settings.apiKey.trim()}`;
}

export function getCachedModels(settings: Pick<AppSettings, "baseUrl" | "apiKey">): string[] {
  return modelsCache.get(modelsCacheKey(settings)) ?? [];
}

function rememberModels(settings: Pick<AppSettings, "baseUrl" | "apiKey">, models: string[]): void {
  modelsCache.set(modelsCacheKey(settings), models);
}

/** 获取服务商的可用模型列表（GET /models） */
export async function listModels(settings: Pick<AppSettings, "baseUrl" | "apiKey">, signal?: AbortSignal): Promise<string[]> {
  if (signal?.aborted) throw new DOMException("任务已停止", "AbortError");
  const url = modelUrl(settings.baseUrl, "/models");
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), MODELS_TIMEOUT_MS);
  const abort = () => ctrl.abort();
  signal?.addEventListener("abort", abort, { once: true });
  try {
    const resp = await fetch(url, {
      headers: {
        Accept: "application/json",
        ...(settings.apiKey.trim() ? { Authorization: `Bearer ${settings.apiKey.trim()}` } : {}),
      },
      signal: ctrl.signal,
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      throw new Error(`获取模型列表失败 (HTTP ${resp.status}): ${text.slice(0, 200)}`);
    }
    const json = await resp.json();
    const catalog = extractModelCatalog(json);
    rememberContextWindows(settings.baseUrl, catalog, settings.apiKey);
    for (const item of catalog) {
      if (item.maxOutputTokens && !getOutputLimit(settings.baseUrl, item.id)) rememberOutputLimit(settings.baseUrl, item.id, item.maxOutputTokens);
    }
    const ids = catalog.map((item) => item.id);
    rememberModels(settings, ids);
    return ids;
  } catch (e: any) {
    if (signal?.aborted) throw new DOMException("任务已停止", "AbortError");
    if (e?.name === "AbortError") throw new Error("获取模型列表超时。可手动输入模型名，或打开完整设置。");
    throw e;
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", abort);
  }
}

function ollamaOrigin(baseUrl: string): string | null {
  try {
    const url = new URL(baseUrl);
    if (!isLoopbackEndpoint(baseUrl) || url.port !== "11434") return null;
    return url.origin;
  } catch {
    return null;
  }
}

async function fetchJson(url: string, init: RequestInit, timeoutMs: number, signal?: AbortSignal): Promise<unknown | null> {
  if (signal?.aborted) throw new DOMException("任务已停止", "AbortError");
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  const abort = () => ctrl.abort();
  signal?.addEventListener("abort", abort, { once: true });
  try {
    const resp = await fetch(url, { ...init, signal: ctrl.signal });
    if (!resp.ok) return null;
    return await resp.json();
  } catch {
    if (signal?.aborted) throw new DOMException("任务已停止", "AbortError");
    return null;
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener("abort", abort);
  }
}

/**
 * 优先读取本地运行配置；其他接口按模型列表、详情、带来源的估算依次探测。
 */
export async function probeContextWindowInfo(
  settings: Pick<AppSettings, "baseUrl" | "apiKey" | "model">,
  options: { force?: boolean; signal?: AbortSignal } = {}
): Promise<ContextWindowInfo | null> {
  if (options.signal?.aborted) throw new DOMException("任务已停止", "AbortError");
  const model = settings.model.trim();
  if (!model) return null;
  // 在任何旁路探测之前验证地址，保持与正常聊天相同的协议和凭证限制。
  modelUrl(settings.baseUrl, "/models");
  const cached = getCachedContextInfo(settings.baseUrl, model, settings.apiKey);
  if (cached && !options.force) return cached;
  if (options.force) forgetContextWindow(settings.baseUrl, model, settings.apiKey);
  const remember = (tokens: number, source: ContextWindowInfo["source"]) =>
    rememberContextWindow(settings.baseUrl, model, tokens, source, settings.apiKey);
  const headers = { Accept: "application/json", ...requestHeaders(settings.apiKey) };

  // Ollama 的训练窗口与实际 num_ctx 不同，优先读取正在运行的配置。
  const ollama = ollamaOrigin(settings.baseUrl);
  if (ollama) {
    const running: any = await fetchJson(`${ollama}/api/ps`, { headers }, 4000, options.signal);
    const sameModel = (id: unknown) => typeof id === "string" &&
      (id === model || id === `${model}:latest` || `${id}:latest` === model);
    const item = Array.isArray(running?.models)
      ? running.models.find((m: any) => sameModel(m.name) || sameModel(m.model)) : undefined;
    const active = pickContextTokens(item);
    if (active) return remember(active, "ollama-running");
    const shown: any = await fetchJson(`${ollama}/api/show`, {
      method: "POST", headers, body: JSON.stringify({ model }),
    }, 4000, options.signal);
    const configured = typeof shown?.parameters === "string"
      ? parseTokenCount(/(?:^|\n)\s*num_ctx\s+(\d+)\b/.exec(shown.parameters)?.[1])
      : pickContextTokens(shown?.parameters);
    if (configured) return remember(configured, "ollama-config");
    // 仅有 model_info 中的训练上限无法证明实际运行窗口，继续查兼容接口。
  }

  try {
    await listModels(settings, options.signal);
    const fromList = getCachedContextInfo(settings.baseUrl, model, settings.apiKey);
    if (fromList?.source === "catalog") return fromList;
  } catch {
    if (options.signal?.aborted) throw new DOMException("任务已停止", "AbortError");
    /* 列表接口不可用时继续其他探测 */
  }

  try {
    const detail = await fetchJson(modelUrl(settings.baseUrl, `/models/${encodeURIComponent(model)}`), { headers }, 8000, options.signal);
    const n = pickContextTokens(detail);
    if (n) {
      return remember(n, "detail");
    }
  } catch {
    if (options.signal?.aborted) throw new DOMException("任务已停止", "AbortError");
    /* ignore */
  }

  // 本地模型的名字也不能推导 num_ctx；不拿训练上限冒充运行配置。
  const guessed = ollama ? null : guessContextWindow(model);
  if (guessed) {
    return remember(guessed, "estimate");
  }
  return null;
}

export async function probeContextWindow(settings: Pick<AppSettings, "baseUrl" | "apiKey" | "model">): Promise<number | null> {
  return (await probeContextWindowInfo(settings))?.tokens ?? null;
}

export interface ConnectionTestResult {
  message: string;
  contextWindow?: number;
  contextInfo?: ContextWindowInfo;
  toolCalling?: ToolProbeResult;
}

export interface ToolProbeResult {
  status: "native" | "text" | "unverified";
  message: string;
}

/** 仅调用虚拟工具，不访问工作簿。第二次请求必须能读取工具结果中的随机回执。 */
export async function probeToolCalling(settings: AppSettings): Promise<ToolProbeResult> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 30000);
  const tools: ToolDefinition[] = [{ type: "function", function: {
    name: "connection_probe", description: "连接诊断专用的虚拟工具，返回回执；不会读取或修改任何文件。",
    parameters: { type: "object", properties: { value: { type: "integer", enum: [7] } }, required: ["value"], additionalProperties: false },
  } }];
  const messages: ChatMessage[] = [{ role: "user", content: "请调用 connection_probe，value 填数字 7。收到工具结果后，只回复其中的 receipt。" }];
  const probeSettings = { ...settings, maxOutputTokens: 1024 };
  try {
    const first = await streamChatCompletionOnce(probeSettings, messages, tools, {}, ctrl.signal);
    const call = first.toolCalls[0];
    if (first.truncated || first.toolCalls.length !== 1 || call?.function.name !== "connection_probe") {
      return { status: "unverified", message: "工具调用未验证：模型未按要求返回完整调用" };
    }
    let args: any;
    try { args = JSON.parse(call.function.arguments); } catch { /* 下方统一报告 */ }
    if (!args || args.value !== 7 || Object.keys(args).length !== 1) {
      return { status: "unverified", message: "工具调用未验证：参数未符合工具定义" };
    }
    const receipt = `probe_${Math.random().toString(36).slice(2)}_${Date.now().toString(36)}`;
    messages.push({ role: "assistant", content: first.content || null, tool_calls: first.toolCalls });
    messages.push({ role: "tool", tool_call_id: call.id, content: JSON.stringify({ receipt }) });
    const second = await streamChatCompletionOnce(probeSettings, messages, tools, {}, ctrl.signal);
    if (second.truncated || second.toolCalls.length || second.content.trim() !== receipt) {
      return { status: "unverified", message: "工具调用未验证：模型未正确读取工具回执" };
    }
    return first.toolCallsFromText
      ? { status: "text", message: "文本兼容工具调用往返通过（非原生调用）" }
      : { status: "native", message: "原生工具调用、参数和结果回填测试通过" };
  } catch (e) {
    return { status: "unverified", message: ctrl.signal.aborted ? "工具调用测试超时，尚未验证" : `工具调用未验证：${String((e as Error)?.message || e).slice(0, 350)}` };
  } finally { clearTimeout(timer); }
}

export interface CompleteOnceOptions {
  maxOutputTokens?: number;
  temperature?: number;
  /** 默认 30 秒；后台任务（复盘、探测）不该无限期挂住 */
  timeoutMs?: number;
}

/**
 * 非流式的一次性补全，供连接测试、技能复盘等后台请求使用。
 * 与主对话共用同一套地址解析（含 OpenCode Go 代理）和鉴权头，避免各处自己拼 URL 再各自出错。
 * 返回 assistant 正文；纯推理模型正文为空时退回思考内容，避免被误判为“没有回复”。
 */
export async function completeOnce(
  settings: AppSettings,
  messages: ChatMessage[],
  options: CompleteOnceOptions = {}
): Promise<string> {
  const url = modelUrl(settings.baseUrl, "/chat/completions");
  const timeoutMs = options.timeoutMs ?? 30000;
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const resp = await fetch(url, {
      method: "POST",
      headers: requestHeaders(settings.apiKey),
      body: JSON.stringify(buildChatRequestBody(
        {
          model: settings.model,
          temperature: options.temperature ?? settings.temperature,
          maxOutputTokens: effectiveMaxOutputTokens({
            ...settings,
            maxOutputTokens: options.maxOutputTokens ?? settings.maxOutputTokens,
          }),
        },
        messages, [], false
      )),
      signal: ctrl.signal,
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      throw new Error(`HTTP ${resp.status}: ${text.slice(0, 300)}`);
    }
    const json = await resp.json();
    if (json?.error) {
      const detail = typeof json.error === "string" ? json.error : json.error?.message;
      throw new Error(`接口未返回有效的模型回复${detail ? `：${String(detail).slice(0, 300)}` : "，请检查 API 地址、模型和输出设置。"}`);
    }
    const message = json?.choices?.[0]?.message ?? {};
    const content = typeof message.content === "string" ? message.content : "";
    if (content.trim()) return content;
    const reasoning = message.reasoning_content ?? message.reasoning;
    if (typeof reasoning === "string" && reasoning.trim()) return reasoning;
    throw new Error("接口未返回有效的模型回复，请检查 API 地址、模型和输出设置。");
  } catch (e) {
    if (ctrl.signal.aborted) throw new Error(`请求超时（${Math.round(timeoutMs / 1000)} 秒），请检查 API 地址和网络后重试。`);
    throw e;
  } finally {
    clearTimeout(timer);
  }
}

/** 用一个最小请求测试连接配置是否可用，并顺带探测上下文窗口 */
export async function testConnection(settings: AppSettings): Promise<ConnectionTestResult> {
  // 推理模型会先把预算花在思考上，给太少的 max_tokens 会让正文为空而误报失败
  const reply = await completeOnce(
    settings,
    [{ role: "user", content: "ping，请回复 pong" }],
    { maxOutputTokens: 256, timeoutMs: 30000 }
  ).catch((e) => {
    throw new Error(String(e?.message || e).replace(/^请求超时/, "测试连接超时"));
  });
  const [contextInfo, toolCalling] = await Promise.all([
    probeContextWindowInfo(settings, { force: true }).catch(() => null), probeToolCalling(settings),
  ]);
  const windowHint = contextInfo ? `上下文 ${describeContextWindow(contextInfo)}` : "上下文窗口未探测到，可手动填写";
  return {
    message: `对话连接成功，模型回复: ${String(reply).slice(0, 50)}。${toolCalling.message}。${windowHint}。`,
    contextWindow: contextInfo?.tokens,
    contextInfo: contextInfo || undefined,
    toolCalling,
  };
}
