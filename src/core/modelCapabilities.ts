/**
 * 运行中学到的模型/接口能力：流式用量统计是否可用、单次输出上限、token 估算校准。
 * 只存在内存里，按接口地址（+ 模型）区分；换模型或重开窗格后重新学习，不会把一次误判固化下来。
 */
import type { AppSettings } from "./types";

const endpointKey = (baseUrl: string) => baseUrl.trim().replace(/\/+$/, "");
const modelKey = (baseUrl: string, model: string) => JSON.stringify([endpointKey(baseUrl), model.trim()]);

/* ---------------- 流式 usage（stream_options.include_usage） ---------------- */

const streamUsage = new Map<string, boolean>();

/** 没被接口明确拒绝过就请求用量统计；拒绝过的接口不再发这个字段。 */
export function shouldRequestStreamUsage(baseUrl: string): boolean {
  return streamUsage.get(endpointKey(baseUrl)) !== false;
}

/** 接口是否已经确认接受 stream_options（确认后不再为它做 400 回退重试）。 */
export function streamUsageConfirmed(baseUrl: string): boolean {
  return streamUsage.get(endpointKey(baseUrl)) === true;
}

export function markStreamUsage(baseUrl: string, supported: boolean): void {
  streamUsage.set(endpointKey(baseUrl), supported);
}

/* ---------------- 单次输出上限 ---------------- */

const outputLimits = new Map<string, number>();

function validOutputLimit(n: unknown): number | null {
  return typeof n === "number" && Number.isFinite(n) && n >= 256 && n <= 2_000_000 ? Math.round(n) : null;
}

export function rememberOutputLimit(baseUrl: string, model: string, tokens: number): void {
  const n = validOutputLimit(tokens);
  if (!n || !model.trim()) return;
  if (outputLimits.size >= 200) outputLimits.delete(outputLimits.keys().next().value!);
  outputLimits.set(modelKey(baseUrl, model), n);
}

export function getOutputLimit(baseUrl: string, model: string): number | undefined {
  return outputLimits.get(modelKey(baseUrl, model));
}

/** 实际发送的 max_tokens：用户设置值，但不超过接口告诉过我们的上限；0 表示不发送。 */
export function effectiveMaxOutputTokens(settings: Pick<AppSettings, "baseUrl" | "model" | "maxOutputTokens">): number {
  const requested = settings.maxOutputTokens > 0 ? Math.round(settings.maxOutputTokens) : 0;
  const limit = getOutputLimit(settings.baseUrl, settings.model);
  return requested && limit ? Math.min(requested, limit) : requested;
}

const OUTPUT_LIMIT_PATTERNS: RegExp[] = [
  // DeepSeek / 通义："the valid range of max_tokens is [1, 8192]"
  /max_(?:completion_|output_|new_)?tokens[^\n[]{0,60}\[\s*\d+\s*,\s*(\d[\d,]*)\s*\]/i,
  // OpenAI："This model supports at most 16384 completion tokens"
  /supports at most\s+(\d[\d,]*)\s+(?:completion|output)\s+tokens/i,
  // Anthropic 风格："64001 > 64000, which is the maximum allowed number of output tokens"
  /(\d[\d,]*),?\s+which is the maximum allowed number of output tokens/i,
  // "The maximum value of max_tokens is 8192"
  /maximum (?:value )?(?:of|for) max_(?:completion_|output_|new_)?tokens (?:is|=)\s*[`'"]?(\d[\d,]*)/i,
  // 通用："max_tokens must be less than or equal to 8192" / "<= 8192" / "at most 8192"
  /max_(?:completion_|output_|new_)?tokens[^\n]{0,80}?(?:less than or equal to|<=|≤|at most|not exceed|cannot exceed|no more than|no greater than|maximum(?: value)?(?: of| is)?)\s*[`'"]?(\d[\d,]*)/i,
];

/** 从 4xx 报错里读出「单次输出最多 N tokens」；上下文总长度的报错不算。 */
export function outputLimitFromError(detail: string): number | null {
  for (const re of OUTPUT_LIMIT_PATTERNS) {
    const match = re.exec(detail);
    if (!match) continue;
    const n = validOutputLimit(Number(match[1].replace(/,/g, "")));
    if (n) return n;
  }
  return null;
}

/* ---------------- token 估算校准 ---------------- */

interface Calibration {
  ratio: number;
  samples: number;
}

const calibrations = new Map<string, Calibration>();
const MIN_RATIO = 0.5;
const MAX_RATIO = 2;

/**
 * 用接口返回的 prompt_tokens 校准本地估算。偏大的观测权重更高：
 * 低估窗口占用会导致超限报错，高估只会多压缩一点历史。
 */
export function recordPromptUsage(baseUrl: string, model: string, estimatedTokens: number, actualTokens: number): void {
  if (!model.trim() || !(estimatedTokens >= 500) || !(actualTokens > 0)) return;
  const observed = Math.min(3, Math.max(0.3, actualTokens / estimatedTokens));
  const key = modelKey(baseUrl, model);
  const prev = calibrations.get(key);
  const ratio = !prev
    ? observed
    : observed > prev.ratio ? prev.ratio * 0.3 + observed * 0.7 : prev.ratio * 0.7 + observed * 0.3;
  if (calibrations.size >= 200 && !prev) calibrations.delete(calibrations.keys().next().value!);
  calibrations.set(key, { ratio, samples: (prev?.samples ?? 0) + 1 });
}

/** 实际 token ÷ 本地估算；还没有观测时为 1。 */
export function tokenEstimateRatio(baseUrl: string, model: string): number {
  const entry = calibrations.get(modelKey(baseUrl, model));
  if (!entry) return 1;
  return Math.min(MAX_RATIO, Math.max(MIN_RATIO, entry.ratio));
}

/** 接口在 usage 里报告的用量；字段名在各家之间不统一。 */
export interface TokenUsage {
  promptTokens?: number;
  completionTokens?: number;
  cachedTokens?: number;
}

export function parseUsage(raw: unknown): TokenUsage | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const u = raw as Record<string, any>;
  const num = (...values: unknown[]) => {
    for (const v of values) if (typeof v === "number" && Number.isFinite(v) && v >= 0) return v;
    return undefined;
  };
  const usage: TokenUsage = {
    promptTokens: num(u.prompt_tokens, u.input_tokens),
    completionTokens: num(u.completion_tokens, u.output_tokens),
    cachedTokens: num(u.prompt_tokens_details?.cached_tokens, u.prompt_cache_hit_tokens, u.cache_read_input_tokens),
  };
  return usage.promptTokens !== undefined || usage.completionTokens !== undefined ? usage : undefined;
}
