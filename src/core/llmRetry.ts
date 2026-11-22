/** 模型请求重试策略：纯函数，便于不启动网络也能核对。 */

export const LLM_MAX_ATTEMPTS = 3;

/** 只认浏览器/fetch 的整句中止提示；服务商 5xx 正文里偶然出现的 "aborted" 不算用户中止。 */
const ABORT_MESSAGE = /^\s*(aborted|The user aborted a request\.?|The operation was aborted\.?|signal is aborted without reason|This operation was aborted\.?)\s*$/i;

export function isAbortLike(error: unknown): boolean {
  if (error == null) return false;
  if (typeof error === "object") {
    const e = error as { name?: string; message?: string };
    if (e.name === "AbortError") return true;
    return ABORT_MESSAGE.test(e.message || "");
  }
  return ABORT_MESSAGE.test(String(error));
}

export function isRetryableLlmError(error: unknown): boolean {
  if (isAbortLike(error)) return false;
  const msg = error instanceof Error ? error.message : String(error);
  if (/当前模型可能不支持图片/.test(msg)) return false;
  if (/HTTP 40[014]|HTTP 403|HTTP 404|HTTP 413|HTTP 422/.test(msg)) return false;
  if (/模型请求失败 \(HTTP (408|429|500|502|503|504)\)/.test(msg)) return true;
  if (/HTTP (408|429|500|502|503|504)\b/.test(msg)) return true;
  if (/Failed to fetch|NetworkError|network error|ECONNRESET|ETIMEDOUT|ERR_NETWORK|Load failed|fetch failed|模型响应没有内容流/i.test(msg)) {
    return true;
  }
  if (/模型流式响应返回错误/.test(msg) && /(rate.?limit|too many requests|overloaded|busy|temporar|try again)/i.test(msg)) {
    return true;
  }
  return false;
}

/** failedAttempt 从 1 计：第一次失败等 800ms，之后 2000ms。 */
export function llmRetryDelayMs(failedAttempt: number): number {
  return failedAttempt <= 1 ? 800 : 2000;
}

export function sleepAbortable(ms: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal.aborted) {
      reject(abortFrom(signal));
      return;
    }
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, Math.max(0, ms));
    const onAbort = () => {
      clearTimeout(timer);
      reject(abortFrom(signal));
    };
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

function abortFrom(signal: AbortSignal): Error {
  const reason = (signal as AbortSignal & { reason?: unknown }).reason;
  if (reason instanceof Error) return reason;
  const err = new Error(reason != null ? String(reason) : "aborted");
  err.name = "AbortError";
  return err;
}

export interface LlmRetryHooks {
  signal: AbortSignal;
  maxAttempts?: number;
  delaysMs?: number[];
  alreadyStarted?: () => boolean;
  onAttempt?: (attempt: number, maxAttempts: number) => void;
  onRetryWait?: (attempt: number, maxAttempts: number, waitMs: number) => void;
}

export async function withLlmRetries<T>(run: () => Promise<T>, options: LlmRetryHooks): Promise<T> {
  const maxAttempts = options.maxAttempts ?? LLM_MAX_ATTEMPTS;
  let lastError: unknown;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    if (options.signal.aborted) throw abortFrom(options.signal);
    try {
      options.onAttempt?.(attempt, maxAttempts);
      return await run();
    } catch (e) {
      lastError = e;
      if (options.alreadyStarted?.()) throw e;
      if (isAbortLike(e) || options.signal.aborted) throw e;
      if (!isRetryableLlmError(e) || attempt >= maxAttempts) throw e;
      const waitMs = options.delaysMs?.[attempt - 1] ?? llmRetryDelayMs(attempt);
      options.onRetryWait?.(attempt, maxAttempts, waitMs);
      await sleepAbortable(waitMs, options.signal);
    }
  }
  throw lastError instanceof Error ? lastError : new Error(String(lastError));
}
