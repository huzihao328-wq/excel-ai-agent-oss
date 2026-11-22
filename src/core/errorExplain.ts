/**
 * 把模型请求的失败翻译成用户能照着做的一句话，并指出下一步该去哪：
 * 界面上只留结论和按钮，服务商的原话作为次要信息附在后面，不再把整段 JSON 甩给用户。
 * 纯函数，不依赖网络和界面。
 */

/** 报错卡片上除「重试」之外的出路 */
export type ErrorAction = "settings" | "new-chat";

export interface ExplainedError {
  /** 给用户看的结论 */
  text: string;
  /** 服务商原话（已从 JSON 里取出），可能为空 */
  detail?: string;
  action?: ErrorAction;
}

export interface ExplainContext {
  /** 当前 API 地址是否指向本机（Ollama / LM Studio） */
  loopback?: boolean;
}

const NETWORK_RE = /Failed to fetch|NetworkError|network error|Load failed|fetch failed|ERR_NETWORK|ERR_CONNECTION|ECONNRESET|ECONNREFUSED|ETIMEDOUT|ENOTFOUND/i;
const OVERFLOW_RE = /context[_ ]length|maximum context|context window|too many tokens|token limit|prompt is too long|input is too long/i;
const BALANCE_RE = /insufficient|balance|quota|billing|arrears|payment required|credit|余额|欠费|额度/i;
const AUTH_RE = /(invalid|incorrect|wrong|expired|missing)[^.\n]{0,24}(api[ _-]?key|token|credential)|unauthori[sz]ed|authentication|api[ _-]?key[^.\n]{0,16}(invalid|incorrect|expired|not found)|无效[^。\n]{0,8}(key|密钥|令牌)|(key|密钥|令牌)[^。\n]{0,8}(无效|错误|过期)/i;
const TOOLS_RE = /(tool|function)[^.\n]{0,60}(not support|unsupported|does not support|unknown parameter|unrecognized)|(not support|unsupported|does not support)[^.\n]{0,60}(tool|function)/i;

/** 从接口返回的正文里取出一句人能读的原因；整页 HTML、空正文返回空串 */
export function providerMessage(raw: string): string {
  const body = raw.trim();
  if (!body) return "";
  if (/^<(!doctype|html)/i.test(body)) return "";
  let text = body;
  try {
    const parsed = JSON.parse(body);
    const pick = (v: unknown): string => {
      if (typeof v === "string") return v;
      if (v && typeof v === "object") {
        const o = v as Record<string, unknown>;
        return pick(o.message) || pick(o.msg) || pick(o.error) || pick(o.detail) || pick(o.error_msg);
      }
      return "";
    };
    text = pick(parsed) || body;
  } catch {
    /* 不是 JSON：按原文处理 */
  }
  text = text.replace(/\s+/g, " ").trim();
  return text.length > 220 ? text.slice(0, 220) + "…" : text;
}

function messageOf(error: unknown): string {
  if (error instanceof Error) return error.message;
  if (error && typeof error === "object" && typeof (error as { message?: unknown }).message === "string") {
    return (error as { message: string }).message;
  }
  return String(error ?? "");
}

export function explainRunError(error: unknown, ctx: ExplainContext = {}): ExplainedError {
  const message = messageOf(error).trim();
  const fields = (error && typeof error === "object" ? error : {}) as { status?: unknown; contextOverflow?: unknown };

  // 「模型请求失败 (HTTP 401): {…}\n\n提示：…」「HTTP 401: {…}」「获取模型列表失败 (HTTP 401): …」
  const http = /HTTP (\d{3})\)?[:：]\s*([\s\S]*?)(?:\n\n提示：[\s\S]*)?$/.exec(message);
  const status = typeof fields.status === "number" ? fields.status : http ? Number(http[1]) : 0;

  if (!status) {
    if (NETWORK_RE.test(message)) {
      return ctx.loopback
        ? { text: "连不上本机的模型服务。请确认 Ollama / LM Studio 已启动，并允许跨域访问（Ollama 需以 OLLAMA_ORIGINS=* 启动）。", action: "settings" }
        : { text: "连不上模型服务。请检查网络和 API 地址是否正确；公司内网可能拦截了外部接口。", action: "settings" };
    }
    if (/上下文窗口|上下文预算/.test(message)) return { text: message, action: "new-chat" };
    if (/不支持图片|API 地址|必须使用 HTTPS/.test(message)) return { text: message, action: "settings" };
    return { text: message || "出错了，但没有拿到具体原因。请重试。" };
  }

  const said = providerMessage(http ? http[2] : "");
  const detail = said ? `服务商返回（HTTP ${status}）：${said}` : `HTTP ${status}`;
  const probe = `${said} ${message}`;

  if (fields.contextOverflow === true || status === 413 || OVERFLOW_RE.test(probe)) {
    return { text: "对话内容超出了模型的上下文窗口。请新建对话重新开始，或在设置里核对「上下文窗口」。", detail, action: "new-chat" };
  }
  if (status === 402 || (BALANCE_RE.test(said) && [400, 401, 403, 429].includes(status))) {
    return { text: "账户余额不足或额度已用完。请到服务商控制台充值，或换一个可用的 API Key。", detail, action: "settings" };
  }
  if (status === 401 || AUTH_RE.test(said)) {
    return { text: "API Key 无效或已过期。请在设置里重新填写。", detail, action: "settings" };
  }
  if (status === 403) {
    return { text: "当前 API Key 没有权限使用这个模型或接口。请换一个模型，或检查 Key 的权限。", detail, action: "settings" };
  }
  if (TOOLS_RE.test(said)) {
    return { text: "当前模型或接口不支持工具调用，没法读写表格。请换用支持工具调用的模型。", detail, action: "settings" };
  }
  if (status === 404) {
    return { text: "模型名不存在，或 API 地址不对。请在设置里核对后重试。", detail, action: "settings" };
  }
  if (status === 429) {
    return { text: "请求太频繁，被服务商限流了。已自动重试仍未成功，请稍等一会儿再试。", detail };
  }
  if (status === 408 || status >= 500) {
    return { text: "模型服务暂时不可用。已自动重试仍未成功，请稍后再试。", detail };
  }
  return { text: "模型接口拒绝了这次请求。", detail, action: "settings" };
}
