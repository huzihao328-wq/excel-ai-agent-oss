import type { ToolCall } from "./types";
import { appendStreamFragment, stripReasoningText, textFromContent } from "./streamText";
import { parseUsage, TokenUsage } from "./modelCapabilities";

/** 读取原生增量与 message 快照。JSON 参数必须逐字拼接，不能按重复文本去重。 */
export async function readChatResponse(
  response: Response,
  onText: (delta: string, fullText: string) => void,
  signal: AbortSignal
): Promise<{ content: string; toolCalls: ToolCall[]; finishReason: string; usage?: TokenUsage }> {
  let content = "";
  let finishReason = "";
  let usage: TokenUsage | undefined;
  let done = false;
  const calls = new Map<number, { id: string; name: string; args: string; lastArgs: string; snapshotArgs?: string }>();
  const isObjectJson = (text: string) => {
    try { const value = JSON.parse(text); return !!value && typeof value === "object" && !Array.isArray(value); }
    catch { return false; }
  };

  const consume = (json: any, snapshot = false) => {
    if (json?.error) {
      const error = typeof json.error === "string" ? json.error : json.error.message || JSON.stringify(json.error);
      throw new Error(`模型流式响应返回错误: ${String(error).slice(0, 500)}`);
    }
    // 用量可能在最后一个 choices 为空的块里，也可能每块都带（累计值），取最后一次。
    usage = parseUsage(json?.usage) ?? usage;
    const choice = json?.choices?.[0];
    if (!choice) return;
    if (choice.finish_reason) finishReason = String(choice.finish_reason);
    const fullMessage = snapshot || (!choice.delta && !!choice.message);
    const message = fullMessage ? choice.message : choice.delta;
    const piece = textFromContent(message?.content ?? choice.text);
    if (piece) {
      content = fullMessage ? piece : content + piece;
      const visible = stripReasoningText(content);
      if (visible) onText(piece, visible);
    }
    if (!Array.isArray(message?.tool_calls)) return;
    for (const [position, tc] of message.tool_calls.entries()) {
      const index = tc.index ?? position;
      if (!Number.isInteger(index) || index < 0) throw new Error("模型返回了无效的工具调用序号");
      const acc = calls.get(index) || { id: `call_${Date.now()}_${index}`, name: "", args: "", lastArgs: "" };
      if (typeof tc.id === "string" && tc.id) acc.id = tc.id;
      if (typeof tc.function?.name === "string") {
        acc.name = fullMessage ? tc.function.name : appendStreamFragment(acc.name, tc.function.name);
      }
      if (tc.function?.arguments !== undefined) {
        if (typeof tc.function.arguments !== "string") throw new Error("模型返回的工具 arguments 必须是 JSON 字符串");
        const incoming = tc.function.arguments;
        if (incoming) {
          // 少数网关把完整参数快照放在 delta 中。只记录逐步扩展且完整的 JSON 备选，
          // 先优先采用原生拼接结果，不能让去重改变 11、1000、重复字符串等数据。
          acc.snapshotArgs = acc.lastArgs.trim().startsWith("{") && incoming.startsWith(acc.lastArgs) && isObjectJson(incoming)
            ? incoming : undefined;
          acc.lastArgs = incoming;
        }
        acc.args = fullMessage ? incoming : acc.args + incoming;
      }
      calls.set(index, acc);
    }
  };

  if (response.headers.get("content-type")?.includes("application/json")) {
    consume(await response.json(), true);
    done = true;
  } else {
    if (!response.body) throw new Error("模型响应没有内容流");
    const reader = response.body.getReader();
    const cancelOnAbort = () => { void reader.cancel().catch(() => undefined); };
    signal.addEventListener("abort", cancelOnAbort, { once: true });
    const decoder = new TextDecoder();
    let buffer = "";
    let data: string[] = [];
    const event = () => {
      if (!data.length) return;
      const payload = data.join("\n").trim();
      data = [];
      if (payload === "[DONE]") { done = true; return; }
      let json: unknown;
      try { json = JSON.parse(payload); }
      catch { throw new Error("模型返回了无效的流式 JSON，任务已停止，请重试。"); }
      consume(json);
    };
    const line = (text: string) => {
      if (!text.trim()) event();
      else if (text.startsWith("data:")) data.push(text.slice(5).replace(/^ /, ""));
    };
    try {
      while (!done) {
        if (signal.aborted) throw new DOMException("任务已停止", "AbortError");
        const chunk = await reader.read();
        buffer += decoder.decode(chunk.value, { stream: !chunk.done });
        const lines = buffer.split(/\r?\n/);
        buffer = lines.pop() || "";
        for (const text of lines) {
          line(text);
          if (done) break;
        }
        if (chunk.done) {
          if (!done && buffer) line(buffer);
          if (!done) event();
          break;
        }
      }
    } finally {
      signal.removeEventListener("abort", cancelOnAbort);
      await reader.cancel().catch(() => undefined);
      reader.releaseLock();
    }
  }
  if (signal.aborted) throw new DOMException("任务已停止", "AbortError");
  if (!done && !finishReason) throw new Error("模型响应在完成前中断，请重试；本轮工具尚未执行。");
  const toolCalls: ToolCall[] = [...calls.entries()].sort(([a], [b]) => a - b).map(([, call]) => {
    if (!call.name.trim()) throw new Error("模型返回了缺少名称的工具调用，本轮工具尚未执行。");
    const args = call.args.trim() || "{}";
    const complete = finishReason !== "length" && finishReason !== "max_tokens";
    const argumentsJson = complete && !isObjectJson(args) && call.snapshotArgs ? call.snapshotArgs : args;
    return { id: call.id, type: "function", function: { name: call.name.trim(), arguments: argumentsJson } };
  });
  if (new Set(toolCalls.map((call) => call.id)).size !== toolCalls.length) throw new Error("模型返回了重复的工具调用 ID，请重试。");
  const visible = stripReasoningText(content);
  if (!visible && !toolCalls.length) {
    throw new Error(`模型没有返回可用回复或工具调用${finishReason ? `（${finishReason}）` : ""}，请重试或更换模型。`);
  }
  return { content: visible, toolCalls, finishReason, usage };
}
