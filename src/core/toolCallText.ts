import type { ToolCall } from "./types";

/**
 * 有些模型/网关不走原生 tool_calls，而是把调用写进正文：
 * `<tool_call>{...}</tool_call>`、`<function=name>{...}</function>`、```json 代码块、裸 JSON、`name({...})`。
 * 这里把这些形式解析成标准 ToolCall，只接受已知工具名，其余文字原样保留。
 */

export interface ExtractedToolCalls {
  content: string;
  toolCalls: ToolCall[];
}

interface Segment {
  start: number;
  end: number;
  calls: ToolCall[];
}

const NAME_KEYS = ["name", "tool", "function", "function_name", "tool_name", "action"];
const ARG_KEYS = ["arguments", "parameters", "params", "input", "args", "action_input"];

let seq = 0;
function newId(): string {
  seq += 1;
  return `text_call_${Date.now().toString(36)}_${seq}`;
}

function argsToJson(raw: unknown): string {
  if (raw == null) return "{}";
  if (typeof raw === "string") {
    const t = raw.trim();
    if (!t) return "{}";
    try {
      const parsed = JSON.parse(t);
      return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? JSON.stringify(parsed) : "{}";
    } catch {
      return "{}";
    }
  }
  if (typeof raw === "object" && !Array.isArray(raw)) return JSON.stringify(raw);
  return "{}";
}

/** 把各种形状的 JSON 候选归一化成 ToolCall；不是已知工具则返回空数组。 */
function normalizeCandidate(value: unknown, known: Set<string>, depth = 0): ToolCall[] {
  if (depth > 3 || value == null || typeof value !== "object") return [];
  if (Array.isArray(value)) {
    const out: ToolCall[] = [];
    for (const item of value) out.push(...normalizeCandidate(item, known, depth + 1));
    return out;
  }
  const rec = value as Record<string, unknown>;
  if (Array.isArray(rec.tool_calls)) return normalizeCandidate(rec.tool_calls, known, depth + 1);
  if (rec.function && typeof rec.function === "object") {
    const fn = rec.function as Record<string, unknown>;
    if (typeof fn.name === "string" && known.has(fn.name)) {
      const args = ARG_KEYS.map((k) => fn[k]).find((v) => v !== undefined);
      return [
        {
          id: typeof rec.id === "string" && rec.id ? rec.id : newId(),
          type: "function",
          function: { name: fn.name, arguments: argsToJson(args) },
        },
      ];
    }
  }
  let name = "";
  for (const key of NAME_KEYS) {
    const v = rec[key];
    if (typeof v === "string" && known.has(v.trim())) {
      name = v.trim();
      break;
    }
  }
  if (!name) return [];
  const args = ARG_KEYS.map((k) => rec[k]).find((v) => v !== undefined);
  // 没有 arguments 字段时，把其余字段当作参数（如 {"name":"read_range","sheet":"Sheet1"}）
  let argJson: string;
  if (args !== undefined) {
    argJson = argsToJson(args);
  } else {
    const rest: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(rec)) if (!NAME_KEYS.includes(k) && k !== "id" && k !== "type") rest[k] = v;
    argJson = JSON.stringify(rest);
  }
  return [{ id: newId(), type: "function", function: { name, arguments: argJson } }];
}

/** 从 from 开始找到第一个能完整闭合并解析的 JSON 对象/数组。 */
function scanJson(text: string, from: number): { start: number; end: number; value: unknown } | null {
  for (let i = from; i < text.length; i++) {
    const ch = text[i];
    if (ch !== "{" && ch !== "[") continue;
    const stack: string[] = [];
    let inStr = false;
    let esc = false;
    for (let j = i; j < text.length; j++) {
      const c = text[j];
      if (inStr) {
        if (esc) esc = false;
        else if (c === "\\") esc = true;
        else if (c === '"') inStr = false;
        continue;
      }
      if (c === '"') inStr = true;
      else if (c === "{" || c === "[") stack.push(c);
      else if (c === "}" || c === "]") {
        stack.pop();
        if (!stack.length) {
          const slice = text.slice(i, j + 1);
          try {
            return { start: i, end: j + 1, value: JSON.parse(slice) };
          } catch {
            break;
          }
        }
      }
    }
  }
  return null;
}

function pushSegment(list: Segment[], seg: Segment): void {
  if (!seg.calls.length) return;
  if (list.some((s) => seg.start < s.end && s.start < seg.end)) return;
  list.push(seg);
}

export function extractTextToolCalls(content: string, knownToolNames: Iterable<string>): ExtractedToolCalls {
  const known = new Set(knownToolNames);
  if (!content || !known.size) return { content, toolCalls: [] };
  const segments: Segment[] = [];

  // 1) <tool_call>…</tool_call>（Qwen / Hermes）、<tool_call> 未闭合也接受
  const tagRe = /<(tool_call|tool_calls|function_call|invoke)>([\s\S]*?)(?:<\/\1>|$)/gi;
  for (let m = tagRe.exec(content); m; m = tagRe.exec(content)) {
    const found = scanJson(m[2], 0);
    if (!found) continue;
    pushSegment(segments, { start: m.index, end: m.index + m[0].length, calls: normalizeCandidate(found.value, known) });
  }

  // 2) <function=name>{…}</function>（Llama 3.x）
  const fnRe = /<function=([\w.-]+)>([\s\S]*?)(?:<\/function>|$)/gi;
  for (let m = fnRe.exec(content); m; m = fnRe.exec(content)) {
    if (!known.has(m[1])) continue;
    const found = scanJson(m[2], 0);
    pushSegment(segments, {
      start: m.index,
      end: m.index + m[0].length,
      calls: [{ id: newId(), type: "function", function: { name: m[1], arguments: argsToJson(found?.value) } }],
    });
  }

  // 3) ```json / ```tool_call 代码块
  const fenceRe = /```[a-zA-Z_-]*\s*\n?([\s\S]*?)```/g;
  for (let m = fenceRe.exec(content); m; m = fenceRe.exec(content)) {
    const found = scanJson(m[1], 0);
    if (!found) continue;
    pushSegment(segments, { start: m.index, end: m.index + m[0].length, calls: normalizeCandidate(found.value, known) });
  }

  // 4) 裸 JSON（整段就是一个对象/数组，或正文里夹着 {"name": "已知工具", …}）
  if (!segments.length) {
    let pos = 0;
    for (let guard = 0; guard < 20; guard++) {
      const found = scanJson(content, pos);
      if (!found) break;
      pushSegment(segments, { start: found.start, end: found.end, calls: normalizeCandidate(found.value, known) });
      pos = found.end;
    }
  }

  // 5) 独占一行的 name({...}) / name()（可包在 [] 里）
  if (!segments.length) {
    const lineRe = /^[ \t]*\[?\s*([a-z_][\w]*)\s*\(([\s\S]*?)\)\s*\]?[ \t]*$/gm;
    for (let m = lineRe.exec(content); m; m = lineRe.exec(content)) {
      if (!known.has(m[1])) continue;
      const inner = m[2].trim();
      const found = inner ? scanJson(inner, 0) : null;
      if (inner && (!found || found.start !== 0 || found.end !== inner.length)) continue;
      pushSegment(segments, {
        start: m.index,
        end: m.index + m[0].length,
        calls: [{ id: newId(), type: "function", function: { name: m[1], arguments: argsToJson(found?.value) } }],
      });
    }
  }

  if (!segments.length) return { content, toolCalls: [] };
  segments.sort((a, b) => a.start - b.start);
  const toolCalls: ToolCall[] = [];
  let rest = "";
  let cursor = 0;
  for (const seg of segments) {
    rest += content.slice(cursor, seg.start);
    cursor = seg.end;
    toolCalls.push(...seg.calls);
  }
  rest += content.slice(cursor);
  const cleaned = rest
    .replace(/<\/?(tool_call|tool_calls|function_call|invoke|function)>/gi, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
  return { content: cleaned, toolCalls };
}

const PLAN_INTENT =
  /(我|让我|接下来|下一步|首先|然后|现在|马上)(将|会|来|先|要|需要|准备|打算|开始)|我将|我会先|让我先|接下来我|开始(处理|执行|分析|操作)|以下是(我的)?(计划|步骤|方案)|步骤\s*[1一]/;

/**
 * 判断一段没有工具调用的回复是否只是「口头计划」：
 * 有明确的将要动手的措辞，又提到了具体工具或工具操作，且没有在向用户提问。
 */
export function looksLikePlanOnly(content: string, knownToolNames: Iterable<string>): boolean {
  const text = content.trim();
  if (text.length < 12) return false;
  const tail = text.slice(-60);
  if (/[?？]\s*$/.test(text) || /(请|您|你)[^。\n]{0,20}(选择|确认|告诉|提供|指定)/.test(tail)) return false;
  if (!PLAN_INTENT.test(text)) return false;
  for (const name of knownToolNames) {
    if (text.includes(name)) return true;
  }
  return /工具|调用|读取(工作簿|表格|数据)|查看(工作簿|表格|结构)/.test(text);
}

/** 工具名拼错时给出最接近的候选（编辑距离 ≤ 3 或前缀/子串匹配）。 */
export function suggestToolName(name: string, knownToolNames: Iterable<string>): string | null {
  const target = name.trim().toLowerCase();
  if (!target) return null;
  let best: { name: string; score: number } | null = null;
  for (const candidate of knownToolNames) {
    const c = candidate.toLowerCase();
    let score: number;
    if (c === target) return candidate;
    if (c.includes(target) || target.includes(c)) score = 1;
    else score = editDistance(target, c);
    if (score <= 3 && (!best || score < best.score)) best = { name: candidate, score };
  }
  return best?.name ?? null;
}

function editDistance(a: string, b: string): number {
  if (Math.abs(a.length - b.length) > 3) return 99;
  const prev = new Array(b.length + 1).fill(0).map((_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let last = prev[0];
    prev[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = prev[j];
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, last + (a[i - 1] === b[j - 1] ? 0 : 1));
      last = tmp;
    }
  }
  return prev[b.length];
}
