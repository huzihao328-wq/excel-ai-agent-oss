import { ChatMessage, ContentPart, UiItem } from "./types";

const INDEX_KEY = "excel-ai-agent-conv-index";
const CONV_PREFIX = "excel-ai-agent-conv:";
const MAX_CONVERSATIONS = 40;

export interface ConversationMeta {
  id: string;
  title: string;
  createdAt: number;
  updatedAt: number;
}

export interface ConversationRecord extends ConversationMeta {
  items: UiItem[];
  messages: ChatMessage[];
}

interface ConversationIndex {
  currentId: string | null;
  items: ConversationMeta[];
}

function readIndex(): ConversationIndex {
  try {
    const raw = localStorage.getItem(INDEX_KEY);
    if (!raw) return { currentId: null, items: [] };
    const parsed = JSON.parse(raw);
    return {
      currentId: typeof parsed.currentId === "string" ? parsed.currentId : null,
      items: Array.isArray(parsed.items) ? parsed.items.filter((item: unknown): item is ConversationMeta => {
        if (!item || typeof item !== "object") return false;
        const entry = item as ConversationMeta;
        return typeof entry.id === "string" && !!entry.id && typeof entry.title === "string" &&
          Number.isFinite(entry.createdAt) && Number.isFinite(entry.updatedAt);
      }) : [],
    };
  } catch {
    return { currentId: null, items: [] };
  }
}

function writeIndex(index: ConversationIndex): void {
  localStorage.setItem(INDEX_KEY, JSON.stringify(index));
}

export function listConversations(): ConversationMeta[] {
  return readIndex().items;
}

export function getCurrentConversationId(): string | null {
  return readIndex().currentId;
}

export function setCurrentConversationId(id: string | null): boolean {
  const index = readIndex();
  index.currentId = id;
  try {
    writeIndex(index);
    return true;
  } catch {
    return false;
  }
}

export function loadConversation(id: string): ConversationRecord | null {
  try {
    const raw = localStorage.getItem(CONV_PREFIX + id);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as ConversationRecord;
    if (!parsed || parsed.id !== id || !Array.isArray(parsed.items) || !Array.isArray(parsed.messages)) {
      return null;
    }
    return {
      ...parsed,
      items: sanitizeItems(parsed.items),
      messages: sanitizeMessages(parsed.messages),
    };
  } catch {
    return null;
  }
}

export function saveConversation(record: ConversationRecord): boolean {
  let sanitized: ConversationRecord = {
    ...record,
    items: sanitizeItems(record.items),
    messages: sanitizeMessages(record.messages),
    updatedAt: Date.now(),
  };

  const persist = () => {
    localStorage.setItem(CONV_PREFIX + sanitized.id, JSON.stringify(sanitized));
    const index = readIndex();
    const meta: ConversationMeta = {
      id: sanitized.id,
      title: sanitized.title,
      createdAt: sanitized.createdAt,
      updatedAt: sanitized.updatedAt,
    };
    index.items = [meta, ...index.items.filter((x) => x.id !== meta.id)];
    if (index.items.length > MAX_CONVERSATIONS) {
      const dropped = index.items.slice(MAX_CONVERSATIONS);
      index.items = index.items.slice(0, MAX_CONVERSATIONS);
      for (const d of dropped) localStorage.removeItem(CONV_PREFIX + d.id);
    }
    writeIndex(index);
  };

  try {
    persist();
    return true;
  } catch {
    /* 存储配额不够：先给自己瘦身，再考虑腾别人的位置 */
  }

  // 第一步：去掉图片（气泡里的 dataUrl 与消息里的 image_url 往往就是几 MB 的主因）
  try {
    sanitized = slimConversation(sanitized);
    persist();
    return true;
  } catch {
    /* 继续 */
  }

  // 第二步：从最旧的对话开始删（不删当前正在看的、也不删正在保存的），直到能写进去
  try {
    const snapshot = readIndex();
    const working: ConversationIndex = { currentId: snapshot.currentId, items: [...snapshot.items] };
    const droppedIds: string[] = [];
    for (let i = working.items.length - 1; i >= 0; i--) {
      const candidate = working.items[i];
      if (candidate.id === sanitized.id || candidate.id === snapshot.currentId) continue;
      working.items.splice(i, 1);
      localStorage.removeItem(CONV_PREFIX + candidate.id);
      droppedIds.push(candidate.id);
      try {
        writeIndex(working);
        persist();
        return true;
      } catch {
        /* keep dropping until the current blob fits */
      }
    }
    writeIndex({
      currentId: snapshot.currentId,
      items: snapshot.items.filter((item) => !droppedIds.includes(item.id)),
    });
  } catch {
    /* 实在写不进去就放弃这次落盘，不能让调用方（UI 的 finally）抛错 */
  }
  return false;
}

/** 丢掉图片数据，保留文字与工具轨迹。 */
function slimConversation(record: ConversationRecord): ConversationRecord {
  const items = record.items.map((it) => (it.kind === "user" && it.images?.length ? { ...it, images: [] } : it));
  const messages = record.messages.map((m) => {
    if (!Array.isArray(m.content)) return m;
    const textParts = m.content.filter((p): p is Extract<ContentPart, { type: "text" }> => p.type === "text");
    const text = textParts.map((p) => p.text).join("\n");
    const dropped = m.content.length - textParts.length;
    return { ...m, content: dropped ? `${text}\n[已移除 ${dropped} 张图片以节省存储]` : text };
  });
  return { ...record, items, messages };
}

export function deleteConversation(id: string): void {
  localStorage.removeItem(CONV_PREFIX + id);
  const index = readIndex();
  index.items = index.items.filter((x) => x.id !== id);
  if (index.currentId === id) index.currentId = null;
  writeIndex(index);
}

export function titleFromText(text: string): string {
  const t = text.replace(/\s+/g, " ").trim();
  if (!t) return "新对话";
  return t.length > 28 ? `${t.slice(0, 28)}…` : t;
}

export function sanitizeItems(items: UiItem[]): UiItem[] {
  // 存储里混进坏数据（null、缺 kind/id）时直接丢弃，不能让一条脏记录把整个界面渲染崩掉
  const valid = items.filter(
    (it): it is UiItem => !!it && typeof it === "object" && typeof (it as UiItem).kind === "string" && typeof (it as UiItem).id === "string"
  );
  return valid.map((it) => {
    if (it.kind === "assistant") return { ...it, streaming: false };
    // 窗格关闭时还没送达的补充说明：模型没看到过它，如实标成未送达，原文留给用户复制
    if (it.kind === "user" && it.pending) return { kind: "notice", id: it.id, text: `补充说明未送达：${it.text}` };
    if (it.kind === "tool" && it.status === "running") {
      return { ...it, status: "error", detail: it.detail || "已中断" };
    }
    if (it.kind === "confirm" && !it.resolved) {
      return { ...it, resolved: "rejected" };
    }
    if (it.kind === "choice" && !it.resolved) {
      return { ...it, resolved: "cancelled" };
    }
    return it;
  });
}

/** 给悬空的 tool_calls 补上失败结果，避免下次请求因缺 tool 消息被接口 400。 */
export function sanitizeMessages(input: ChatMessage[]): ChatMessage[] {
  const messages = input.filter((m): m is ChatMessage => !!m && typeof m === "object" && typeof m.role === "string");
  const out: ChatMessage[] = [];
  let i = 0;
  while (i < messages.length) {
    const m = messages[i];
    out.push(m);
    i++;
    if (m.role !== "assistant" || !m.tool_calls?.length) continue;
    const needed = new Set(m.tool_calls.map((tc) => tc.id));
    while (i < messages.length && messages[i].role === "tool") {
      out.push(messages[i]);
      const callId = messages[i].tool_call_id;
      if (callId) needed.delete(callId);
      i++;
    }
    for (const tc of m.tool_calls) {
      if (!needed.has(tc.id)) continue;
      out.push({
        role: "tool",
        content: JSON.stringify({ ok: false, error: "操作未完成（会话中断），该工具未执行。" }),
        tool_call_id: tc.id,
      });
    }
  }
  return out;
}
