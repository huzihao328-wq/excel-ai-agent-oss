/* global Office */
import { MEMORY_STORAGE_KEY, flushUserData, schedulePersistUserData } from "./durableStore";
import { ToolResult } from "./types";

export type MemoryScope = "global" | "workbook";

export interface MemoryEntry {
  id: string;
  scope: MemoryScope;
  /** 同一 key 再次记住会覆盖，避免重复条目 */
  key?: string;
  content: string;
  createdAt: number;
  updatedAt: number;
}

const WB_STORE_KEY = "excel-ai-agent-memory";
const WB_ID_KEY = "excel-ai-agent-wb-id";
const WB_LOCAL_PREFIX = "excel-ai-agent-memory-wb:";
const MAX_PER_SCOPE = 40;
const MAX_CONTENT = 400;
const SECRET_RE = /(api[_-]?key|password|secret|token|sk-[a-z0-9]{12,})/i;

let listeners: (() => void)[] = [];

export function onMemoryChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((f) => f !== fn);
  };
}

function notify() {
  listeners.forEach((f) => f());
}

function parseEntries(raw: unknown, scope: MemoryScope): MemoryEntry[] {
  if (typeof raw !== "string" || !raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .filter((x) => x && typeof x.content === "string" && x.id)
      .map((x) => ({
        id: String(x.id),
        scope,
        content: String(x.content),
        createdAt: Number(x.createdAt) || Date.now(),
        updatedAt: Number(x.updatedAt) || Date.now(),
        ...(x.key ? { key: String(x.key) } : {}),
      }));
  } catch {
    return [];
  }
}

function loadGlobal(): MemoryEntry[] {
  try {
    return parseEntries(localStorage.getItem(MEMORY_STORAGE_KEY), "global");
  } catch {
    return [];
  }
}

function saveGlobal(items: MemoryEntry[]): void {
  localStorage.setItem(MEMORY_STORAGE_KEY, JSON.stringify(items));
  schedulePersistUserData();
}

function getWorkbookId(): string {
  try {
    const existing = Office.context.document.settings.get(WB_ID_KEY);
    if (typeof existing === "string" && existing) return existing;
    const id = `wb_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
    Office.context.document.settings.set(WB_ID_KEY, id);
    Office.context.document.settings.saveAsync();
    return id;
  } catch {
    return "local";
  }
}

function loadWorkbook(): MemoryEntry[] {
  try {
    const fromDoc = Office.context.document.settings.get(WB_STORE_KEY);
    const docItems = parseEntries(fromDoc, "workbook");
    if (docItems.length) return docItems;
  } catch {
    /* 非 Excel 或尚未就绪 */
  }
  try {
    return parseEntries(localStorage.getItem(WB_LOCAL_PREFIX + getWorkbookId()), "workbook");
  } catch {
    return [];
  }
}

function saveWorkbook(items: MemoryEntry[]): void {
  const raw = JSON.stringify(items);
  try {
    Office.context.document.settings.set(WB_STORE_KEY, raw);
    Office.context.document.settings.saveAsync();
  } catch {
    /* ignore */
  }
  try {
    localStorage.setItem(WB_LOCAL_PREFIX + getWorkbookId(), raw);
  } catch {
    /* ignore */
  }
}

export function loadAllMemories(): MemoryEntry[] {
  return [...loadGlobal(), ...loadWorkbook()];
}

export function formatMemoryFile(entries: MemoryEntry[]): string {
  if (!entries.length) return "";
  const global = entries.filter((e) => e.scope === "global");
  const workbook = entries.filter((e) => e.scope === "workbook");
  const lines = ["[长期记忆] 做事时遵守这些约定。用 remember / forget 更新，不要重复提问已经记住的事。"];
  if (global.length) {
    lines.push("## 全局习惯");
    for (const e of global) lines.push(`- (${e.key || e.id}) ${e.content}`);
  }
  if (workbook.length) {
    lines.push("## 当前工作簿");
    for (const e of workbook) lines.push(`- (${e.key || e.id}) ${e.content}`);
  }
  return lines.join("\n");
}

function nextId(): string {
  return `mem_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
}

function slugKey(key: string | undefined): string | undefined {
  if (!key) return undefined;
  const s = key.trim().replace(/\s+/g, "_").slice(0, 40);
  return s || undefined;
}

export function upsertMemory(input: {
  content: string;
  scope?: MemoryScope;
  key?: string;
  id?: string;
}): MemoryEntry {
  const content = input.content.replace(/\s+/g, " ").trim().slice(0, MAX_CONTENT);
  if (!content) throw new Error("记忆内容不能为空");
  if (SECRET_RE.test(content)) throw new Error("不要把密码、API Key 或令牌写入记忆");

  const scope: MemoryScope = input.scope === "workbook" ? "workbook" : "global";
  const key = slugKey(input.key);
  const list = scope === "global" ? loadGlobal() : loadWorkbook();
  const now = Date.now();

  let entry = input.id ? list.find((e) => e.id === input.id) : undefined;
  if (!entry && key) entry = list.find((e) => e.key === key);

  if (entry) {
    entry.content = content;
    entry.updatedAt = now;
    if (key) entry.key = key;
  } else {
    entry = { id: nextId(), scope, content, createdAt: now, updatedAt: now, ...(key ? { key } : {}) };
    list.unshift(entry);
  }

  const trimmed = list
    .sort((a, b) => b.updatedAt - a.updatedAt)
    .slice(0, MAX_PER_SCOPE);
  if (scope === "global") saveGlobal(trimmed);
  else saveWorkbook(trimmed);
  notify();
  return entry;
}

export function removeMemory(idOrKey: string): MemoryEntry | null {
  const needle = idOrKey.trim();
  if (!needle) return null;
  const tryRemove = (scope: MemoryScope): MemoryEntry | null => {
    const list = scope === "global" ? loadGlobal() : loadWorkbook();
    const idx = list.findIndex((e) => e.id === needle || e.key === needle);
    if (idx < 0) return null;
    const [removed] = list.splice(idx, 1);
    if (scope === "global") saveGlobal(list);
    else saveWorkbook(list);
    return removed;
  };
  const removed = tryRemove("workbook") || tryRemove("global");
  if (removed) notify();
  return removed;
}

export function clearMemories(scope?: MemoryScope): void {
  if (!scope || scope === "global") saveGlobal([]);
  if (!scope || scope === "workbook") saveWorkbook([]);
  notify();
  if (!scope || scope === "global") void flushUserData();
}

export function toolRemember(args: any): ToolResult {
  try {
    const entry = upsertMemory({
      content: String(args.content || ""),
      scope: args.scope === "workbook" ? "workbook" : "global",
      key: args.key ? String(args.key) : undefined,
    });
    return {
      ok: true,
      data: { id: entry.id, key: entry.key, scope: entry.scope, content: entry.content },
      summary: `已记住（${entry.scope === "workbook" ? "本工作簿" : "全局"}）：${entry.content}`,
    };
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
}

export function toolForget(args: any): ToolResult {
  const idOrKey = String(args.id || args.key || "").trim();
  if (!idOrKey) return { ok: false, error: "forget 需要 id 或 key" };
  const removed = removeMemory(idOrKey);
  if (!removed) return { ok: false, error: `没有找到记忆「${idOrKey}」` };
  return {
    ok: true,
    data: { removed: removed.id, key: removed.key },
    summary: `已忘掉：${removed.content}`,
  };
}

export function toolListMemories(): ToolResult {
  const items = loadAllMemories();
  return {
    ok: true,
    data: {
      memories: items.map((e) => ({
        id: e.id,
        key: e.key,
        scope: e.scope,
        content: e.content,
      })),
    },
    summary: items.length ? `共 ${items.length} 条记忆` : "还没有记忆",
  };
}
