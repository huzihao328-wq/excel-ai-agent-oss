import { pickContextTokens } from "./contextWindow";

/** Parse OpenAI-compatible /models payloads into a unique, sorted id list. */

const ID_KEYS = ["id", "name", "model", "model_name", "modelId"] as const;
const LIST_KEYS = ["data", "models", "list", "items", "result"] as const;

export interface CatalogModel {
  id: string;
  contextWindow?: number;
  /** 目录里声明的单次输出上限（如 OpenRouter 的 top_provider.max_completion_tokens） */
  maxOutputTokens?: number;
}

const OUTPUT_KEY = /^(max[_-]?(completion|output)[_-]?tokens|max[_-]?output[_-]?length|output[_-]?token[_-]?limit)$/i;

/** 在模型对象（含 top_provider / limits 等嵌套块）里找单次输出上限。 */
function pickOutputTokens(value: unknown, depth = 0): number | null {
  if (depth > 3 || !value || typeof value !== "object" || Array.isArray(value)) return null;
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    if (!OUTPUT_KEY.test(key)) continue;
    const n = typeof raw === "number" ? raw : typeof raw === "string" ? Number(raw) : NaN;
    if (Number.isFinite(n) && n >= 256 && n <= 2_000_000) return Math.round(n);
  }
  for (const [key, raw] of Object.entries(value as Record<string, unknown>)) {
    if (raw && typeof raw === "object" && /provider|limit|meta|capabilit|spec/i.test(key)) {
      const nested = pickOutputTokens(raw, depth + 1);
      if (nested) return nested;
    }
  }
  return null;
}

function pickModelId(value: unknown): string | null {
  if (typeof value === "string") {
    const id = value.trim();
    return id ? id : null;
  }
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  for (const key of ID_KEYS) {
    const raw = record[key];
    if (typeof raw === "string" && raw.trim()) return raw.trim();
  }
  return null;
}

function catalogFromItem(value: unknown): CatalogModel | null {
  const id = pickModelId(value);
  if (!id) return null;
  const contextWindow = typeof value === "object" && value ? pickContextTokens(value) ?? undefined : undefined;
  const maxOutputTokens = pickOutputTokens(value) ?? undefined;
  return { id, ...(contextWindow ? { contextWindow } : {}), ...(maxOutputTokens ? { maxOutputTokens } : {}) };
}

/**
 * Accepts the common shapes used by OpenAI, Ollama, DashScope, GLM, Ark, and
 * aggregators: `{ data: [{ id }] }`, `{ models: [{ name }] }`, nested `data.data`,
 * or a raw string array.
 */
export function extractModelCatalog(payload: unknown): CatalogModel[] {
  const found: CatalogModel[] = [];
  const visit = (value: unknown, depth: number) => {
    if (depth > 5 || value == null) return;
    if (typeof value === "string") {
      const item = catalogFromItem(value);
      if (item) found.push(item);
      return;
    }
    if (Array.isArray(value)) {
      for (const item of value) visit(item, depth + 1);
      return;
    }
    if (typeof value !== "object") return;
    const record = value as Record<string, unknown>;
    const item = catalogFromItem(record);
    if (item) found.push(item);
    for (const key of LIST_KEYS) {
      if (key in record) visit(record[key], depth + 1);
    }
  };
  visit(payload, 0);
  const byId = new Map<string, CatalogModel>();
  for (const item of found) {
    const prev = byId.get(item.id);
    if (!prev) byId.set(item.id, item);
    else if ((item.contextWindow && !prev.contextWindow) || (item.maxOutputTokens && !prev.maxOutputTokens)) {
      byId.set(item.id, { ...prev, ...item, contextWindow: prev.contextWindow ?? item.contextWindow, maxOutputTokens: prev.maxOutputTokens ?? item.maxOutputTokens });
    }
  }
  return [...byId.values()].sort((a, b) => a.id.localeCompare(b.id));
}

export function extractModelIds(payload: unknown): string[] {
  return extractModelCatalog(payload).map((item) => item.id);
}

export function mergeModelLists(...lists: (string[] | undefined)[]): string[] {
  const ids = new Set<string>();
  for (const list of lists) {
    if (!list) continue;
    for (const item of list) {
      const id = item.trim();
      if (id) ids.add(id);
    }
  }
  return [...ids];
}
