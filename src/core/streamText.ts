/**
 * Merge streamed chat chunks. Some OpenAI-compatible gateways send incremental
 * deltas; others resend the full text (or the last sentence) on every event.
 * Concatenating snapshots makes the UI repeat the same sentence.
 * Reasoning / chain-of-thought must stay out of the user-visible reply.
 */

const REPLAY_MIN_CHARS = 8;
const SKIP_CONTENT_TYPES = /^(thinking|thought|reasoning|reasoning_content|reasoning_text|redacted_reasoning)$/i;

const THINK_BLOCK =
  /<(think|thinking|thought|reason|reasoning)(?:\s[^>]*)?>[\s\S]*?<\/\1>/gi;
const THINK_SPECIAL =
  /<\|(?:begin_of_thought|redacted_reasoning)\|>[\s\S]*?<\|(?:end_of_thought|end_redacted_reasoning|redacted_reasoning)\|>/gi;
const THINK_KIMI = /◁think▷[\s\S]*?(?:◁\/think▷|$)/g;
const THINK_BRACKET = /\[\/?(?:think|thinking|thought)\][\s\S]*?\[\/(?:think|thinking|thought)\]/gi;
const THINK_OPEN_TAIL =
  /<(?:think|thinking|thought|reason|reasoning)(?:\s[^>]*)?>[\s\S]*$/i;
const THINK_SPECIAL_TAIL = /<\|(?:begin_of_thought|redacted_reasoning)\|>[\s\S]*$/i;
const THINK_KIMI_TAIL = /◁think▷[\s\S]*$/;

function shouldCollapseUnit(text: string): boolean {
  const t = text.trim();
  if (t.length < 4) return false;
  if (t.length >= 12) return true;
  return /[。！？.!?；;\n]/.test(t);
}

/** Drop an exact doubled body, or a consecutive duplicate sentence. */
export function collapseDuplicatedText(text: string): string {
  let t = text;
  if (t.length >= 8 && t.length % 2 === 0) {
    const half = t.length / 2;
    const unit = t.slice(0, half);
    if (unit === t.slice(half) && shouldCollapseUnit(unit)) t = unit;
  }
  const paras = t.split(/\n+/);
  if (paras.length === 2 && paras[0].trim() === paras[1].trim() && shouldCollapseUnit(paras[0])) {
    t = paras[0].trim();
  }
  const parts = t.split(/([。！？\n]+)/);
  if (parts.length >= 4) {
    const out: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      if (
        i >= 2 &&
        parts[i] &&
        parts[i] === parts[i - 2] &&
        parts[i].trim().length >= 4 &&
        (i + 1 >= parts.length || parts[i + 1] === parts[i - 1])
      ) {
        if (i + 1 < parts.length && parts[i + 1] === parts[i - 1]) i++;
        continue;
      }
      out.push(parts[i]);
    }
    t = out.join("");
  }
  return t;
}

/** Remove chain-of-thought blocks, including a still-open tag while streaming. */
export function stripReasoningText(text: string): string {
  let t = text;
  t = t.replace(THINK_BLOCK, "");
  t = t.replace(THINK_SPECIAL, "");
  t = t.replace(THINK_KIMI, "");
  t = t.replace(THINK_BRACKET, "");
  t = t.replace(THINK_OPEN_TAIL, "");
  t = t.replace(THINK_SPECIAL_TAIL, "");
  t = t.replace(THINK_KIMI_TAIL, "");
  return t.replace(/^\s+/, "");
}

/**
 * Fold one incoming stream piece into the text already shown.
 * Incremental tokens are appended; cumulative snapshots replace.
 * Does not collapse duplicate sentences — safe for tool names and JSON arguments.
 */
export function appendStreamFragment(previous: string, incoming: string): string {
  const piece = incoming;
  if (!piece) return previous;
  if (!previous) return piece;
  if (piece === previous) return previous;
  if (piece.startsWith(previous)) return piece;
  if (previous.startsWith(piece) && piece.length >= 8) return previous;
  if (previous.endsWith(piece) && piece.length >= REPLAY_MIN_CHARS) return previous;
  return previous + piece;
}

export function appendStreamText(previous: string, incoming: string): string {
  return collapseDuplicatedText(appendStreamFragment(previous, incoming));
}

/**
 * Always return valid JSON, shrinking oversized tool results instead of slicing the string.
 * Prefers trimming the biggest arrays/strings in place, so the model still gets parseable
 * rows plus an `omitted` note saying exactly what was cut; falls back to a text preview.
 */
export function clipJsonToBudget(value: unknown, maxChars: number): string {
  const full = JSON.stringify(value);
  if (full.length <= maxChars) return full;
  const hint = "结果过长已截断，请用更小的范围重新查询";
  const ok = !(value && typeof value === "object" && (value as { ok?: unknown }).ok === false);
  const structured = shrinkStructured(value, maxChars, ok, hint);
  if (structured) return structured;
  let budget = Math.max(0, maxChars - 96);
  for (;;) {
    const out = JSON.stringify({ ok, truncated: true, hint, preview: full.slice(0, budget) });
    if (out.length <= maxChars || budget === 0) {
      return out.length <= maxChars ? out : JSON.stringify({ ok, truncated: true, hint });
    }
    budget = Math.max(0, budget - (out.length - maxChars));
  }
}

interface ShrinkSlot {
  holder: any;
  key: string | number;
  path: string;
  size: number;
}

/** The largest array (with more than one item) or long string anywhere in the value. */
function largestShrinkable(root: Record<string, unknown>): ShrinkSlot | null {
  let best: ShrinkSlot | null = null;
  const visit = (holder: any, key: string | number, path: string) => {
    const v = holder[key];
    if (typeof v === "string") {
      if (v.length > 200 && (!best || v.length > best.size)) best = { holder, key, path, size: v.length };
      return;
    }
    if (!v || typeof v !== "object") return;
    if (Array.isArray(v)) {
      const size = JSON.stringify(v).length;
      if (v.length > 1 && (!best || size > best.size)) best = { holder, key, path, size };
      v.forEach((_, i) => visit(v, i, `${path}[${i}]`));
    } else {
      for (const k of Object.keys(v)) visit(v, k, `${path}.${k}`);
    }
  };
  for (const k of Object.keys(root)) visit(root, k, k);
  return best;
}

function shrinkStructured(value: unknown, maxChars: number, ok: boolean, hint: string): string | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const body = JSON.parse(JSON.stringify(value)) as Record<string, unknown>;
  delete body.ok;
  const omitted = new Map<string, { total: number; kept: number; unit: string }>();
  const render = () => JSON.stringify({
    ok,
    truncated: true,
    hint,
    omitted: [...omitted].map(([path, o]) => `${path} 只保留了前 ${o.kept}/${o.total} ${o.unit}`),
    ...body,
  });
  for (let guard = 0; guard < 60; guard++) {
    const text = render();
    if (text.length <= maxChars) return text;
    const slot = largestShrinkable(body);
    if (!slot) return null;
    const current = slot.holder[slot.key] as unknown[] | string;
    // 按超出比例缩短，至少缩掉一项/一段，避免在大数组上一次只砍一个元素。
    const ratio = Math.max(0, (slot.size - (text.length - maxChars) - 40) / slot.size) * 0.9;
    const prev = omitted.get(slot.path);
    if (typeof current === "string") {
      const keep = Math.min(current.length - 40, Math.max(40, Math.floor(current.length * ratio)));
      if (keep >= current.length) return null;
      slot.holder[slot.key] = current.slice(0, keep) + "…";
      omitted.set(slot.path, { total: prev?.total ?? current.length, kept: keep, unit: "字" });
    } else {
      const keep = Math.min(current.length - 1, Math.max(1, Math.floor(current.length * ratio)));
      slot.holder[slot.key] = current.slice(0, keep);
      omitted.set(slot.path, { total: prev?.total ?? current.length, kept: keep, unit: "项" });
    }
  }
  return null;
}

export function textFromContent(content: unknown): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  let out = "";
  for (const part of content) {
    if (typeof part === "string") {
      out += part;
      continue;
    }
    if (!part || typeof part !== "object") continue;
    const rec = part as { type?: unknown; text?: unknown };
    const type = typeof rec.type === "string" ? rec.type : "";
    if (type && SKIP_CONTENT_TYPES.test(type)) continue;
    if (typeof rec.text === "string") out += rec.text;
  }
  return out;
}

/** Prefer delta.content; never surface reasoning_content / thinking fields. */
export function streamPieceFromChoice(choice: unknown): string {
  if (!choice || typeof choice !== "object") return "";
  const c = choice as { delta?: { content?: unknown }; message?: { content?: unknown }; text?: unknown };
  const fromDelta = textFromContent(c.delta?.content);
  if (fromDelta) return fromDelta;
  const fromMessage = textFromContent(c.message?.content);
  if (fromMessage) return fromMessage;
  return typeof c.text === "string" ? c.text : "";
}

