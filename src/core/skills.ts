/*
 * 技能库：把做过的任务沉淀为可复用的"技能"（playbook），
 * 下次遇到相似任务自动检索注入上下文，实现自我进化。
 *
 * - 任务完成后由 reflectAndLearn 调用一次轻量 LLM 复盘，决定 新建/更新/不学。
 * - 每次用户发消息时 matchSkills 做本地关键词匹配，命中的技能注入 system 上下文。
 */
import { SKILLS_STORAGE_KEY, flushUserData, schedulePersistUserData } from "./durableStore";
import { completeOnce } from "./llm";
import { AppSettings, ChatMessage, ContentPart, ToolResult } from "./types";

export interface Skill {
  id: string;
  name: string;
  /** 一句话说明什么时候用 */
  description: string;
  /** 触发关键词，用于检索 */
  triggers: string[];
  /** 工具级步骤与注意点（泛化的 playbook） */
  steps: string;
  source: "learned" | "manual";
  useCount: number;
  createdAt: number;
  updatedAt: number;
  lastUsedAt?: number;
  /**
   * 模型自己写的新技能在用户确认前不会注入上下文。
   * 复盘时读到的表格内容可能夹带指令，不审核就会被当成「经验」长期注入后续每次任务。
   */
  pending?: boolean;
  /** 模型对已确认技能提出的修改，确认后才替换现有内容 */
  proposal?: { description: string; triggers: string[]; steps: string; updatedAt: number };
}

const MAX_SKILLS = 50;
const MAX_STEPS_CHARS = 1500;

let listeners: (() => void)[] = [];

export function onSkillsChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((f) => f !== fn);
  };
}

function notify() {
  listeners.forEach((f) => f());
}

export function loadSkills(): Skill[] {
  try {
    const raw = localStorage.getItem(SKILLS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter((s) => s && s.id && s.name && s.steps);
  } catch {
    return [];
  }
}

function saveSkills(items: Skill[]): void {
  localStorage.setItem(SKILLS_STORAGE_KEY, JSON.stringify(items));
  schedulePersistUserData();
  notify();
}

function nextId(): string {
  return `skl_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;
}

function normName(name: string): string {
  return name.replace(/\s+/g, "").toLowerCase();
}

export function upsertSkill(input: {
  id?: string;
  name: string;
  description?: string;
  triggers?: string[];
  steps: string;
  source?: "learned" | "manual";
  /** 模型产出的内容：新建为待确认，修改已确认技能时只记为提议 */
  review?: boolean;
}): { skill: Skill; created: boolean } {
  const name = input.name.trim().slice(0, 40);
  const steps = input.steps.trim().slice(0, MAX_STEPS_CHARS);
  if (!name || !steps) throw new Error("技能需要 name 和 steps");
  const triggers = (input.triggers || [])
    .map((t) => String(t).trim())
    .filter(Boolean)
    .slice(0, 10);

  const list = loadSkills();
  const now = Date.now();
  let skill = input.id ? list.find((s) => s.id === input.id) : undefined;
  if (!skill) skill = list.find((s) => normName(s.name) === normName(name));

  let created = false;
  if (skill && input.review && !skill.pending) {
    skill.proposal = {
      description: (input.description || skill.description).trim().slice(0, 200),
      triggers: triggers.length ? triggers : skill.triggers,
      steps,
      updatedAt: now,
    };
    skill.updatedAt = now;
  } else if (skill) {
    if (!input.review) {
      delete skill.pending;
      delete skill.proposal;
    }
    skill.name = name;
    if (input.description) skill.description = input.description.trim().slice(0, 200);
    if (triggers.length) skill.triggers = triggers;
    skill.steps = steps;
    skill.updatedAt = now;
  } else {
    created = true;
    skill = {
      id: nextId(),
      name,
      description: (input.description || "").trim().slice(0, 200),
      triggers,
      steps,
      source: input.source || "manual",
      useCount: 0,
      createdAt: now,
      updatedAt: now,
      ...(input.review ? { pending: true } : {}),
    };
    list.unshift(skill);
  }

  saveSkills(list.sort((a, b) => b.updatedAt - a.updatedAt).slice(0, MAX_SKILLS));
  return { skill, created };
}

export function removeSkill(id: string): Skill | null {
  const list = loadSkills();
  const idx = list.findIndex((s) => s.id === id);
  if (idx < 0) return null;
  const [removed] = list.splice(idx, 1);
  saveSkills(list);
  return removed;
}

/** 确认待审核的新技能或修改提议，之后才会被检索注入 */
export function approveSkill(id: string): Skill | null {
  const list = loadSkills();
  const skill = list.find((s) => s.id === id);
  if (!skill) return null;
  if (skill.proposal) {
    skill.description = skill.proposal.description;
    skill.triggers = skill.proposal.triggers;
    skill.steps = skill.proposal.steps;
    delete skill.proposal;
  }
  delete skill.pending;
  skill.updatedAt = Date.now();
  saveSkills(list);
  return skill;
}

/** 拒绝：待审核的新技能直接删除；已确认技能只丢弃修改提议 */
export function rejectSkill(id: string): void {
  const list = loadSkills();
  const skill = list.find((s) => s.id === id);
  if (!skill) return;
  if (skill.pending) {
    removeSkill(id);
    return;
  }
  delete skill.proposal;
  saveSkills(list);
}

export function clearSkills(): void {
  saveSkills([]);
  void flushUserData();
}

export function markSkillsUsed(ids: string[]): void {
  if (!ids.length) return;
  const list = loadSkills();
  const now = Date.now();
  for (const s of list) {
    if (ids.includes(s.id)) {
      s.useCount++;
      s.lastUsedAt = now;
    }
  }
  saveSkills(list);
}

// ---------------- 检索匹配 ----------------

/** 中文按双字切分 + 英文数字按词切分 */
export function tokenize(text: string): Set<string> {
  const out = new Set<string>();
  const lower = text.toLowerCase();
  for (const m of lower.matchAll(/[a-z0-9_]+/g)) out.add(m[0]);
  const han = lower.match(/[\u4e00-\u9fff]+/g) || [];
  for (const seg of han) {
    for (let i = 0; i < seg.length; i++) {
      if (i + 1 < seg.length) out.add(seg.slice(i, i + 2));
    }
    if (seg.length === 1) out.add(seg);
  }
  return out;
}

export function overlap(a: Set<string>, b: Set<string>): number {
  let n = 0;
  for (const t of a) if (b.has(t)) n++;
  return n;
}

/** 根据用户消息匹配最相关的技能（本地检索，无 API 消耗） */
export function matchSkills(query: string, limit = 2): Skill[] {
  const skills = loadSkills().filter((s) => !s.pending);
  if (!skills.length || !query.trim()) return [];
  const qTokens = tokenize(query);
  const qLower = query.toLowerCase();

  const scored = skills
    .map((s) => {
      let score = 0;
      // 触发词整词命中，强信号
      for (const t of s.triggers) {
        const tl = t.toLowerCase();
        if (tl.length >= 2 && qLower.includes(tl)) score += 5;
      }
      score += overlap(qTokens, tokenize(s.triggers.join(" "))) * 1.5;
      score += overlap(qTokens, tokenize(s.name)) * 2;
      score += overlap(qTokens, tokenize(s.description)) * 1;
      return { s, score };
    })
    .filter((x) => x.score >= 4)
    .sort((a, b) => b.score - a.score);

  return scored.slice(0, limit).map((x) => x.s);
}

export function formatSkillsBlock(skills: Skill[]): string {
  if (!skills.length) return "";
  const lines = [
    "[相关技能] 以下是过去总结的做法，与当前任务相似时请参考其步骤执行（结合实际数据灵活调整）：",
  ];
  for (const s of skills) {
    lines.push(`### ${s.name}${s.description ? `（${s.description}）` : ""}`);
    lines.push(s.steps);
  }
  return lines.join("\n");
}

// ---------------- 模型工具 ----------------

export function toolSaveSkill(args: any): ToolResult {
  try {
    const { skill, created } = upsertSkill({
      name: String(args.name || ""),
      description: args.description ? String(args.description) : undefined,
      triggers: Array.isArray(args.triggers) ? args.triggers : [],
      steps: String(args.steps || ""),
      source: "learned",
      review: true,
    });
    return {
      ok: true,
      data: { id: skill.id, name: skill.name, created, awaiting_user_review: true },
      summary: `${created ? "新技能" : "技能修改"}「${skill.name}」已提交，待用户在技能库确认后生效`,
    };
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
}

export function toolListSkills(): ToolResult {
  const items = loadSkills();
  return {
    ok: true,
    data: {
      skills: items.map((s) => ({
        id: s.id,
        name: s.name,
        description: s.description,
        triggers: s.triggers,
        use_count: s.useCount,
        ...(s.pending ? { awaiting_user_review: true } : {}),
      })),
    },
    summary: items.length ? `技能库共 ${items.length} 个技能` : "技能库为空",
  };
}

// ---------------- 任务后反思学习 ----------------

function textOf(content: string | null | ContentPart[]): string {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .filter((p): p is Extract<ContentPart, { type: "text" }> => p.type === "text")
      .map((p) => p.text)
      .join("\n");
  }
  return "";
}

/** 把本轮任务（最后一条 user 消息之后）压缩成复盘用的文字记录 */
export function buildTaskTranscript(history: ChatMessage[]): string {
  let lastUser = -1;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i].role === "user") {
      lastUser = i;
      break;
    }
  }
  if (lastUser < 0) return "";
  const lines: string[] = [];
  for (let i = lastUser; i < history.length; i++) {
    const m = history[i];
    if (m.role === "user") {
      let t = textOf(m.content);
      const marker = t.indexOf("[用户消息]");
      if (marker >= 0) t = t.slice(marker + 6).trim();
      lines.push(`用户: ${t.slice(0, 600)}`);
    } else if (m.role === "assistant") {
      for (const tc of m.tool_calls || []) {
        lines.push(`调用工具 ${tc.function.name}(${tc.function.arguments.slice(0, 200)})`);
      }
      const t = textOf(m.content);
      if (t) lines.push(`助手: ${t.slice(0, 400)}`);
    } else if (m.role === "tool") {
      const t = textOf(m.content);
      lines.push(`工具结果: ${t.slice(0, 200)}`);
    }
  }
  let text = lines.join("\n");
  if (text.length > 6000) text = text.slice(0, 6000) + "\n…(已截断)";
  return text;
}

interface ReflectionDecision {
  action: "create" | "update" | "none";
  reason?: string;
  target_id?: string;
  skill?: { name: string; description?: string; triggers?: string[]; steps: string };
}

function parseReflection(raw: string): ReflectionDecision | null {
  let text = raw.trim();
  const fence = /```(?:json)?\s*([\s\S]*?)```/.exec(text);
  if (fence) text = fence[1].trim();
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  try {
    const obj = JSON.parse(text.slice(start, end + 1));
    if (obj && (obj.action === "create" || obj.action === "update" || obj.action === "none")) return obj;
    return null;
  } catch {
    return null;
  }
}

/**
 * 任务完成后的自我反思：判断是否值得沉淀/更新技能。
 * 返回给用户看的学习通知文本；不学习或失败时返回 null。
 */
export async function reflectAndLearn(settings: AppSettings, history: ChatMessage[]): Promise<string | null> {
  const transcript = buildTaskTranscript(history);
  if (!transcript || transcript.length < 200) return null;

  const existing = loadSkills();
  const existingBrief = existing.length
    ? existing.map((s) => `- id=${s.id} | ${s.name} | 触发词: ${s.triggers.join(",")} | ${s.description}`).join("\n")
    : "（技能库为空）";

  const prompt = `你是 Excel 智能体的复盘模块。根据下面一次已完成的任务记录，判断其中是否包含值得沉淀为「可复用技能」的通用工作流程。

[已有技能]
${existingBrief}

[本次任务记录]
${transcript}

判定规则：
- 值得沉淀：多步骤、有方法论、用户将来很可能再做类似的事（如跨表合并流程、特定清洗套路、某类汇总+图表组合）。
- 不值得（action=none）：单步读写、简单问答、失败/被用户中止的任务、纯一次性操作。
- 与已有技能高度相似 → action=update 并给出 target_id，把新经验合并进 steps。
- steps 要写成"工具级"的泛化步骤和注意点（1. 2. 3. …），不要包含具体文件名、Sheet 名、单元格值等一次性细节。
- triggers 给 3~6 个用户可能说出的关键词（中文为主）。

只输出 JSON，不要其他文字：
{"action":"create|update|none","reason":"一句话理由","target_id":"更新时给","skill":{"name":"技能名(≤12字)","description":"一句话适用场景","triggers":["…"],"steps":"1. …\\n2. …"}}`;

  try {
    // 复用主对话的地址解析与鉴权（含 OpenCode Go 代理、空 Key 不带 Authorization），并限时，避免后台复盘挂死
    const raw = await completeOnce(
      settings,
      [{ role: "user", content: prompt }],
      { temperature: 0.2, maxOutputTokens: 900, timeoutMs: 45000 }
    );
    const decision = parseReflection(raw);
    if (!decision || decision.action === "none" || !decision.skill?.name || !decision.skill?.steps) return null;

    const { skill, created } = upsertSkill({
      id: decision.action === "update" ? decision.target_id : undefined,
      name: decision.skill.name,
      description: decision.skill.description,
      triggers: decision.skill.triggers,
      steps: decision.skill.steps,
      source: "learned",
      review: true,
    });
    return created
      ? `🧠 从本次任务总结出新技能「${skill.name}」，请到技能库确认后才会在以后的任务里使用`
      : `🧠 建议更新技能「${skill.name}」，请到技能库确认修改`;
  } catch {
    return null;
  }
}
