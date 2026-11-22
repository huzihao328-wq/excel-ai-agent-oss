import { loadAllMemories, MemoryEntry } from "./memory";
import { Skill, loadSkills, overlap, tokenize } from "./skills";
import { SelectionProfile } from "./tools/excelTools";

export interface SuggestedPrompt {
  id: string;
  label: string;
  prompt: string;
  source: "selection" | "memory" | "skill" | "both";
}

const DEFAULT_PROMPTS: SuggestedPrompt[] = [
  {
    id: "d1",
    label: "按共同字段合并多张表",
    prompt: "先查看当前工作簿各工作表的表头，找出可以关联的字段，向我确认要合并的表和关联键，再把结果写入新表。",
    source: "selection",
  },
  {
    id: "d2",
    label: "分析数据，生成分类汇总",
    prompt: "先查看当前活动表的数据结构，建议合适的分类字段和统计指标，确认后在新表生成汇总。不要假定表中存在特定字段。",
    source: "selection",
  },
  {
    id: "d3",
    label: "找出重复的行并标黄",
    prompt: "找出重复的行并标黄",
    source: "selection",
  },
  {
    id: "d4",
    label: "统一日期和数字的显示格式",
    prompt: "检查当前选区的日期和数字格式，识别不一致之处并提出统一格式建议，确认后处理。保留原始数值，不要将编号误转成数字或日期。",
    source: "selection",
  },
];

/** 像「1月」「Q3」「第二季度」「2024年」「2024-03」这样的期间表头：连着出现 3 个以上多半是宽表 */
const PERIOD_HEADER = /^(\d{1,2}月份?|[一二三四五六七八九十]{1,3}月|Q[1-4]|第?[一二三四1-4]季度?|(19|20)\d{2}年?(\d{1,2}月)?|(19|20)\d{2}[-/.]\d{1,2})$/i;

function colLetter(index: number): string {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function add(
  out: SuggestedPrompt[],
  id: string,
  label: string,
  prompt: string,
  source: SuggestedPrompt["source"],
  cap = 8
) {
  if (out.length >= cap || out.some((x) => x.id === id || x.label === label)) return;
  out.push({ id, label, prompt, source });
}

function clip(text: string, n: number): string {
  const t = text.replace(/\s+/g, " ").trim();
  return t.length <= n ? t : t.slice(0, n) + "…";
}

function scoreMemory(m: MemoryEntry, profile: SelectionProfile): { score: number; hit?: string } {
  const blob = `${m.key || ""} ${m.content}`;
  let score = 0;
  let hit: string | undefined;
  for (const h of profile.headers) {
    if (h && h.length >= 2 && blob.includes(h)) {
      score += 6;
      if (!hit) hit = h;
    }
  }
  if (profile.sheet.length >= 2 && blob.includes(profile.sheet)) {
    score += 4;
    if (!hit) hit = profile.sheet;
  }
  score += overlap(tokenize([profile.sheet, ...profile.headers].join(" ")), tokenize(blob)) * 1.5;
  if (/关联|合并|订单号|join|lookup/i.test(m.content) && profile.sheetNames.length > 1) score += 3;
  if (/透视|汇总|交叉/i.test(m.content)) score += 2;
  return { score, hit };
}

function scoreSkill(s: Skill, profile: SelectionProfile, memoryText: string): number {
  const headerText = profile.headers.join(" ");
  const q = `${profile.sheet} ${headerText} ${memoryText}`;
  const qLower = q.toLowerCase();
  let score = 0;
  for (const t of s.triggers) {
    const tl = t.toLowerCase();
    if (tl.length >= 2 && qLower.includes(tl)) score += 5;
    for (const h of profile.headers) {
      if (!h || h.length < 2) continue;
      const hl = h.toLowerCase();
      if (hl.includes(tl) || (tl.includes(hl) && hl.length >= 2)) score += 3;
    }
  }
  const qTokens = tokenize(q);
  score += overlap(qTokens, tokenize(s.triggers.join(" "))) * 1.5;
  score += overlap(qTokens, tokenize(s.name)) * 2;
  score += overlap(qTokens, tokenize(s.description)) * 1;
  score += overlap(qTokens, tokenize(s.steps)) * 0.7;
  if (s.lastUsedAt && Date.now() - s.lastUsedAt < 14 * 24 * 3600 * 1000) score += 1;
  return score;
}

function memoryLines(memories: MemoryEntry[]): string {
  if (!memories.length) return "";
  return `\n同时遵守这些记忆：\n${memories.map((m) => `- ${m.content}`).join("\n")}`;
}

function skillPrompt(s: Skill, where: string, dest: string, memories: MemoryEntry[]): string {
  return `请按技能「${s.name}」处理当前选区 ${where}。${s.description ? `适用：${s.description}。` : ""}

执行流程：
${s.steps}
${memoryLines(memories)}
本次对象就是这块选区，结合实际表头和数据灵活调整。${dest}`;
}

function preferNewSheet(memories: MemoryEntry[]): boolean {
  return memories.some((m) => /新\s*Sheet|写到新|不要改原|不动原|默认新表/i.test(m.content));
}

function destHint(memories: MemoryEntry[]): string {
  return preferNewSheet(memories) ? "结果写入新 Sheet，不要改原数据。" : "默认写入新 Sheet，除非我另有说明。";
}

/** 框选数据后，结合选区画像、记忆和技能生成推荐问法；未框选则返回默认示例 */
export function buildSuggestedPrompts(profile: SelectionProfile | null, useMemory: boolean): SuggestedPrompt[] {
  if (!profile || !profile.isRange) return DEFAULT_PROMPTS;

  const memories = useMemory ? loadAllMemories() : [];
  const { sheet, address, headers, columnKinds, hasError, hasFormula, emptyRatio, sheetNames, rowCount, columnCount } =
    profile;
  const where = `${sheet}!${address}`;
  const dest = destHint(memories);
  const headerList = headers.filter(Boolean).slice(0, 8).join("、");
  const urgent: SuggestedPrompt[] = [];
  const linked: SuggestedPrompt[] = [];
  const generic: SuggestedPrompt[] = [];

  if (hasError) {
    add(
      urgent,
      "fix-err",
      "修复选区里的公式错误",
      `当前选区 ${where} 里有公式错误。请先 inspect_range 查看错误类型，找出原因并直接修好（优先原地修正，修之前用一两句话说明）。`,
      "selection"
    );
  }
  if (hasFormula && !hasError) {
    add(
      generic,
      "explain-fx",
      "解释选区里的公式",
      `请查看当前选区 ${where} 的公式，解释它们在算什么、引用了哪些单元格、有没有风险。`,
      "selection"
    );
  }

  const relatedMemories = memories
    .map((m) => ({ m, ...scoreMemory(m, profile) }))
    .filter((x) => x.score >= 4)
    .sort((a, b) => b.score - a.score)
    .slice(0, 3);

  const memoryText = relatedMemories.map((x) => `${x.m.key || ""} ${x.m.content}`).join(" ");
  const relatedSkills = loadSkills()
    .map((s) => ({ s, score: scoreSkill(s, profile, memoryText) }))
    .filter((x) => x.score >= 3)
    .sort((a, b) => b.score - a.score)
    .slice(0, 2);

  const citedMemories = relatedMemories.map((x) => x.m);

  if (relatedSkills.length && citedMemories.length) {
    const best = relatedSkills[0].s;
    add(
      linked,
      `both-${best.id}`,
      `用技能「${best.name}」（含记忆）`,
      skillPrompt(best, where, dest, citedMemories),
      "both"
    );
    for (const { s } of relatedSkills.slice(1)) {
      add(linked, `skill-${s.id}`, `用技能「${s.name}」`, skillPrompt(s, where, dest, citedMemories), "skill");
    }
  } else if (relatedSkills.length) {
    for (const { s } of relatedSkills) {
      add(linked, `skill-${s.id}`, `用技能「${s.name}」`, skillPrompt(s, where, dest, []), "skill");
    }
  } else {
    for (const { m, hit } of relatedMemories) {
      add(
        linked,
        `mem-${m.id}`,
        hit ? `按记忆处理「${hit}」` : `按记忆：${clip(m.content, 16)}`,
        `当前选区是 ${where}${hit ? `，表头包含「${hit}」` : ""}。请遵守这条记忆：${m.content}。据此处理这块数据。${dest}`,
        "memory"
      );
    }
    if (/关联|合并|订单号|join|lookup/i.test(memories.map((m) => m.content).join(" ")) && sheetNames.length > 1) {
      const joinMem = memories.find((m) => /关联|合并|订单号|join|lookup/i.test(m.content));
      if (joinMem && !linked.some((x) => x.id === `mem-${joinMem.id}`)) {
        add(
          linked,
          `mem-join-${joinMem.id}`,
          "按记忆跨表合并",
          `当前选区 ${where}。请遵守记忆「${joinMem.content}」，把相关 Sheet 合并到新表。`,
          "memory"
        );
      }
    }
  }

  const dateIdx = columnKinds.findIndex((k) => k === "date");
  if (dateIdx >= 0) {
    const name = headers[dateIdx] || `${colLetter(dateIdx)} 列`;
    add(
      generic,
      "dates",
      `把「${name}」日期统一格式`,
      `当前选区 ${where} 里「${name}」看起来是日期。请用 clean_column（to_date）把该列统一成 yyyy-mm-dd，${dest}`,
      "selection"
    );
  }

  const idHeader = headers.find((h) => /身份证|证件号/.test(h));
  if (idHeader) {
    add(
      generic,
      "id-card",
      `从「${idHeader}」算出生日期/性别/年龄`,
      `当前选区 ${where} 里「${idHeader}」是身份证号。请用 extract_text 在右侧新增出生日期、性别、年龄和身份证校验四列，并汇报校验不通过的个数。`,
      "selection"
    );
  }

  const periodHeaders = headers.filter((h) => PERIOD_HEADER.test(h.trim()));
  if (periodHeaders.length >= 3) {
    add(
      generic,
      "unpivot",
      "把横排的期间列转成长表",
      `当前选区 ${where} 的「${periodHeaders.slice(0, 3).join("、")}」等 ${periodHeaders.length} 列是按期间横着排的。请用 unpivot_columns 把它们转成「期间 + 值」两列的长表放到新 Sheet，方便汇总和做透视。`,
      "selection"
    );
  }

  const textCols = headers.filter((_, i) => columnKinds[i] === "text");
  const numCols = headers.filter((_, i) => columnKinds[i] === "number");
  if (textCols.length && numCols.length) {
    add(
      generic,
      "summarize",
      `按「${textCols[0]}」汇总「${numCols[0]}」`,
      `对当前选区 ${where}（表头：${headerList}）按「${textCols[0]}」汇总「${numCols[0]}」的合计，做成汇总表。${dest}`,
      "selection"
    );
  }

  if (emptyRatio >= 0.15) {
    add(
      generic,
      "clean",
      "清洗选区（空值/重复）",
      `检查当前选区 ${where} 的空值、重复行和格式问题，清洗后${preferNewSheet(memories) ? "写入新 Sheet" : "写入新 Sheet，保留原数据"}。`,
      "selection"
    );
  }

  if (columnCount === 1 && rowCount >= 8) {
    add(
      generic,
      "dups",
      "找出重复值并标黄",
      `当前选区 ${where} 是一列数据。请找出重复项并标黄，再告诉我重复了多少。`,
      "selection"
    );
  }

  if (columnCount >= 2 && rowCount >= 2) {
    add(
      generic,
      "copy",
      "把选区复制到新表",
      `请用 copy_range 把当前选区 ${where} 复制到一个新 Sheet（先 create_sheet），不要把数据读出来再写入。`,
      "selection"
    );
  }
  add(
    generic,
    "analyze",
    "分析这块选区",
    `分析当前选区 ${where}${headerList ? `（表头：${headerList}）` : ""}：概述结构，给出关键统计和你发现的规律或异常。`,
    "selection"
  );
  add(
    generic,
    "chart",
    "根据选区做图表",
    `根据当前选区 ${where} 选择最合适的图表类型画出来，并说明理由。`,
    "selection"
  );
  add(
    generic,
    "pivot",
    "把选区做成透视表",
    `根据当前选区 ${where}${headerList ? `（表头：${headerList}）` : ""} 做一张透视表，放到新 Sheet。行/列/值按数据含义选择。`,
    "selection"
  );
  add(
    generic,
    "table",
    "转成 Excel 表并加筛选",
    `把当前选区 ${where} 转成 Excel 表并加上自动筛选。`,
    "selection"
  );

  const linkedCap = linked.length ? Math.min(3, linked.length) : 0;
  return [...urgent, ...linked.slice(0, linkedCap), ...generic].slice(0, linkedCap ? 5 : 4);
}
