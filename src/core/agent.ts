import { AttachedImage } from "./images";
import { streamChatCompletion, AssistantTurn, probeContextWindowInfo } from "./llm";
import { clipJsonToBudget } from "./streamText";
import { formatArtifactsBlock } from "./artifacts";
import { formatMemoryFile, loadAllMemories } from "./memory";
import { buildContextBlock, buildRuntimeBlock, buildSystemPrompt } from "./prompts";
import { TOOL_DEFINITIONS, TOOL_LABELS } from "./tools/definitions";
import { beginCheckpointSession } from "./tools/checkpoint";
import { executeTool, getQuickContext, TOOL_REGISTRY, ToolExecContext } from "./tools/excelTools";
import { getExternalFiles, getExternalSelection } from "./tools/externalFiles";
import { formatSkillsBlock, markSkillsUsed, matchSkills, reflectAndLearn } from "./skills";
import { looksLikePlanOnly, suggestToolName } from "./toolCallText";
import { compactHistoryForRequest, estimateTokens, estimateMessagesTokens, getCachedContextWindow, guessContextWindow } from "./contextWindow";
import { adaptToolsForEndpoint, compactToolDefinitions } from "./toolSchema";
import { normalizeToolArguments, parseToolArgumentsJson } from "./toolArguments";
import { effectiveMaxOutputTokens, recordPromptUsage, tokenEstimateRatio } from "./modelCapabilities";
import { AppSettings, ChatMessage, ContentPart, ToolCall, ToolDefinition, ToolResult } from "./types";

/** 单条工具结果的字符上限（小窗口）；大窗口按比例放宽到 CEILING */
const MAX_TOOL_RESULT_CHARS = 20000;
const MAX_TOOL_RESULT_CHARS_CEILING = 60000;
const isMemoryTool = (name: string) => !!TOOL_REGISTRY[name]?.memory;
/** 探测不到上下文窗口时的保守假设 */
const FALLBACK_CONTEXT_TOKENS = 64_000;
/** 给模型输出预留的 token（未设置 max_tokens 时） */
const DEFAULT_OUTPUT_RESERVE = 4096;
/** 请求最多用到窗口的这个比例，给估算误差留余量 */
const WINDOW_SHARE = 0.92;
/** 系统提示 + 工具定义超过可用输入预算的这个比例时，改用精简工具说明 */
const COMPACT_TOOLS_SHARE = 0.4;

export interface AgentCallbacks {
  /** assistant 文本流式更新 */
  onAssistantDelta: (fullText: string) => void;
  /**
   * 一轮 assistant 文本结束。interim=true 表示这段文字只是调用工具前的过程叙述，
   * 不是最终结论（界面应弱化显示；fullText 为空时应移除气泡）。
   */
  onAssistantDone: (fullText: string, meta?: { interim?: boolean }) => void;
  /** 工具开始执行 */
  onToolStart: (callId: string, name: string, label: string, argsSummary: string) => void;
  /** 工具执行结束 */
  onToolEnd: (callId: string, result: ToolResult) => void;
  /** 命中技能库并注入上下文 */
  onSkillsMatched?: (names: string[]) => void;
  /** 任务后反思学到/更新了技能 */
  onLearned?: (message: string) => void;
  /** 需要让用户知道的系统提示（输出被截断、上下文被压缩等） */
  onNotice?: (text: string) => void;
  /** 用户在任务进行中补充的说明：是否有还没送达的 */
  hasSteering?: () => boolean;
  /** 取走还没送达的补充说明；取走即视为已送达 */
  takeSteering?: () => string[];
  /** 每一轮模型调用开始时的进度（从 1 计） */
  onProgress?: (info: {
    round: number;
    maxRounds: number;
    phase: "model" | "tools" | "retry";
    retryAttempt?: number;
    retryMax?: number;
  }) => void;
}

/** 卡片摘要里出现的参数及其叫法；没列在这里的参数不上摘要（展开卡片仍能看到完整结果） */
const ARG_LABELS: Record<string, string> = {
  name: "名称", new_name: "新名称", title: "标题", columns: "列", column: "列", count: "数量", rows: "行", key: "键", file: "文件",
  text: "文本", formula: "公式", what: "内容", chart_type: "图表", type: "类型", sort_by: "排序", filter: "筛选",
  start_row: "起始行", end_row: "结束行", before_row: "插在第几行前", start_column: "起始列", before_column: "插在哪列前",
  keys: "按", key_column: "按", group_by: "分组", metrics: "指标", date_column: "日期列", date_part: "分段",
  find: "查找", replace: "替换为", lookup_sheet: "查找表", lookup_key: "匹配列", dest_key: "匹配列", return_columns: "取回", return_column: "取回",
  dest_start_column: "写到", dest_column: "写到", dest_header: "新列", dest_headers: "新列", source_column: "按列",
  operations: "操作", operation: "操作", targets: "提取", mode: "方式", rule: "规则", delimiter: "分隔符", separator: "连接符",
  left_sheet: "左表", right_sheet: "右表", old_sheet: "旧版", new_sheet: "新版", sheets: "表", id_columns: "标识列", value_columns: "值列",
  number_format: "数字格式", format: "格式", filename: "文件名", orientation: "方向", paper_size: "纸张", data_range: "数据",
  content: "内容", scope: "范围", question: "问题", cell: "位置", target: "对象", purpose: "用途", max_rows: "最多行数",
};

/** 参数值写成一小段人能读的文字；对象数组取每项的列名/名称，纯对象不上摘要 */
function plainArg(value: unknown): string {
  if (typeof value === "string") return value.trim();
  if (typeof value === "number") return String(value);
  if (!Array.isArray(value)) return "";
  const names = value
    .map((item) => {
      if (typeof item === "string" || typeof item === "number") return String(item);
      if (!item || typeof item !== "object") return "";
      const record = item as Record<string, unknown>;
      const named = ["column", "field", "name", "label", "header", "sheet", "op", "type"].map((k) => record[k]).find((v) => typeof v === "string" && v);
      return typeof named === "string" ? named : "";
    })
    .filter(Boolean);
  return names.length ? names.join("、") : value.length ? `${value.length} 项` : "";
}

/** 工具卡片上的一行摘要：用「Sheet!区域 · 写入 3×4」这样的人话，而不是原始参数名 */
export function summarizeArgs(argsJson: string): string {
  try {
    const args = JSON.parse(argsJson || "{}") as Record<string, unknown>;
    const used = new Set<string>();
    const pick = (...keys: string[]) => {
      for (const key of keys) {
        const value = args[key];
        if (typeof value === "string" && value.trim()) { used.add(key); return value.trim(); }
      }
      return "";
    };
    const ref = (sheet: string, range: string) => (sheet && range ? `${sheet}!${range}` : sheet || range);
    const parts: string[] = [];
    const main = ref(pick("sheet"), pick("range", "address", "start_cell", "target_cell"));
    if (main) parts.push(main);
    const source = ref(pick("source_sheet"), pick("source_range"));
    if (source) parts.push(`来源 ${source}`);
    const dest = ref(pick("dest_sheet", "target_sheet", "target_sheet_name"), pick("dest_cell", "dest_range"));
    if (dest) parts.push(`→ ${dest}`);
    for (const key of ["values", "formulas"]) {
      const grid = args[key];
      if (!Array.isArray(grid)) continue;
      used.add(key);
      parts.push(Array.isArray(grid[0])
        ? `写入${key === "formulas" ? "公式 " : " "}${grid.length}×${(grid[0] as unknown[]).length}`
        : `${grid.length} 项`);
    }
    if (typeof args.code === "string") { used.add("code"); parts.push(`脚本 ${args.code.length} 字符`); }
    for (const [key, value] of Object.entries(args)) {
      if (parts.length >= 3) break;
      const label = ARG_LABELS[key];
      if (used.has(key) || !label) continue;
      const text = plainArg(value);
      if (!text) continue;
      parts.push(`${label} ${text.length > 32 ? text.slice(0, 32) + "…" : text}`);
    }
    return parts.join(" · ");
  } catch {
    return argsJson.slice(0, 80);
  }
}

function buildUserMessage(text: string, images?: AttachedImage[]): ChatMessage {
  if (!images?.length) return { role: "user", content: text };
  const parts: ContentPart[] = [{ type: "text", text }];
  for (const img of images) {
    parts.push({ type: "image_url", image_url: { url: img.dataUrl } });
  }
  return { role: "user", content: parts };
}

/** 把实时上下文拼到最后一条 user 消息前面，只影响本次请求，不改历史。 */
function withTurnContext(messages: ChatMessage[], prefix: string): ChatMessage[] {
  if (!prefix) return messages;
  let idx = -1;
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i].role === "user") {
      idx = i;
      break;
    }
  }
  if (idx < 0) return messages;
  const target = messages[idx];
  const wrap = (text: string) => `${prefix}\n\n[用户消息]\n${text}`;
  let content: ChatMessage["content"];
  if (typeof target.content === "string" || target.content == null) {
    content = wrap(target.content ?? "");
  } else {
    let replaced = false;
    content = target.content.map((part) => {
      if (!replaced && part.type === "text") {
        replaced = true;
        return { ...part, text: wrap(part.text) };
      }
      return part;
    });
    if (!replaced) content = [{ type: "text", text: wrap("") }, ...target.content];
  }
  const out = messages.slice();
  out[idx] = { ...target, content };
  return out;
}

export interface AgentOutcome {
  /** 是否因达到工具调用轮数上限而结束（任务可能未完成） */
  hitRoundLimit: boolean;
  incomplete?: boolean;
}

/**
 * Agent 主循环：发送消息 → 若模型请求工具则执行并回填 → 直到模型给出最终回复。
 * history 会被原地追加（包含 assistant/tool 消息），供下一轮对话复用。
 * 历史里只存用户原话；工作簿状态、记忆、技能等实时上下文在每次请求时临时拼接，不会随对话累积。
 */
export async function runAgent(
  userText: string,
  history: ChatMessage[],
  settings: AppSettings,
  execCtx: ToolExecContext,
  callbacks: AgentCallbacks,
  signal: AbortSignal,
  images?: AttachedImage[],
  options?: { skipCheckpointSession?: boolean; skipUserPush?: boolean }
): Promise<AgentOutcome> {
  execCtx = { ...execCtx, settings, signal };
  if (settings.autoCheckpoint !== false && !options?.skipCheckpointSession) beginCheckpointSession(userText);

  // 技能检索：本地匹配相似任务的历史经验，注入上下文
  let skillsBlock = "";
  try {
    const matched = matchSkills(userText);
    if (matched.length) {
      skillsBlock = formatSkillsBlock(matched);
      markSkillsUsed(matched.map((s) => s.id));
      callbacks.onSkillsMatched?.(matched.map((s) => s.name));
    }
  } catch {
    /* 技能检索失败不影响任务 */
  }

  let detectedWindow = getCachedContextWindow(settings.baseUrl, settings.model, settings.apiKey);
  if (!settings.contextWindowTokens && !detectedWindow) {
    // 长时间使用后缓存会过期；发新任务时重新读取部署信息，不能悄悄回到过大的模型名估算。
    detectedWindow = (await probeContextWindowInfo(settings, { signal }).catch(() => null))?.tokens;
  }
  if (signal.aborted) return { hitRoundLimit: false };
  let windowTokens =
    settings.contextWindowTokens ||
    detectedWindow ||
    guessContextWindow(settings.model) ||
    FALLBACK_CONTEXT_TOKENS;

  // 附带轻量实时上下文，让模型少走一步
  let workbookBlock = "";
  try {
    const ctx = await getQuickContext(settings.sendSelectionContext, true);
    const externalFiles = getExternalFiles().map(
      (f) => `${f.caption}（${f.sheets.map((s) => `${s.name} ${s.rows.length}行`).join("、")}）`
    );
    const extSel = getExternalSelection();
    const extSelFile = extSel ? getExternalFiles().find((f) => f.name === extSel.file) : undefined;
    workbookBlock = buildContextBlock({
      ...ctx,
      selectionPreview: settings.sendSelectionContext ? ctx.selectionPreview : undefined,
      selectionFormula: settings.sendSelectionContext ? ctx.selectionFormula : undefined,
      selectionErrors: settings.sendSelectionContext ? ctx.selectionErrors : undefined,
      externalFiles,
      externalSelection: extSel
        ? `${extSelFile?.caption ?? `文件「${extSel.file}」`} 工作表「${extSel.sheet}」区域 ${extSel.range}`
        : undefined,
      memoryFile: settings.useMemory ? formatMemoryFile(loadAllMemories()) : undefined,
      skillsBlock: skillsBlock || undefined,
      artifactsBlock: formatArtifactsBlock() || undefined,
    });
  } catch {
    /* Excel 不可用时忽略 */
  }
  if (!workbookBlock && skillsBlock) workbookBlock = skillsBlock;
  const contextBlock = [
    buildRuntimeBlock({ maxToolRounds: settings.maxToolRounds, toolResultChars: maxToolResultChars(windowTokens) }),
    workbookBlock,
  ].filter(Boolean).join("\n\n");

  let text = userText;
  if (images?.length) {
    text += `\n\n[用户附加了 ${images.length} 张图片: ${images.map((i) => i.name).join("、")}]`;
  }
  // 上一次请求失败后重试：同一条用户消息已经在历史末尾，不要再推一份
  const last = history[history.length - 1];
  const sameAsLast = last?.role === "user" && userTextOf(last) === text;
  if (!options?.skipUserPush && !sameAsLast) history.push(buildUserMessage(text, images));

  // 统计实质性工具调用（用于判断任务是否值得复盘学习）；meta 工具不算
  let substantiveCalls = 0;
  const availableTools = settings.useMemory
    ? TOOL_DEFINITIONS
    : TOOL_DEFINITIONS.filter((tool) => !isMemoryTool(tool.function.name));
  const enabledTools = availableTools.map((tool) => ({
    ...tool, function: { ...tool.function, parameters: { ...tool.function.parameters, additionalProperties: false } },
  }));
  const enabledToolNames = enabledTools.map((t) => t.function.name);
  const systemPrompt = buildSystemPrompt(settings.askDirection, settings.useMemory);

  // 预算都按本地估算的 token 算；ratio 是接口报告的真实用量 ÷ 估算，用来把真实窗口换算成估算单位。
  const outputReserve = effectiveMaxOutputTokens(settings) || DEFAULT_OUTPUT_RESERVE;
  const inputBudget = () =>
    Math.floor((windowTokens * WINDOW_SHARE - outputReserve) / tokenEstimateRatio(settings.baseUrl, settings.model));
  const toolTokens = (list: ToolDefinition[]) => estimateTokens(JSON.stringify(adaptToolsForEndpoint(list, settings.baseUrl)));
  const promptTokens = estimateTokens(systemPrompt) + estimateTokens(contextBlock);
  let requestTools: ToolDefinition[] = enabledTools;
  let requestToolTokens = toolTokens(enabledTools);
  if (promptTokens + requestToolTokens > inputBudget() * COMPACT_TOOLS_SHARE) {
    // 小窗口模型：工具说明精简到第一句，名称/参数结构不变，给对话和工具结果腾地方。
    requestTools = compactToolDefinitions(enabledTools);
    requestToolTokens = toolTokens(requestTools);
  }
  const fixedTokens = promptTokens + requestToolTokens;
  /** 上下文超限后收紧的上限；之后校准也不会再放宽到超过它 */
  let budgetCap = Infinity;
  const currentHistoryBudget = () => Math.min(budgetCap, inputBudget() - fixedTokens);
  let historyBudget = currentHistoryBudget();
  let compactNoticeShown = false;
  const failedCalls = new Map<string, number>();
  /** 自上次改表以来，只读调用的签名 → 成功次数；同样的读取第三次时提醒模型推进 */
  const readCounts = new Map<string, number>();

  const maybeLearn = () => {
    if (!settings.autoLearnSkills || signal.aborted || substantiveCalls < 2) return;
    // 后台复盘，不阻塞回复展示；失败静默
    void reflectAndLearn(settings, history.slice())
      .then((msg) => {
        if (msg) callbacks.onLearned?.(msg);
      })
      .catch(() => undefined);
  };

  /** 下一次请求临时附带的提示（纠偏、轮次提醒），不进历史 */
  let nudgeMessages: ChatMessage[] = [];
  let nudged = false;
  const maxRounds = settings.maxToolRounds;

  // round === maxRounds 是额外的收尾请求：工具额度已用完，只让模型总结进度。
  for (let round = 0; round <= maxRounds; round++) {
    if (signal.aborted) return { hitRoundLimit: false };
    const wrapUp = round === maxRounds;
    callbacks.onProgress?.({ round: Math.min(round + 1, maxRounds), maxRounds, phase: "model" });

    // 任务进行中用户插入的补充说明：放在上一轮工具结果之后，本轮模型就能看到
    const steering = wrapUp ? [] : callbacks.takeSteering?.() ?? [];
    if (steering.length) {
      history.push({ role: "user", content: formatSteering(steering) });
      failedCalls.clear();
      // 旧计划的催执行提示不能排在用户的新要求之后。
      nudgeMessages = [];
    }
    const remaining = maxRounds - round;
    if (wrapUp) {
      nudgeMessages = [{ role: "user", content: WRAP_UP_NOTE }];
    } else if (round > 0 && maxRounds >= 5 && remaining <= 2 && history[history.length - 1]?.role !== "user") {
      nudgeMessages = appendNote(nudgeMessages, roundBudgetNote(remaining));
    }

    const requestHistory = history;
    const turnNudges = nudgeMessages;
    nudgeMessages = [];
    const buildMessages = (budget: number): ChatMessage[] => {
      if (budget < 256) {
        throw new Error(`当前上下文窗口（${windowTokens.toLocaleString()} tokens）不足以容纳工具定义和输出预留。请核对窗口配置、降低单次回复上限，或换用更大窗口的模型。`);
      }
      const compacted = compactHistoryForRequest(requestHistory, { budgetTokens: budget - estimateMessagesTokens(turnNudges) });
      if (estimateMessagesTokens(compacted) + estimateMessagesTokens(turnNudges) > budget) {
        throw new Error("压缩历史后，当前任务和最近工具结果仍超过模型上下文预算。请缩小读取区域、减少附件内容，或核对上下文窗口设置后重试；已完成的操作会保留。");
      }
      if (compacted !== requestHistory && !compactNoticeShown) {
        compactNoticeShown = true;
        callbacks.onNotice?.("对话较长，已自动省略较早的工具结果以腾出上下文；如需旧数据模型会重新读取。");
      }
      return [
        { role: "system", content: systemPrompt },
        ...withTurnContext(compacted, contextBlock),
        ...turnNudges,
      ];
    };
    let messages = buildMessages(historyBudget);

    const requestTurn = () => streamChatCompletion(
      settings,
      messages,
      requestTools,
      {
        onTextDelta: (_d, full) => callbacks.onAssistantDelta(full),
        onAttempt: () => {
          callbacks.onProgress?.({ round: Math.min(round + 1, maxRounds), maxRounds, phase: "model" });
        },
        onRetryWait: (attempt, maxAttempts) => {
          callbacks.onProgress?.({
            round: Math.min(round + 1, maxRounds),
            maxRounds,
            phase: "retry",
            retryAttempt: attempt,
            retryMax: maxAttempts,
          });
        },
      },
      signal
    );
    let turn: AssistantTurn;
    try {
      turn = await requestTurn();
    } catch (e) {
      const error = e as { contextOverflow?: boolean; contextLimit?: number } | null;
      // 收尾总结失败不该让整个任务报错：已完成的操作都在，退回固定说明即可。
      if (wrapUp && !signal.aborted) return finishAtRoundLimit(history, callbacks);
      if (!error?.contextOverflow || signal.aborted) throw e;
      if (error.contextLimit && error.contextLimit < windowTokens) windowTokens = error.contextLimit;
      const smallerBudget = Math.min(Math.floor(historyBudget * 0.65), currentHistoryBudget());
      const smallerMessages = buildMessages(smallerBudget);
      // 只在请求确实缩小时重试一次；错误发生在 HTTP 层，没有执行新工具。
      if (estimateMessagesTokens(smallerMessages) >= estimateMessagesTokens(messages)) throw e;
      budgetCap = smallerBudget;
      historyBudget = smallerBudget;
      messages = smallerMessages;
      callbacks.onNotice?.("接口报告上下文超限，已进一步压缩历史并重试；已完成的工具不会重复执行。");
      turn = await requestTurn();
    }

    // 用接口报告的真实用量校准本地估算，后续轮次按校准后的预算压缩历史。
    if (turn.usage?.promptTokens && !hasImages(messages)) {
      recordPromptUsage(settings.baseUrl, settings.model, estimateMessagesTokens(messages) + requestToolTokens, turn.usage.promptTokens);
      historyBudget = currentHistoryBudget();
    }

    if (wrapUp) {
      // 收尾请求里模型仍想调用工具：不执行，改用固定说明，历史里不留悬空的 tool_calls。
      if (turn.toolCalls.length || !turn.content.trim()) return finishAtRoundLimit(history, callbacks);
      history.push({ role: "assistant", content: turn.content });
      callbacks.onAssistantDone(turn.content);
      return { hitRoundLimit: true, incomplete: true };
    }

    if (!turn.toolCalls.length) {
      // 模型只描述了计划没动手：不进历史，弱化展示，下一轮提醒它直接调用工具（每次任务最多一次）
      if (!turn.truncated && !nudged && looksLikePlanOnly(turn.content, enabledToolNames)) {
        nudged = true;
        callbacks.onAssistantDone(turn.content, { interim: true });
        nudgeMessages = [
          { role: "assistant", content: turn.content },
          {
            role: "user",
            content:
              "[系统提示] 你上一条回复只描述了计划，没有调用任何工具，用户看不到进展。" +
              "现在请直接调用工具执行（一次可以并行发出多个 tool_calls），不要再复述计划；" +
              "如果确实需要用户补充信息，用一句话提问即可。",
          },
        ];
        continue;
      }
      history.push({ role: "assistant", content: turn.content || null });
      if (turn.content) callbacks.onAssistantDone(turn.content);
      // 模型收尾时用户刚好插了话：不结束任务，下一轮带上补充说明继续
      if (!turn.truncated && callbacks.hasSteering?.()) continue;
      if (turn.truncated) {
        callbacks.onNotice?.("这条回复达到了单次输出 token 上限被截断。可以让我「继续」，或在设置里调大「单次回复 token 上限」。");
      }
      if (!turn.truncated) maybeLearn();
      return { hitRoundLimit: false, incomplete: turn.truncated || looksLikePlanOnly(turn.content, enabledToolNames) };
    }

    // 参数先规范化再进历史：能安全修好的 JSON / 字段名问题就地修好，
    // 后续请求里的历史参数都是合法 JSON，模型也能从自己的历史里看到正确写法。
    const prepared = turn.toolCalls.map((call) => prepareToolCall(call, turn.truncated));
    const toolCalls = prepared.map((p) => p.call);
    history.push({
      role: "assistant",
      content: turn.content || null,
      tool_calls: toolCalls,
    });
    // 不管是用户中止还是界面回调抛错，都要给剩余的 tool_calls 补齐结果消息，
    // 否则悬空的 tool_calls 会让下一次 API 请求被拒绝（400），整个会话报废。
    const backfill = (from: number, reason: string) => {
      for (let j = from; j < toolCalls.length; j++) {
        history.push({
          role: "tool",
          content: JSON.stringify({ ok: false, error: reason }),
          tool_call_id: toolCalls[j].id,
        });
      }
    };

    // 为下一轮输入留空间；单个结果不应独占历史预算，超出时按结构裁剪并明确标记。
    const resultChars = Math.min(
      maxToolResultChars(windowTokens),
      Math.max(800, Math.floor(historyBudget * 0.45 / toolCalls.length))
    );
    let i = 0;
    /** 本批第一个失败的工具；之后会改表的调用跳过，只读调用照常执行 */
    let blockedBy = "";
    try {
      // 这些 UI 回调也可能抛错，必须位于补全历史的保护范围内。
      callbacks.onAssistantDone(turn.content, { interim: true });
      callbacks.onProgress?.({ round: round + 1, maxRounds, phase: "tools" });
      for (; i < toolCalls.length;) {
        const call = toolCalls[i];
        if (signal.aborted) {
          backfill(i, "用户中止了任务，该工具未执行。");
          return { hitRoundLimit: false };
        }
        if (callbacks.hasSteering?.()) {
          // 当前工具完成后先把补充交给模型；同批旧计划里的剩余操作不能继续改表。
          backfill(i, "用户补充了新的要求，该工具未执行。请先阅读补充说明，再决定后续操作。");
          break;
        }
        const name = call.function.name;
        const spec = TOOL_REGISTRY[name];
        if (blockedBy && !spec?.readOnly) {
          history.push({
            role: "tool",
            content: JSON.stringify({
              ok: false,
              error: `本批前面的 ${blockedBy} 失败，本工具未执行。请先根据失败原因调整方案，再决定后续调用；不要重复已成功的操作。`,
            }),
            tool_call_id: call.id,
          });
          i++;
          continue;
        }
        const label = TOOL_LABELS[name] || name;
        callbacks.onToolStart(call.id, name, label, summarizeArgs(call.function.arguments));

        if (!spec?.meta) substantiveCalls++;
        let result: ToolResult;
        const signature = toolSignature(name, prepared[i].signatureArgs);
        if ((failedCalls.get(signature) || 0) >= 2) {
          result = { ok: false, error: "相同工具和参数已失败两次，本次未执行。请先读取相关范围核对状态、修正参数或换用其他工具，不要重复原调用。" };
        } else if (!enabledToolNames.includes(name)) {
          const guess = suggestToolName(name, enabledToolNames);
          result = {
            ok: false,
            error: `未知工具 "${name}"。${guess ? `你可能想调用 ${guess}。` : ""}只能使用工具列表里给出的名称。`,
          };
        } else if (!settings.useMemory && isMemoryTool(name)) {
          result = { ok: false, error: "长期记忆已关闭，本次不会读取或写入记忆。" };
        } else if (prepared[i].error) {
          result = { ok: false, error: prepared[i].error };
        } else {
          try {
            result = await executeTool(name, call.function.arguments, execCtx);
          } catch (e: any) {
            result = { ok: false, error: `执行 ${name} 出错: ${e?.message || String(e)}` };
          }
        }
        const mutated = typeof spec?.mutating === "function" ? spec.mutating(safeParse(call.function.arguments)) : !!spec?.mutating;
        if (result.ok && mutated) {
          failedCalls.clear();
          readCounts.clear();
        } else if (result.ok) failedCalls.delete(signature);
        else failedCalls.set(signature, (failedCalls.get(signature) || 0) + 1);

        let payload: Record<string, unknown> = result.ok ? { ok: true, data: result.data } : { ok: false, error: result.error };
        if (result.ok && spec?.readOnly) {
          const seen = (readCounts.get(signature) || 0) + 1;
          readCounts.set(signature, seen);
          if (seen >= 3) payload = { ...payload, note: REPEATED_READ_NOTE };
        }
        const content = clipJsonToBudget(payload, resultChars);
        history.push({ role: "tool", content, tool_call_id: call.id });
        // 先记录真实执行结果再更新 UI，避免 UI 出错后把成功写入误报成未执行并重复写入。
        i++;
        callbacks.onToolEnd(call.id, result);
        if (!result.ok && !blockedBy) blockedBy = name;
      }
    } catch (e) {
      backfill(i, "工具执行过程中断，该工具未执行。");
      throw e;
    }
  }

  // 正常流程在收尾请求里就返回了；这里只是兜底。
  return finishAtRoundLimit(history, callbacks);
}

const LIMIT_MESSAGE = "已达到单次任务的工具调用轮数上限，任务可能未完成。你可以对我说“继续”让我接着做。";

const WRAP_UP_NOTE =
  "[系统提示] 本次任务的工具调用轮次已经用完，之后发出的工具调用都不会执行。" +
  "请不要再调用工具，直接用 2～4 句话告诉用户：已经完成了哪些操作、结果在哪个 Sheet 的哪个区域；还有哪些步骤没做；" +
  "最后提示用户可以说「继续」让你接着完成。只陈述工具结果里确认过的事实。";

const REPEATED_READ_NOTE =
  "本任务中已用完全相同的参数成功读取过至少两次，期间你没有修改工作簿，结果与之前相同。请直接基于已有结果推进下一步，不要再重复这次读取。";

function roundBudgetNote(remaining: number): string {
  return remaining <= 1
    ? "[系统提示] 这是本次任务最后一轮可以执行工具的机会。请只做最关键的一步；做不完的部分会在下一次请求里总结给用户。"
    : `[系统提示] 本次任务（含本轮）还剩 ${remaining} 轮工具调用。请优先完成最关键的写入，能并行的调用放在同一轮发出。`;
}

/** 追加一条临时提示；上一条临时消息已是 user 时合并，避免连续两条 user 消息（部分接口会拒绝）。 */
function appendNote(nudges: ChatMessage[], note: string): ChatMessage[] {
  const lastNudge = nudges[nudges.length - 1];
  if (lastNudge?.role === "user" && typeof lastNudge.content === "string") {
    return [...nudges.slice(0, -1), { ...lastNudge, content: `${lastNudge.content}\n${note}` }];
  }
  return [...nudges, { role: "user", content: note }];
}

function finishAtRoundLimit(history: ChatMessage[], callbacks: AgentCallbacks): AgentOutcome {
  history.push({ role: "assistant", content: LIMIT_MESSAGE });
  callbacks.onAssistantDone(LIMIT_MESSAGE);
  return { hitRoundLimit: true, incomplete: true };
}

/** 单条工具结果的字符上限：大窗口模型可以一次看完整页数据，小窗口保持 2 万字符。 */
function maxToolResultChars(windowTokens: number): number {
  return Math.min(MAX_TOOL_RESULT_CHARS_CEILING, Math.max(MAX_TOOL_RESULT_CHARS, Math.floor(windowTokens * 0.12)));
}

interface PreparedToolCall {
  /** 参数已规范化的调用，写进历史并交给执行器 */
  call: ToolCall;
  /** 用于识别重复失败的参数文本（解析失败时用原文，避免不同的坏参数被当成同一个） */
  signatureArgs: string;
  /** 参数无法使用时的错误，不会执行 */
  error?: string;
}

function prepareToolCall(call: ToolCall, truncated: boolean): PreparedToolCall {
  const raw = call.function.arguments ?? "";
  const withArgs = (args: string): ToolCall => (args === raw ? call : { ...call, function: { ...call.function, arguments: args } });
  if (truncated && !isJsonObject(raw)) {
    return {
      call: withArgs("{}"),
      signatureArgs: raw,
      error: "工具参数在输出达到 token 上限时被截断，未执行。请减少单次写入的数据量，分多次调用。",
    };
  }
  const parsed = parseToolArgumentsJson(raw);
  if (!parsed.ok) {
    return { call: withArgs("{}"), signatureArgs: raw, error: `${parsed.error}。本次未执行，请重新发出完整、合法的 JSON 参数。` };
  }
  let value: unknown = parsed.value;
  const definition = TOOL_DEFINITIONS.find((tool) => tool.function.name === call.function.name);
  if (definition) {
    try {
      value = normalizeToolArguments(parsed.value, { ...definition.function.parameters, additionalProperties: false });
    } catch {
      /* 保留原参数，由执行器给出带字段清单的错误 */
    }
  }
  const canonical = JSON.stringify(value);
  return { call: withArgs(canonical), signatureArgs: canonical };
}

function hasImages(messages: ChatMessage[]): boolean {
  return messages.some((m) => Array.isArray(m.content) && m.content.some((p) => p.type === "image_url"));
}

function toolSignature(name: string, args: string): string {
  const canonical = (value: any): any => Array.isArray(value) ? value.map(canonical)
    : value && typeof value === "object" ? Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical(value[key])])) : value;
  try { return `${name}:${JSON.stringify(canonical(JSON.parse(args)))}`; }
  catch { return `${name}:${args.trim()}`; }
}

function formatSteering(notes: string[]): string {
  const body = notes.length === 1 ? notes[0] : notes.map((note, i) => `${i + 1}. ${note}`).join("\n");
  return `[任务进行中用户补充] ${notes.length > 1 ? "\n" : ""}${body}`;
}

function userTextOf(m: ChatMessage): string {
  if (typeof m.content === "string") return m.content;
  if (!Array.isArray(m.content)) return "";
  return m.content
    .filter((p): p is Extract<ContentPart, { type: "text" }> => p.type === "text")
    .map((p) => p.text)
    .join("\n");
}

function safeParse(text: string): unknown {
  try {
    return JSON.parse(text || "{}");
  } catch {
    return {};
  }
}

function isJsonObject(text: string): boolean {
  try {
    const v = JSON.parse(text || "{}");
    return !!v && typeof v === "object" && !Array.isArray(v);
  } catch {
    return false;
  }
}
