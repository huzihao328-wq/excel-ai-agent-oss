/*
 * 批量任务：对多个附加的外部文件逐个执行同一条指令。
 * 每个文件用全新的对话历史跑一次 Agent，结果写入当前工作簿
 * （导入/产出的 Sheet 会带上文件标识，避免互相覆盖）。
 */
import { runAgent } from "./agent";
import { getExternalFiles } from "./tools/externalFiles";
import { beginCheckpointSession } from "./tools/checkpoint";
import { ToolExecContext } from "./tools/excelTools";
import { AppSettings, ChatMessage } from "./types";

export type BatchFileStatus = "pending" | "running" | "done" | "incomplete" | "failed" | "aborted";

export interface BatchToolTrace {
  label: string;
  argsSummary: string;
  status: "running" | "ok" | "error";
}

export interface BatchFileState {
  fileName: string;
  status: BatchFileStatus;
  /** 完成时 assistant 最终回复的摘要 */
  summary?: string;
  error?: string;
  toolTrace: BatchToolTrace[];
}

export interface BatchRunResult {
  states: BatchFileState[];
  aborted: boolean;
}

function describeExternalFile(fileName: string): string {
  const f = getExternalFiles().find((x) => x.name === fileName);
  if (!f) return "（未找到该文件）";
  return f.sheets
    .map((s) => {
      const headers = (s.rows[0] ?? []).filter((h) => String(h).trim()).slice(0, 15);
      return `「${s.name}」${s.rows.length} 行 × ${s.rows.reduce((max, row) => Math.max(max, row.length), 0)} 列，表头: ${headers.join("、") || "(无)"}`;
    })
    .join("；");
}

/**
 * 顺序对每个文件执行指令。onState 用于 UI 逐文件刷新状态。
 * execCtx 复用主界面的确认/选择机制（确认卡会出现在主聊天流中）。
 */
export async function runBatch(
  fileNames: string[],
  instruction: string,
  settings: AppSettings,
  execCtx: ToolExecContext,
  onState: (state: BatchFileState) => void,
  signal: AbortSignal
): Promise<BatchRunResult> {
  // 批量任务指令明确、无需选方向；复盘学习由主对话承担，避免每个文件都反思一次
  const batchSettings: AppSettings = { ...settings, askDirection: false, autoLearnSkills: false };
  if (settings.autoCheckpoint !== false) {
    beginCheckpointSession(`批量: ${instruction}`.replace(/\s+/g, " ").trim().slice(0, 40));
  }

  const states: BatchFileState[] = fileNames.map((fileName) => ({
    fileName,
    status: "pending",
    toolTrace: [],
  }));
  const emit = (state: BatchFileState) => onState({ ...state, toolTrace: state.toolTrace.map((step) => ({ ...step })) });
  let aborted = false;

  for (const state of states) {
    if (signal.aborted) {
      state.status = "aborted";
      emit(state);
      continue;
    }
    state.status = "running";
    emit(state);

    if (!getExternalFiles().some((file) => file.name === state.fileName)) {
      state.status = "failed";
      state.error = "附加文件已移除，请重新附加后执行。";
      emit(state);
      continue;
    }

    const base = state.fileName.replace(/\.[^.]+$/, "");
    const userText = [
      `[批量任务] 本次只处理附加文件「${getExternalFiles().find((x) => x.name === state.fileName)?.caption ?? state.fileName}」（若上下文列出了其他附加文件，它们与本条无关，不要读取）。`,
      `该文件的结构：${describeExternalFile(state.fileName)}`,
      "",
      `[用户指令]`,
      instruction.trim(),
      "",
      "要求：",
      `- 处理结果写入当前工作簿：导入或产出的 Sheet 名以「${base}」开头（需要时加业务后缀，如「${base}-清洗」），避免与其他文件的结果重名。`,
      "- 只针对目标文件的数据执行指令，不要改动当前工作簿里其他无关 Sheet。",
      "- 完成后用两三句话汇报做了什么、结果在哪个 Sheet。",
    ].join("\n");

    const history: ChatMessage[] = [];
    let lastReply = "";
    try {
      const traceById = new Map<string, BatchToolTrace>();
      const outcome = await runAgent(userText, history, batchSettings, execCtx, {
        onAssistantDelta: () => undefined,
        onAssistantDone: (full, meta) => {
          // 工具前的过程叙述不算汇报，只记最终结论
          if (!meta?.interim && full.trim()) lastReply = full;
        },
        onToolStart: (callId, name, label, argsSummary) => {
          if (name === "ask_user_choice" || name === "suggest_followups") return;
          const trace: BatchToolTrace = { label, argsSummary, status: "running" };
          traceById.set(callId, trace);
          state.toolTrace.push(trace);
          emit(state);
        },
        onToolEnd: (callId, result) => {
          const trace = traceById.get(callId);
          if (trace) trace.status = result.ok ? "ok" : "error";
          emit(state);
        },
      }, signal, undefined, { skipCheckpointSession: true });
      if (signal.aborted) {
        state.status = "aborted";
        state.summary = lastReply ? lastReply.slice(0, 200) : "已中止";
      } else if (outcome.hitRoundLimit || outcome.incomplete || state.toolTrace.some((step) => step.status === "error")) {
        state.status = "incomplete";
        state.summary = lastReply.replace(/\s+/g, " ").slice(0, 300);
        state.error = outcome.hitRoundLimit ? "已达到执行轮数上限，任务尚未完成。" : outcome.incomplete ? "模型回复未完成，请检查后继续。" : "执行中有工具失败，请检查结果后重试未完成步骤。";
      } else {
        state.status = "done";
        state.summary = lastReply ? lastReply.replace(/\s+/g, " ").slice(0, 300) : "完成（无文字汇报）";
      }
    } catch (e: any) {
      if (signal.aborted) {
        state.status = "aborted";
        state.summary = "已中止";
      } else {
        state.status = "failed";
        state.error = e?.message || String(e);
      }
    }
    for (const step of state.toolTrace) if (step.status === "running") step.status = "error";
    emit(state);
  }

  if (signal.aborted) aborted = true;
  return { states, aborted };
}
