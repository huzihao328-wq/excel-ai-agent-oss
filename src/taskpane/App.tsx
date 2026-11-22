/* global Office */
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { heuristicFollowups } from "../core/followups";
import { runAgent } from "../core/agent";
import { getCachedModels, listModels, probeContextWindowInfo } from "../core/llm";
import { mergeModelLists } from "../core/modelList";
import { formatContextWindow, getCachedContextWindow, getCachedContextInfo, ContextWindowInfo, describeContextWindow } from "../core/contextWindow";
import { Artifact, getArtifacts, onArtifactsChange } from "../core/artifacts";
import {
  ConversationMeta,
  deleteConversation,
  getCurrentConversationId,
  listConversations,
  loadConversation,
  saveConversation,
  sanitizeItems,
  sanitizeMessages,
  setCurrentConversationId,
  titleFromText,
} from "../core/conversations";
import { isLoopbackEndpoint, loadSettings, saveSettings } from "../core/settings";
import { explainRunError } from "../core/errorExplain";
import { onMemoryChange } from "../core/memory";
import { onSkillsChange } from "../core/skills";
import { buildSuggestedPrompts, SuggestedPrompt } from "../core/recommendations";
import {
  applyUndoTo,
  clearUndoStack,
  getQuickContext,
  getSelectionProfile,
  selectRangeInWorkbook,
  getUndoStack,
  onUndoStackChange,
  SelectionProfile,
  ToolExecContext,
} from "../core/tools/excelTools";
import {
  discardCheckpoint,
  getCheckpoint,
  hydrateCheckpoint,
  onCheckpointChange,
  restoreCheckpoint,
  TaskCheckpoint,
} from "../core/tools/checkpoint";
import {
  AttachedImage,
  MAX_IMAGES_PER_MESSAGE,
  collectClipboardImages,
  isImageFile,
  loadAttachedImage,
} from "../core/images";
import {
  ExternalFile,
  ExternalSelection,
  downloadAttachedFile,
  getExternalFiles,
  getExternalSelection,
  hydrateExternalFiles,
  loadExternalFile,
  onExternalFilesChange,
  onExternalSelectionChange,
  removeExternalFile,
  setExternalSelection,
  toolImportExternalSheet,
} from "../core/tools/externalFiles";
import { toolExportFile } from "../core/tools/exportWorkbook";
import FilePreview from "./components/FilePreview";
import { AppSettings, AuditIssue, ChatMessage, ChoiceOption, ChoiceResolution, ConfirmAnswer, FollowupAction, UiItem, UndoEntry } from "../core/types";
import {
  AssistantBubble,
  AuditCard,
  ChoiceCard,
  ChoiceDecision,
  ConfirmCard,
  ConfirmDecision,
  ErrorBubble,
  FollowupBar,
  NoticeBubble,
  ProcessGroup,
  ThinkingIndicator,
  ToolCard,
  UserBubble,
} from "./components/ChatItems";
import {
  IconArrowDown,
  IconBookmark,
  IconChart,
  IconChats,
  IconChevronDown,
  IconClip,
  IconClose,
  IconDots,
  IconDownload,
  IconFile,
  IconGear,
  IconHistory,
  IconLayers,
  IconPivot,
  IconPlus,
  IconSelection,
  IconSend,
  IconSheet,
  IconSparkle,
  IconStop,
  IconTemplate,
  IconUpload,
  IconZap,
} from "./components/Icons";
import TaskLauncher, { ContextAction } from "./components/TaskLauncher";
import { WorkbookLinks, WorkbookLinksContext } from "./components/Markdown";
import RollbackBar from "./components/RollbackBar";
import WelcomePanel from "./components/WelcomePanel";
import { useAutoSizeTextarea } from "./hooks/useAutoSizeTextarea";
import { MOTION, springScroll } from "./motion";
import Collapse from "./components/Collapse";
import MotionList from "./components/MotionList";
import BatchPanel from "./components/BatchPanel";
import HistoryPanel from "./components/HistoryPanel";
import MemoryPanel from "./components/MemoryPanel";
import Presence from "./components/Presence";
import PromptTemplatesPanel from "./components/PromptTemplatesPanel";
import SkillsPanel from "./components/SkillsPanel";
import SettingsPanel from "./components/SettingsPanel";
import TimelinePanel from "./components/TimelinePanel";

let uidCounter = 0;
const uid = () => `ui_${Date.now()}_${uidCounter++}`;

const PROMPT_HISTORY_KEY = "excel-ai-agent-prompt-history";
const COMPOSER_DRAFT_KEY = "excel-ai-agent-composer-draft";

/** 输入框里没发出去的草稿：关掉任务窗格或插件重新加载后还在 */
function loadComposerDraft(): string {
  try {
    return localStorage.getItem(COMPOSER_DRAFT_KEY) || "";
  } catch {
    return "";
  }
}

function loadPromptHistory(): string[] {
  try {
    const raw = JSON.parse(localStorage.getItem(PROMPT_HISTORY_KEY) || "[]");
    return Array.isArray(raw) ? raw.filter((x) => typeof x === "string" && x.trim()).slice(-50) : [];
  } catch {
    return [];
  }
}

function conversationTitle(items: UiItem[]): string {
  const firstUser = items.find((it) => it.kind === "user");
  if (!firstUser || firstUser.kind !== "user") return "新对话";
  if (firstUser.text.trim()) return titleFromText(firstUser.text);
  if (firstUser.images?.length) return firstUser.images.length === 1 ? "图片" : `图片（${firstUser.images.length}）`;
  return "新对话";
}

/** 按 Excel 的写法引用表名：含空格、符号或以数字开头时加单引号 */
function sheetRefName(name: string): string {
  return /^[A-Za-z_一-鿿][\w一-鿿]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function hasPendingRunInput(items: UiItem[]): boolean {
  return items.some(
    (it) =>
      (it.kind === "confirm" && !it.resolved) || (it.kind === "choice" && !it.resolved)
  );
}

/** 会话里是否真的有用户发过的消息（只有系统提示的会话不值得存进历史） */
function hasUserMessage(items: UiItem[]): boolean {
  return items.some((it) => it.kind === "user");
}

type RunProgressInfo = {
  round: number;
  maxRounds: number;
  phase: "model" | "tools" | "retry";
  retryAttempt?: number;
  retryMax?: number;
};

type RunSession = {
  convId: string;
  history: ChatMessage[];
  items: UiItem[];
  createdAt: number;
  progress: RunProgressInfo | null;
  toolLabel: string | null;
};

function pushPromptHistory(text: string): void {
  const t = text.trim();
  if (!t) return;
  const next = [...loadPromptHistory().filter((x) => x !== t), t].slice(-50);
  try {
    localStorage.setItem(PROMPT_HISTORY_KEY, JSON.stringify(next));
  } catch {
    /* ignore quota */
  }
}

const SUGGESTION_TAGS: Record<SuggestedPrompt["source"], string | undefined> = {
  memory: "记忆",
  skill: "技能",
  both: "记忆·技能",
  selection: undefined,
};

/** 一次任务的起点：最后一条真正发起任务的用户消息（任务中插入的补充说明不算） */
function lastTaskStart(items: UiItem[]): number {
  for (let i = items.length - 1; i >= 0; i--) {
    const it = items[i];
    if (it.kind === "user" && !it.steer) return i;
  }
  return -1;
}

/** 属于执行过程、结束后可以折叠起来的条目 */
function isProcessItem(it: UiItem): boolean {
  switch (it.kind) {
    case "tool":
      return it.status !== "running";
    case "assistant":
      return !!it.interim && !it.streaming;
    case "confirm":
      // 已经答复过的确认也是过程的一部分；还在等答复的留在外面
      return !!it.resolved;
    case "notice":
      return true;
    default:
      return false;
  }
}

type ChatNode = { kind: "item"; item: UiItem } | { kind: "group"; items: UiItem[] };

/**
 * 把已结束任务里连续的执行过程合成一组。正在查看的运行中任务不折叠，
 * 步骤太少的也不折叠（折起来反而多点一下）。
 */
function groupChatItems(items: UiItem[], liveFrom: number): ChatNode[] {
  const nodes: ChatNode[] = [];
  let segment: UiItem[] = [];
  const flush = () => {
    const tools = segment.filter((it) => it.kind === "tool").length;
    if (tools >= 2 && segment.length >= 3) nodes.push({ kind: "group", items: segment });
    else segment.forEach((item) => nodes.push({ kind: "item", item }));
    segment = [];
  };
  items.forEach((item, index) => {
    if (index < liveFrom && isProcessItem(item)) {
      segment.push(item);
      return;
    }
    flush();
    nodes.push({ kind: "item", item });
  });
  flush();
  return nodes;
}



export default function App() {
  const [settings, setSettings] = useState<AppSettings>(loadSettings);
  const [showSettings, setShowSettings] = useState(false);
  const [showModelMenu, setShowModelMenu] = useState(false);
  const [availableModels, setAvailableModels] = useState<string[]>([]);
  const [modelFilter, setModelFilter] = useState("");
  const [modelMenuStatus, setModelMenuStatus] = useState<"idle" | "loading" | "error">("idle");
  const [modelMenuError, setModelMenuError] = useState("");
  const [items, setItems] = useState<UiItem[]>([]);
  const [running, setRunning] = useState(false);
  const [input, setInput] = useState(loadComposerDraft);
  const [undoEntries, setUndoEntries] = useState<UndoEntry[]>([]);
  const [undoBusy, setUndoBusy] = useState(false);
  const [workbookBusy, setWorkbookBusy] = useState(false);
  const [storageWarning, setStorageWarning] = useState(false);
  const [showJumpLatest, setShowJumpLatest] = useState(false);
  const [checkpoint, setCheckpoint] = useState<TaskCheckpoint | null>(getCheckpoint);
  const [cpBusy, setCpBusy] = useState(false);
  const [extFiles, setExtFiles] = useState<ExternalFile[]>([]);
  const [extSelection, setExtSelection] = useState<ExternalSelection | null>(null);
  const [previewFile, setPreviewFile] = useState<string | null>(null);
  const [restoredFiles, setRestoredFiles] = useState(0);
  const [dragging, setDragging] = useState(false);
  const [excelCtx, setExcelCtx] = useState<{
    activeSheet: string;
    selection: string;
    hasFormula: boolean;
    hasError: boolean;
  }>({ activeSheet: "", selection: "", hasFormula: false, hasError: false });
  /** 工作簿里的 Sheet 名，回复中的地址据此变成定位链接；名字没变时保持同一个数组，避免所有回复重新渲染 */
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [checkpointAction, setCheckpointAction] = useState<"restore" | "discard" | null>(null);
  const [showMenu, setShowMenu] = useState(false);
  const [showMemory, setShowMemory] = useState(false);
  const [showSkills, setShowSkills] = useState(false);
  const [showTemplates, setShowTemplates] = useState(false);
  const [showTimeline, setShowTimeline] = useState(false);
  const [showBatch, setShowBatch] = useState(false);
  const [batchRunning, setBatchRunning] = useState(false);
  const [artifacts, setArtifacts] = useState<Artifact[]>(getArtifacts);
  const [conversations, setConversations] = useState<ConversationMeta[]>(listConversations);
  const [currentConvId, setCurrentConvId] = useState<string | null>(getCurrentConversationId);
  const [pendingImages, setPendingImages] = useState<AttachedImage[]>([]);
  const [selProfile, setSelProfile] = useState<SelectionProfile | null>(null);
  const [memoryTick, setMemoryTick] = useState(0);
  const [queued, setQueued] = useState<{ convId: string; text: string; images: AttachedImage[]; paused?: boolean } | null>(null);
  const [runProgress, setRunProgress] = useState<RunProgressInfo | null>(null);
  const [currentToolLabel, setCurrentToolLabel] = useState<string | null>(null);
  const [runningConvId, setRunningConvId] = useState<string | null>(null);
  const [runNeedsInput, setRunNeedsInput] = useState(false);
  const [composerFlash, setComposerFlash] = useState<string | null>(null);

  const historyRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const currentAssistantIdRef = useRef<string | null>(null);
  const confirmResolversRef = useRef<Map<string, { resolve: (v: ConfirmAnswer) => void; tool?: string }>>(new Map());
  /** 任务执行中插入、还没送到模型的补充说明 */
  const steerRef = useRef<{ id: string; text: string }[]>([]);
  /** 本次任务里用户已勾选「同类操作不再询问」的工具 */
  const autoApproveRef = useRef<Set<string>>(new Set());
  const flashTimerRef = useRef<number | null>(null);
  const choiceResolversRef = useRef<Map<string, (v: ChoiceResolution) => void>>(new Map());
  const listRef = useRef<HTMLDivElement>(null);
  const stickToBottomRef = useRef(true);
  const programmaticScrollRef = useRef(false);
  const stopJumpScrollRef = useRef<(() => void) | null>(null);
  const lastScrollTopRef = useRef(0);
  /** 在此时刻之前发生的上移算用户滚动（滚轮、触摸、翻页键之后的一小段时间） */
  const userScrollUntilRef = useRef(0);
  /** 正按住列表本身（滚动条）拖动 */
  const scrollPointerDownRef = useRef(false);
  const scrollRafRef = useRef<number | null>(null);
  const streamFlushRef = useRef<number | null>(null);
  const streamTextRef = useRef("");
  const menuRef = useRef<HTMLDivElement>(null);
  const modelMenuRef = useRef<HTMLDivElement>(null);
  const undoBusyRef = useRef(false);
  const workbookBusyRef = useRef(false);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  useAutoSizeTextarea(inputRef, input);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragDepthRef = useRef(0);
  const currentConvIdRef = useRef<string | null>(currentConvId);
  const itemsRef = useRef<UiItem[]>([]);
  const pendingImagesRef = useRef<AttachedImage[]>([]);
  const createdAtRef = useRef<number>(Date.now());
  const persistTimerRef = useRef<number | null>(null);
  const runPersistTimerRef = useRef<number | null>(null);
  const recTimerRef = useRef<number | null>(null);
  const contextRequestRef = useRef(0);
  const modelsReqRef = useRef(0);
  const toolsUsedRef = useRef<string[]>([]);
  const pendingFollowupsRef = useRef<FollowupAction[] | null>(null);
  const runningRef = useRef(false);
  const batchRunningRef = useRef(false);
  const runDoneRef = useRef(Promise.resolve());
  const batchDoneRef = useRef(Promise.resolve());
  const batchSettleRef = useRef<(() => void) | null>(null);
  const batchAbortRef = useRef<(() => void) | null>(null);
  const batchOpenedRef = useRef(false);
  if (showBatch) batchOpenedRef.current = true;
  const switchingRef = useRef(false);
  const runSessionRef = useRef<RunSession | null>(null);
  const promptHistIdxRef = useRef(-1);
  const promptDraftRef = useRef("");
  /** 插入引用后要放光标的位置，等输入框渲染出新内容后生效 */
  const pendingCaretRef = useRef<number | null>(null);
  const wasThinkingRef = useRef(false);
  const thinkingSinceRef = useRef(0);
  runningRef.current = running;
  batchRunningRef.current = batchRunning;

  const withWorkbookAction = useCallback(async <T,>(action: () => Promise<T>): Promise<T> => {
    if (runningRef.current || batchRunningRef.current || workbookBusyRef.current) {
      throw new Error("工作簿正在处理其他操作，请等待完成后重试。");
    }
    workbookBusyRef.current = true;
    setWorkbookBusy(true);
    try {
      return await action();
    } finally {
      workbookBusyRef.current = false;
      setWorkbookBusy(false);
    }
  }, []);

  /**
   * 本地提示（附件、撤销、导出、模型切换、排队提示等）追加到用户正在看的会话。
   * 与后台运行中的会话无关，所以不能像 addRunItem 那样一律塞进运行会话——否则用户在别的会话里
   * 附加文件、切换模型时什么反馈都看不到，而那条提示却出现在另一个会话里。
   */
  const addItem = useCallback((item: UiItem) => {
    const run = runSessionRef.current;
    if (run && currentConvIdRef.current === run.convId) {
      run.items = [...run.items, item];
      setItems(run.items.slice());
      return;
    }
    setItems((prev) => [...prev, item]);
  }, []);
  const updateItem = useCallback((id: string, patch: Partial<UiItem>) => {
    setItems((prev) => prev.map((it) => (it.id === id ? ({ ...it, ...patch } as UiItem) : it)));
  }, []);

  const persistSession = useCallback((session: Pick<RunSession, "convId" | "items" | "history" | "createdAt">) => {
    // 只有本地提示（附加了文件、切了模型…）还没发过消息的会话不落盘，否则历史里会堆满空的「新对话」
    if (!hasUserMessage(session.items)) return;
    const saved = saveConversation({
      id: session.convId,
      title: conversationTitle(session.items),
      createdAt: session.createdAt,
      updatedAt: Date.now(),
      items: session.items,
      messages: [...session.history],
    });
    setStorageWarning(!saved);
    setConversations(listConversations());
  }, []);

  const persistRunSession = useCallback(
    (immediate = false) => {
      const run = runSessionRef.current;
      if (!run) return;
      const save = () => persistSession(run);
      if (immediate) {
        if (runPersistTimerRef.current != null) {
          window.clearTimeout(runPersistTimerRef.current);
          runPersistTimerRef.current = null;
        }
        save();
        return;
      }
      if (runPersistTimerRef.current != null) return;
      runPersistTimerRef.current = window.setTimeout(() => {
        runPersistTimerRef.current = null;
        save();
      }, 800);
    },
    [persistSession]
  );

  const addRunItem = useCallback(
    (item: UiItem) => {
      const run = runSessionRef.current;
      if (!run) {
        addItem(item);
        return;
      }
      run.items = [...run.items, item];
      if (currentConvIdRef.current === run.convId) setItems(run.items.slice());
      else persistRunSession();
      if ((item.kind === "confirm" && !item.resolved) || (item.kind === "choice" && !item.resolved)) {
        setRunNeedsInput(true);
      }
    },
    [addItem, persistRunSession]
  );

  const updateRunItem = useCallback(
    (id: string, patch: Partial<UiItem>) => {
      const run = runSessionRef.current;
      if (!run) {
        updateItem(id, patch);
        return;
      }
      run.items = run.items.map((it) => (it.id === id ? ({ ...it, ...patch } as UiItem) : it));
      if (currentConvIdRef.current === run.convId) setItems(run.items.slice());
      else persistRunSession();
    },
    [persistRunSession, updateItem]
  );

  const removeRunItem = useCallback(
    (id: string) => {
      const run = runSessionRef.current;
      if (!run) {
        setItems((prev) => prev.filter((it) => it.id !== id));
        return;
      }
      run.items = run.items.filter((it) => it.id !== id);
      if (currentConvIdRef.current === run.convId) setItems(run.items.slice());
      else persistRunSession();
    },
    [persistRunSession]
  );

  const setCurrentId = useCallback((id: string | null) => {
    currentConvIdRef.current = id;
    setCurrentConvId(id);
    if (!setCurrentConversationId(id)) setStorageWarning(true);
  }, []);

  const persistCurrent = useCallback(() => {
    const run = runSessionRef.current;
    if (run && currentConvIdRef.current === run.convId) {
      persistSession(run);
      return;
    }
    const itemsNow = itemsRef.current;
    if (!hasUserMessage(itemsNow)) return;
    let id = currentConvIdRef.current;
    if (!id) {
      id = uid();
      currentConvIdRef.current = id;
      setCurrentConvId(id);
      setCurrentConversationId(id);
    }
    persistSession({
      convId: id,
      items: itemsNow,
      history: historyRef.current,
      createdAt: createdAtRef.current,
    });
  }, [persistSession]);

  const flushPersist = useCallback(() => {
    if (persistTimerRef.current != null) {
      window.clearTimeout(persistTimerRef.current);
      persistTimerRef.current = null;
    }
    persistCurrent();
  }, [persistCurrent]);

  itemsRef.current = items;
  pendingImagesRef.current = pendingImages;

  // 恢复上次对话
  useEffect(() => {
    const id = getCurrentConversationId();
    if (!id) return;
    const data = loadConversation(id);
    if (!data) {
      currentConvIdRef.current = null;
      setCurrentConvId(null);
      setCurrentConversationId(null);
      return;
    }
    currentConvIdRef.current = data.id;
    createdAtRef.current = data.createdAt;
    historyRef.current = data.messages;
    setCurrentConvId(data.id);
    setItems(sanitizeItems(data.items));
  }, []);

  // 自动滚动到底部。用户滚轮往上翻时立刻停住，避免出字过程把滚动抢走。
  const setStuckToBottom = useCallback((stuck: boolean) => {
    if (stickToBottomRef.current === stuck) return;
    stickToBottomRef.current = stuck;
    setShowJumpLatest(!stuck);
  }, []);

  const scrollChatToBottom = useCallback(() => {
    const el = listRef.current;
    // 回到底部的弹簧滚动会自己追踪不断变化的底部，这里不再抢写 scrollTop。
    if (!el || !stickToBottomRef.current || itemsRef.current.length === 0 || stopJumpScrollRef.current) return;
    programmaticScrollRef.current = true;
    el.scrollTop = el.scrollHeight;
    lastScrollTopRef.current = el.scrollTop;
    requestAnimationFrame(() => {
      programmaticScrollRef.current = false;
    });
  }, []);

  useEffect(() => {
    if (items.length === 0) {
      stopJumpScrollRef.current?.();
      if (listRef.current) listRef.current.scrollTop = 0;
      lastScrollTopRef.current = 0;
      setStuckToBottom(true);
      return;
    }
    if (!stickToBottomRef.current) return;
    if (scrollRafRef.current != null) return;
    scrollRafRef.current = requestAnimationFrame(() => {
      scrollRafRef.current = null;
      scrollChatToBottom();
    });
  }, [items, scrollChatToBottom, setStuckToBottom]);

  // 输入区/附件平滑变高时保持底部锚定；用户翻阅历史时不改变其滚动位置。
  useEffect(() => {
    const element = listRef.current;
    if (!element) return;
    const observer = new ResizeObserver(() => {
      if (!stickToBottomRef.current || scrollRafRef.current != null) return;
      scrollRafRef.current = requestAnimationFrame(() => { scrollRafRef.current = null; scrollChatToBottom(); });
    });
    observer.observe(element);
    return () => { observer.disconnect(); if (scrollRafRef.current != null) cancelAnimationFrame(scrollRafRef.current); scrollRafRef.current = null; };
  }, [scrollChatToBottom]);

  useEffect(() => {
    const el = listRef.current;
    if (!el) return;
    const intend = (ms: number) => {
      userScrollUntilRef.current = Math.max(userScrollUntilRef.current, Date.now() + ms);
    };
    const onWheel = (e: WheelEvent) => {
      intend(600);
      if (itemsRef.current.length > 0 && e.deltaY < 0) setStuckToBottom(false);
    };
    // 只有按在列表本身上（滚动条）才算拖动滚动；点卡片里的按钮不算
    const onPointerDown = (e: PointerEvent) => {
      if (e.target === el) scrollPointerDownRef.current = true;
    };
    const onPointerUp = () => {
      if (!scrollPointerDownRef.current) return;
      scrollPointerDownRef.current = false;
      intend(300);
    };
    // 触摸滑动松手后还有惯性滚动
    const onTouch = () => intend(1500);
    const onKeyDown = (e: KeyboardEvent) => {
      // Tab 聚焦到视野外的按钮时浏览器会滚过去，这也是用户意图
      if (["PageUp", "PageDown", "ArrowUp", "ArrowDown", "Home", "End", " ", "Tab"].includes(e.key)) intend(800);
    };
    el.addEventListener("wheel", onWheel, { passive: true });
    el.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
    el.addEventListener("touchmove", onTouch, { passive: true });
    el.addEventListener("touchend", onTouch, { passive: true });
    el.addEventListener("keydown", onKeyDown);
    return () => {
      el.removeEventListener("wheel", onWheel);
      el.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("pointerup", onPointerUp);
      window.removeEventListener("pointercancel", onPointerUp);
      el.removeEventListener("touchmove", onTouch);
      el.removeEventListener("touchend", onTouch);
      el.removeEventListener("keydown", onKeyDown);
    };
  }, [setStuckToBottom]);

  const onListScroll = useCallback(() => {
    const el = listRef.current;
    if (!el || itemsRef.current.length === 0) return;
    const previousTop = lastScrollTopRef.current;
    const currentTop = el.scrollTop;
    lastScrollTopRef.current = currentTop;
    if (programmaticScrollRef.current) return;

    // 向上滚（包括拖动滚动条）立即锁住当前位置。不能用“离底部 32px”判断，
    // 否则一次小幅上滚会立刻被下一次流式渲染重新吸回底部。
    // 例外：输入区收回后列表变高，浏览器把 scrollTop 夹回最大值，也表现为“变小”，但仍停在底部，不是用户上滚。
    const distanceToBottom = el.scrollHeight - currentTop - el.clientHeight;
    if (currentTop < previousTop - 1 && distanceToBottom > 1) {
      // 只有紧跟在用户操作（滚轮、拖滚动条、触摸、翻页键）之后的上移才算用户上翻。
      // 其余上移是布局位移：输入框长高、点过的按钮被卡片内容替换时 WebView 都会挪动滚动位置，
      // 当成上翻会让对话停止跟随、弹出「回到最新」。这种情况继续贴底。
      const userDriven = scrollPointerDownRef.current || Date.now() < userScrollUntilRef.current;
      if (!userDriven && stickToBottomRef.current) {
        scrollChatToBottom();
        return;
      }
      setStuckToBottom(false);
      return;
    }

    // 锁住后只有用户确实滚到最底部才恢复自动跟随。
    if (distanceToBottom <= 1 && currentTop > previousTop) setStuckToBottom(true);
  }, [scrollChatToBottom, setStuckToBottom]);

  const jumpToLatest = useCallback(() => {
    stickToBottomRef.current = true;
    setShowJumpLatest(false);
    const el = listRef.current;
    if (!el) return;
    stopJumpScrollRef.current?.();
    programmaticScrollRef.current = true;
    let finished = false;
    const stop = springScroll(el, () => el.scrollHeight - el.clientHeight, () => {
      finished = true;
      stopJumpScrollRef.current = null;
      lastScrollTopRef.current = el.scrollTop;
      requestAnimationFrame(() => { programmaticScrollRef.current = false; });
    });
    if (!finished) stopJumpScrollRef.current = stop;
  }, []);
  useEffect(() => () => stopJumpScrollRef.current?.(), []);

  // 对话内容变化后自动落盘
  useEffect(() => {
    if (items.length === 0) return;
    if (persistTimerRef.current != null) window.clearTimeout(persistTimerRef.current);
    persistTimerRef.current = window.setTimeout(() => {
      persistTimerRef.current = null;
      persistCurrent();
    }, 400);
    return () => {
      if (persistTimerRef.current != null) {
        window.clearTimeout(persistTimerRef.current);
        persistTimerRef.current = null;
      }
    };
  }, [items, persistCurrent]);

  // 关闭任务窗格时再刷一次，避免丢最后几条
  useEffect(() => {
    const onVisibility = () => {
      if (document.visibilityState === "hidden") flushPersist();
    };
    window.addEventListener("pagehide", flushPersist);
    document.addEventListener("visibilitychange", onVisibility);
    return () => {
      window.removeEventListener("pagehide", flushPersist);
      document.removeEventListener("visibilitychange", onVisibility);
    };
  }, [flushPersist]);

  // 插入引用后，等新内容渲染进输入框再把光标放到插入点后面（直接设会被随后的赋值挪到末尾）
  useLayoutEffect(() => {
    const caret = pendingCaretRef.current;
    const field = inputRef.current;
    if (caret == null || !field) return;
    pendingCaretRef.current = null;
    field.focus();
    field.setSelectionRange(caret, caret);
  }, [input]);

  useEffect(() => {
    try {
      if (input) localStorage.setItem(COMPOSER_DRAFT_KEY, input);
      else localStorage.removeItem(COMPOSER_DRAFT_KEY);
    } catch {
      /* 存储不可用时只是不保留草稿 */
    }
  }, [input]);

  /** 输入区底部的一句临时提示：回车没发出去时告诉用户原因，而不是静默吞掉 */
  const flashComposer = useCallback((text: string) => {
    if (flashTimerRef.current != null) window.clearTimeout(flashTimerRef.current);
    setComposerFlash(text);
    flashTimerRef.current = window.setTimeout(() => {
      flashTimerRef.current = null;
      setComposerFlash(null);
    }, 4000);
  }, []);

  /** 回复里的地址链接：在 Excel 里切到对应 Sheet 并选中区域；选区变化事件会刷新底部的工作表胶囊 */
  const locateInWorkbook = useCallback(
    async (sheet: string, address?: string) => {
      try {
        await selectRangeInWorkbook(sheet, address);
      } catch (e: any) {
        flashComposer(`没能定位到 ${address ? `${sheet}!${address}` : `「${sheet}」`}：${e?.message || String(e)}`);
      }
    },
    [flashComposer]
  );
  const workbookLinks = useMemo<WorkbookLinks>(
    () => ({ sheetNames, locate: (sheet, address) => void locateInWorkbook(sheet, address) }),
    [sheetNames, locateInWorkbook]
  );

  useEffect(() => {
    return () => {
      if (flashTimerRef.current != null) window.clearTimeout(flashTimerRef.current);
      abortRef.current?.abort();
      batchAbortRef.current?.();
      confirmResolversRef.current.forEach(({ resolve }) => resolve({ accepted: false }));
      confirmResolversRef.current.clear();
      choiceResolversRef.current.forEach((resolve) => resolve({ kind: "cancelled" }));
      choiceResolversRef.current.clear();
    };
  }, []);

  // 下拉菜单：点击外部或按 Esc 关闭
  useEffect(() => {
    if (!showMenu) return;
    const onDown = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) setShowMenu(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setShowMenu(false);
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [showMenu]);

  // Esc 关闭最上层的浮层面板（设置 > 文件预览 > 其余面板），一次只关一层；内部已处理 Esc 的组件会阻止冒泡到这里
  const anyOverlayOpen = showSettings || !!previewFile || showBatch || showTemplates || showTimeline || showSkills || showMemory || showHistory;
  useEffect(() => {
    if (!anyOverlayOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || e.isComposing) return;
      if (showSettings) setShowSettings(false);
      else if (previewFile) setPreviewFile(null);
      else if (showBatch) setShowBatch(false);
      else if (showTemplates) setShowTemplates(false);
      else if (showTimeline) setShowTimeline(false);
      else if (showSkills) setShowSkills(false);
      else if (showMemory) setShowMemory(false);
      else if (showHistory) setShowHistory(false);
    };
    document.addEventListener("keydown", onKey);
    return () => document.removeEventListener("keydown", onKey);
  }, [anyOverlayOpen, showSettings, previewFile, showBatch, showTemplates, showTimeline, showSkills, showMemory, showHistory]);

  useEffect(() => {
    if (!showModelMenu) return;
    const onDown = (e: MouseEvent) => {
      if (modelMenuRef.current && !modelMenuRef.current.contains(e.target as Node)) setShowModelMenu(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setShowModelMenu(false);
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [showModelMenu]);

  useEffect(() => {
    modelsReqRef.current++;
    setAvailableModels(mergeModelLists(getCachedModels(settings), [settings.model]));
    setModelFilter("");
    setModelMenuStatus("idle");
    setModelMenuError("");
    setShowModelMenu(false);
  }, [settings.baseUrl, settings.apiKey]);

  const [detectedContext, setDetectedContext] = useState<ContextWindowInfo | null>(null);
  const displayedContextTokens = settings.contextWindowTokens || detectedContext?.tokens || 0;
  const displayedContextHint = settings.contextWindowTokens
    ? `${formatContextWindow(settings.contextWindowTokens)}（手动设置）`
    : detectedContext ? describeContextWindow(detectedContext) : "";
  useEffect(() => {
    setDetectedContext(null);
    const model = settings.model.trim();
    if (!model) return;
    if (!settings.apiKey.trim() && !isLoopbackEndpoint(settings.baseUrl)) return;
    let cancelled = false;
    void (async () => {
      const n = await probeContextWindowInfo(settings).catch(() => null);
      if (cancelled || !n) return;
      // 自动探测只进短期缓存，不固化为手动配置，避免估算长期覆盖新探测。
      setDetectedContext(n);
    })();
    return () => {
      cancelled = true;
    };
  }, [settings.baseUrl, settings.apiKey, settings.model]);

  // 订阅撤销栈变化
  useEffect(() => {
    setUndoEntries([...getUndoStack()]);
    return onUndoStackChange(() => setUndoEntries([...getUndoStack()]));
  }, []);

  // 找回上次未丢弃的任务检查点
  useEffect(() => {
    const off = onCheckpointChange((cp, event) => {
      setCheckpoint(cp);
      setCpBusy(event === "creating");
      if (event === "created" && cp) {
        addItem({
          kind: "notice",
          id: uid(),
          text:
            cp.storage === "file"
              ? `已备份当前工作簿（${cp.sheets.length} 个 Sheet）。可用底部「恢复到任务开始」整段回退；备份存在本机插件存储里，不会写进工作簿。`
              : `已备份当前工作簿（${cp.sheets.length} 个 Sheet）。可用底部「恢复到任务开始」整段回退；当前 Excel 版本只能把备份存为隐藏表，会随文件保存，不需要时请丢弃。`,
        });
      } else if (event === "cleanup-failed") {
        addItem({
          kind: "notice",
          id: uid(),
          text: "当前任务备份已创建，可正常恢复；部分旧备份未清理，可能增加文件大小。",
        });
      } else if (event === "failed") {
        addItem({
          kind: "notice",
          id: uid(),
          text: "任务检查点创建失败，本次仍会继续执行；删行/删表将无法整段回退。",
        });
      }
    });
    void hydrateCheckpoint();
    return off;
  }, [addItem]);

  // 订阅外部文件变化
  useEffect(() => {
    setExtFiles(getExternalFiles());
    return onExternalFilesChange(() => {
      setExtFiles(getExternalFiles());
      setPreviewFile((prev) => (prev && !getExternalFiles().some((f) => f.name === prev) ? null : prev));
    });
  }, []);

  useEffect(() => {
    void hydrateExternalFiles()
      .then((n) => {
        if (n > 0) setRestoredFiles(n);
      })
      .catch(() => undefined);
  }, []);

  // 订阅外部文件选区变化
  useEffect(() => {
    setExtSelection(getExternalSelection());
    return onExternalSelectionChange(() => setExtSelection(getExternalSelection()));
  }, []);

  const handleFiles = useCallback(
    async (fileList: FileList | File[]) => {
      const spreadsheet = /\.(xlsx|xlsm|xls|csv)$/i;
      const files = Array.from(fileList);
      // 图片先按剩余名额筛一遍：超出的只提示一次，也避免在 setState 更新函数里做副作用
      const imageFiles = files.filter(isImageFile);
      const room = Math.max(0, MAX_IMAGES_PER_MESSAGE - pendingImagesRef.current.length);
      if (imageFiles.length > room) {
        addItem({
          kind: "error",
          id: uid(),
          text: `一次最多发送 ${MAX_IMAGES_PER_MESSAGE} 张图片，已忽略 ${imageFiles.length - room} 张。`,
        });
      }
      const accepted = new Set(imageFiles.slice(0, room));
      for (const file of files) {
        if (isImageFile(file)) {
          if (!accepted.has(file)) continue;
          try {
            const img = await loadAttachedImage(file);
            setPendingImages((prev) => [...prev, img]);
          } catch (e: any) {
            addItem({ kind: "error", id: uid(), text: e?.message || String(e) });
          }
          continue;
        }
        if (!spreadsheet.test(file.name)) {
          addItem({
            kind: "error",
            id: uid(),
            text: `不支持的文件类型: ${file.name}（支持 xlsx / xls / xlsm / csv，以及 png / jpg / gif / webp 图片）`,
          });
          continue;
        }
        try {
          const msg = await loadExternalFile(file);
          addItem({ kind: "notice", id: uid(), text: `${msg}。点击文件胶囊可预览并框选部分数据` });
        } catch (e: any) {
          addItem({ kind: "error", id: uid(), text: `附加文件失败: ${e?.message || String(e)}` });
        }
      }
    },
    [addItem]
  );

  // 跟踪活动 Sheet 与选区
  const refreshExcelCtx = useCallback(async () => {
    const request = ++contextRequestRef.current;
    if (recTimerRef.current != null) window.clearTimeout(recTimerRef.current);
    try {
      const ctx = await getQuickContext(false);
      if (request !== contextRequestRef.current) return;
      setExcelCtx({
        activeSheet: ctx.activeSheet,
        selection: ctx.selectionAddress.split("!").pop() || "",
        hasFormula: !!ctx.selectionHasFormula,
        hasError: !!ctx.selectionHasError,
      });
      setSheetNames((prev) => (prev.join("\u0000") === ctx.sheetNames.join("\u0000") ? prev : ctx.sheetNames));
    } catch {
      if (request !== contextRequestRef.current) return;
      setExcelCtx({ activeSheet: "", selection: "", hasFormula: false, hasError: false });
    }
    if (recTimerRef.current != null) window.clearTimeout(recTimerRef.current);
    recTimerRef.current = window.setTimeout(async () => {
      recTimerRef.current = null;
      try {
        const profile = await getSelectionProfile();
        if (request === contextRequestRef.current) setSelProfile(profile);
      } catch {
        if (request === contextRequestRef.current) setSelProfile(null);
      }
    }, 140);
  }, []);

  useEffect(() => {
    refreshExcelCtx();
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refreshExcelCtx);
    } catch {
      /* 非 Excel 环境 */
    }
    return () => {
      contextRequestRef.current++;
      if (recTimerRef.current != null) window.clearTimeout(recTimerRef.current);
      try {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: refreshExcelCtx });
      } catch { /* Office may already be disconnected. */ }
    };
  }, [refreshExcelCtx]);

  useEffect(() => onMemoryChange(() => setMemoryTick((n) => n + 1)), []);
  useEffect(() => onSkillsChange(() => setMemoryTick((n) => n + 1)), []);
  useEffect(() => onArtifactsChange(() => setArtifacts(getArtifacts())), []);

  const suggestedPrompts = useMemo(
    () => buildSuggestedPrompts(selProfile, settings.useMemory !== false),
    [selProfile, settings.useMemory, memoryTick]
  );
  const suggestionsFromContext = !!selProfile?.isRange;

  const resolveConfirm = useCallback(
    (id: string, decision: ConfirmDecision) => {
      const pending = confirmResolversRef.current.get(id);
      const note = decision.accepted ? undefined : decision.note?.trim() || undefined;
      if (pending) {
        if (decision.accepted && decision.allowSimilar && pending.tool) autoApproveRef.current.add(pending.tool);
        pending.resolve({ accepted: decision.accepted, note });
        confirmResolversRef.current.delete(id);
      }
      updateRunItem(id, {
        resolved: decision.accepted ? "accepted" : "rejected",
        ...(note ? { note } : {}),
      } as Partial<UiItem>);
      setRunNeedsInput(false);
    },
    [updateRunItem]
  );

  const resolveChoice = useCallback(
    (id: string, resolution: ChoiceDecision) => {
      const resolver = choiceResolversRef.current.get(id);
      const custom = typeof resolution === "object" && "custom" in resolution ? resolution.custom : null;
      if (resolver) {
        if (resolution === "skipped") resolver({ kind: "skipped" });
        else if (resolution === "cancelled") resolver({ kind: "cancelled" });
        else if (custom != null) resolver({ kind: "custom", text: custom });
        else if ("id" in resolution) resolver({ kind: "selected", id: resolution.id, label: resolution.label });
        choiceResolversRef.current.delete(id);
      }
      updateRunItem(id, {
        resolved: custom != null ? { id: "custom", label: custom, custom: true } : resolution,
      } as Partial<UiItem>);
      setRunNeedsInput(false);
    },
    [updateRunItem]
  );

  /** 主对话与批量任务共用的工具执行上下文（确认/选择卡片出现在主聊天流中） */
  const buildExecCtx = useCallback((): ToolExecContext => {
    const requestConfirm = (title: string, detail: string, tool?: string) =>
      new Promise<ConfirmAnswer>((resolve) => {
        // 本次任务里用户已允许过同类操作：不再打断，但仍留一条可展开的记录
        if (tool && autoApproveRef.current.has(tool)) {
          addRunItem({ kind: "confirm", id: uid(), title, detail, tool, resolved: "accepted", auto: true });
          resolve({ accepted: true });
          return;
        }
        if (batchRunningRef.current) setShowBatch(false);
        const id = uid();
        confirmResolversRef.current.set(id, { resolve, tool });
        addRunItem({ kind: "confirm", id, title, detail, tool });
      });
    return {
      settings,
      confirm: async (title, detail) => (await requestConfirm(title, detail)).accepted,
      confirmDetailed: (title, detail, meta) => requestConfirm(title, detail, meta.tool),
      choose: (title, question, options: ChoiceOption[]) =>
        new Promise<ChoiceResolution>((resolve) => {
          if (batchRunningRef.current) setShowBatch(false);
          const id = uid();
          choiceResolversRef.current.set(id, resolve);
          addRunItem({ kind: "choice", id, title, question, options });
        }),
    };
  }, [settings, addRunItem]);

  const send = useCallback(
    async (text: string, options?: { skipUserBubble?: boolean; images?: AttachedImage[]; resumeTurn?: boolean; preserveComposer?: boolean }) => {
      const trimmed = text.trim();
      const images = options?.images ?? pendingImages;
      if ((!trimmed && images.length === 0) || runningRef.current || batchRunningRef.current || workbookBusyRef.current) return;
      if (images.length > MAX_IMAGES_PER_MESSAGE) {
        addItem({ kind: "error", id: uid(), text: `一次最多发送 ${MAX_IMAGES_PER_MESSAGE} 张图片，请分批发送。` });
        return;
      }
      if (!settings.model.trim()) {
        addItem({ kind: "notice", id: uid(), text: "请先选择或填写模型名。当前输入已保留。" });
        setShowSettings(true);
        return;
      }
      if (!settings.apiKey.trim() && !isLoopbackEndpoint(settings.baseUrl)) {
        addItem({ kind: "notice", id: uid(), text: "请先在设置里填写 API Key（右上角齿轮图标）。" });
        setShowSettings(true);
        return;
      }
      runningRef.current = true;
      setRunning(true);
      if (!options?.preserveComposer) setInput("");
      promptHistIdxRef.current = -1;
      if (trimmed) pushPromptHistory(trimmed);
      if (!options?.preserveComposer) setPendingImages([]);
      toolsUsedRef.current = [];
      pendingFollowupsRef.current = null;
      steerRef.current = [];
      autoApproveRef.current = new Set();
      setRunProgress(null);
      setCurrentToolLabel(null);
      stickToBottomRef.current = true;
      setShowJumpLatest(false);
      if (streamFlushRef.current != null) {
        window.clearTimeout(streamFlushRef.current);
        streamFlushRef.current = null;
      }
      if (!currentConvIdRef.current) {
        const id = uid();
        createdAtRef.current = Date.now();
        setCurrentId(id);
      }
      const convId = currentConvIdRef.current!;
      const seededItems = itemsRef.current.filter((it) => it.kind !== "followups");
      const runHistory = historyRef.current;
      runSessionRef.current = {
        convId,
        history: runHistory,
        items: seededItems,
        createdAt: createdAtRef.current,
        progress: null,
        toolLabel: null,
      };
      setItems(seededItems);
      setRunningConvId(convId);
      setRunNeedsInput(false);
      const displayText = trimmed || (images.length ? "请查看我发送的图片。" : "");
      if (!options?.skipUserBubble) {
        addRunItem({
          kind: "user",
          id: uid(),
          text: displayText,
          images: images.map((img) => ({ id: img.id, name: img.name, dataUrl: img.dataUrl })),
        });
      }
      const abort = new AbortController();
      abortRef.current = abort;
      currentAssistantIdRef.current = null;
      let settleRun!: () => void;
      runDoneRef.current = new Promise<void>((resolve) => {
        settleRun = resolve;
      });

      let hitRoundLimit = false;
      let incomplete = false;
      let runFailed = false;
      try {
        const outcome = await runAgent(
          displayText,
          runHistory,
          settings,
          buildExecCtx(),
          {
            onProgress: (info) => {
              const run = runSessionRef.current;
              if (run) run.progress = info;
              if (!run || currentConvIdRef.current === run.convId) setRunProgress(info);
              if (info.phase === "model") {
                if (run) run.toolLabel = null;
                if (!run || currentConvIdRef.current === run.convId) setCurrentToolLabel(null);
              }
              if (info.phase === "retry") {
                const label = `网络不稳，正在重试${info.retryAttempt && info.retryMax ? `（${info.retryAttempt}/${info.retryMax}）` : ""}`;
                if (run) run.toolLabel = label;
                if (!run || currentConvIdRef.current === run.convId) setCurrentToolLabel(label);
              }
            },
            onAssistantDelta: (full) => {
              streamTextRef.current = full;
              if (!currentAssistantIdRef.current) {
                const id = uid();
                currentAssistantIdRef.current = id;
                addRunItem({ kind: "assistant", id, text: full, streaming: true });
                return;
              }
              if (streamFlushRef.current != null) return;
              streamFlushRef.current = window.setTimeout(() => {
                streamFlushRef.current = null;
                const id = currentAssistantIdRef.current;
                if (id) updateRunItem(id, { text: streamTextRef.current } as Partial<UiItem>);
              }, 120);
            },
            onAssistantDone: (full, meta) => {
              if (streamFlushRef.current != null) {
                window.clearTimeout(streamFlushRef.current);
                streamFlushRef.current = null;
              }
              const interim = !!meta?.interim;
              const streamingId = currentAssistantIdRef.current;
              currentAssistantIdRef.current = null;
              streamTextRef.current = "";
              if (!full.trim()) {
                // 正文被解析成了工具调用（或本来就没说话）：去掉空气泡
                if (streamingId) removeRunItem(streamingId);
                return;
              }
              if (streamingId) {
                updateRunItem(streamingId, { text: full, streaming: false, interim } as Partial<UiItem>);
              } else {
                addRunItem({ kind: "assistant", id: uid(), text: full, interim });
              }
            },
            onNotice: (text) => {
              addRunItem({ kind: "notice", id: uid(), text });
            },
            hasSteering: () => steerRef.current.length > 0,
            takeSteering: () => {
              const notes = steerRef.current;
              steerRef.current = [];
              notes.forEach((note) => updateRunItem(note.id, { pending: false } as Partial<UiItem>));
              return notes.map((note) => note.text);
            },
            onToolStart: (callId, _name, label, argsSummary) => {
              if (_name === "ask_user_choice" || _name === "suggest_followups") return;
              toolsUsedRef.current.push(_name);
              const run = runSessionRef.current;
              if (run) run.toolLabel = label;
              if (!run || currentConvIdRef.current === run.convId) setCurrentToolLabel(label);
              addRunItem({ kind: "tool", id: callId, name: _name, label, argsSummary, status: "running", startedAt: Date.now() });
            },
            onSkillsMatched: (names) => {
              addRunItem({ kind: "notice", id: uid(), text: `调用技能：${names.join("、")}` });
            },
            onLearned: (message) => {
              const item: UiItem = { kind: "notice", id: uid(), text: message };
              const run = runSessionRef.current;
              if (run?.convId === convId) {
                addRunItem(item);
                return;
              }
              if (currentConvIdRef.current === convId) {
                setItems((prev) => [...prev, item]);
                return;
              }
              const rec = loadConversation(convId);
              if (rec) {
                saveConversation({ ...rec, items: [...rec.items, item] });
                setConversations(listConversations());
              }
            },
            onToolEnd: (callId, result) => {
              const run = runSessionRef.current;
              if (run) run.toolLabel = null;
              if (!run || currentConvIdRef.current === run.convId) setCurrentToolLabel(null);
              if (result.followups?.length) pendingFollowupsRef.current = result.followups;
              if (result.audit) {
                addRunItem({
                  kind: "audit",
                  id: uid(),
                  sheet: result.audit.sheet,
                  range: result.audit.range,
                  reportSheet: result.audit.reportSheet,
                  issues: result.audit.issues,
                });
              }
              let detail = "";
              if (result.ok) {
                detail = result.summary || "";
                if (result.data !== undefined) {
                  const dataStr = JSON.stringify(result.data, null, 2);
                  if (dataStr && dataStr !== "{}") detail += (detail ? "\n" : "") + dataStr.slice(0, 2000);
                }
              } else {
                detail = result.error || "未知错误";
              }
              updateRunItem(callId, {
                status: result.ok ? "ok" : "error",
                detail,
                undoable: result.undoable,
              } as Partial<UiItem>);
              refreshExcelCtx();
            },
          },
          abort.signal,
          images,
          options?.resumeTurn ? { skipUserPush: true, skipCheckpointSession: true } : undefined
        );
        hitRoundLimit = !!outcome?.hitRoundLimit;
        incomplete = !!outcome?.incomplete;
      } catch (e: any) {
        runFailed = true;
        if (abort.signal.aborted) {
          addRunItem({ kind: "notice", id: uid(), text: "已停止。" });
        } else {
          // 用户消息保留在历史里：重试时 runAgent 会识别到同一条而不重复推入，
          // 直接发新消息时模型也仍能看到这个没回答上的问题。
          addRunItem({ kind: "error", id: uid(), ...explainRunError(e, { loopback: isLoopbackEndpoint(settings.baseUrl) }), retryable: true });
        }
      } finally {
        if (runFailed || abort.signal.aborted || hitRoundLimit || incomplete) {
          setQueued((pending) => {
            if (pending?.convId !== convId) return pending;
            return { ...pending, paused: true };
          });
        }
        // 无论怎么结束，都不能让历史里留下没有 tool 结果的 tool_calls（下次请求会被接口 400）
        const repaired = sanitizeMessages(runHistory);
        if (repaired.length !== runHistory.length) runHistory.splice(0, runHistory.length, ...repaired);
        autoApproveRef.current = new Set();
        // 停止或出错时还没送到模型的补充说明：从对话里撤下，原文放回输入框，不能悄悄丢掉
        const undelivered = steerRef.current;
        steerRef.current = [];
        if (undelivered.length) {
          undelivered.forEach((note) => removeRunItem(note.id));
          setInput((draft) => [...undelivered.map((note) => note.text), draft].filter((t) => t.trim()).join("\n"));
          addRunItem({ kind: "notice", id: uid(), text: "任务已结束，补充说明没来得及送达，已放回输入框。" });
        }
        confirmResolversRef.current.forEach(({ resolve }) => resolve({ accepted: false }));
        confirmResolversRef.current.clear();
        choiceResolversRef.current.forEach((resolve) => resolve({ kind: "cancelled" }));
        choiceResolversRef.current.clear();
        const patchPending = (prev: UiItem[]) =>
          prev.map((it) => {
            if (it.kind === "choice" && !it.resolved) return { ...it, resolved: "cancelled" as const };
            if (it.kind === "confirm" && !it.resolved) return { ...it, resolved: "rejected" as const };
            return it;
          });
        if (runSessionRef.current) runSessionRef.current.items = patchPending(runSessionRef.current.items);
        if (!runSessionRef.current || currentConvIdRef.current === runSessionRef.current.convId) {
          setItems((prev) => patchPending(prev));
        }
        if (streamFlushRef.current != null) {
          window.clearTimeout(streamFlushRef.current);
          streamFlushRef.current = null;
        }
        if (currentAssistantIdRef.current) {
          updateRunItem(currentAssistantIdRef.current, {
            ...(streamTextRef.current ? { text: streamTextRef.current } : {}),
            streaming: false,
          } as Partial<UiItem>);
          currentAssistantIdRef.current = null;
        }
        if (!abort.signal.aborted && !runFailed) {
          const extras = hitRoundLimit || incomplete
            ? [{ label: "继续", prompt: "继续完成刚才未完成的任务。", run: true }]
            : pendingFollowupsRef.current || heuristicFollowups(toolsUsedRef.current);
          if (extras.length) addRunItem({ kind: "followups", id: uid(), actions: extras });
        }
        try {
          persistRunSession(true);
        } catch {
          /* 落盘失败不能阻断状态复位 */
        }
        if (runSessionRef.current && currentConvIdRef.current === runSessionRef.current.convId) {
          // itemsRef 只在渲染后更新，这里先同步成最终结果，避免下面的 flushPersist 用旧快照覆盖刚保存的内容
          itemsRef.current = runSessionRef.current.items;
        }
        runSessionRef.current = null;
        setRunNeedsInput(false);
        setRunningConvId(null);
        setRunProgress(null);
        setCurrentToolLabel(null);
        runningRef.current = false;
        setRunning(false);
        abortRef.current = null;
        if (currentConvIdRef.current && currentConvIdRef.current !== convId) {
          try {
            flushPersist();
          } catch {
            /* ignore */
          }
        }
        if (currentConvIdRef.current === convId) inputRef.current?.focus();
        settleRun();
      }
    },
    [addItem, addRunItem, buildExecCtx, flushPersist, persistRunSession, pendingImages, refreshExcelCtx, removeRunItem, setCurrentId, settings, updateRunItem]
  );

  /**
   * 发送一条预设指令（后续操作、上下文按钮、模板直接运行、审计修复）。
   * 这些指令不是用户在输入框里打的，所以不能顺手清掉输入框草稿，也不能把待发图片一起带走。
   */
  const sendPreset = useCallback(
    (prompt: string) => send(prompt, { preserveComposer: true, images: [] }),
    [send]
  );

  useEffect(() => {
    if (running || batchRunning || workbookBusy || !queued || queued.paused) return;
    if (queued.convId !== currentConvIdRef.current) return;
    const next = queued;
    setQueued(null);
    void send(next.text, { images: next.images, preserveComposer: true });
  }, [running, batchRunning, workbookBusy, queued, currentConvId, send]);

  const retryLast = useCallback(
    (errorId: string) => {
      if (runningRef.current || batchRunningRef.current) return;
      const errorIndex = itemsRef.current.findIndex((it) => it.id === errorId);
      if (errorIndex < 0 || itemsRef.current.slice(errorIndex + 1).some((it) => it.kind === "user")) return;
      const lastUser = itemsRef.current.slice(0, errorIndex).reverse().find((it) => it.kind === "user" && !it.steer);
      if (!lastUser || lastUser.kind !== "user") return;
      const nextItems = itemsRef.current.filter((it) => it.id !== errorId);
      itemsRef.current = nextItems;
      setItems(nextItems);
      // 失败请求所属的用户消息已经入历史，末尾也可能是补充说明。
      // 重试沿用同一任务及其检查点，不能再次追加原指令覆盖用户的新要求。
      const resumeTurn = historyRef.current.some((message) => message.role === "user");
      void send(lastUser.text, {
        skipUserBubble: true,
        resumeTurn,
        preserveComposer: true,
        images: (lastUser.images || []).map((img) => ({
          id: img.id,
          name: img.name,
          mime: "image/png",
          dataUrl: img.dataUrl,
          width: 0,
          height: 0,
        })),
      });
    },
    [send]
  );

  const resolvePendingUi = useCallback(() => {
    confirmResolversRef.current.forEach(({ resolve }) => resolve({ accepted: false }));
    confirmResolversRef.current.clear();
    choiceResolversRef.current.forEach((resolve) => resolve({ kind: "cancelled" }));
    choiceResolversRef.current.clear();
    const patch = (prev: UiItem[]) =>
      prev.map((it) => {
        if (it.kind === "choice" && !it.resolved) return { ...it, resolved: "cancelled" as const };
        if (it.kind === "confirm" && !it.resolved) return { ...it, resolved: "rejected" as const };
        return it;
      });
    const run = runSessionRef.current;
    if (run) {
      run.items = patch(run.items);
      if (currentConvIdRef.current === run.convId) setItems(run.items.slice());
    } else {
      setItems(patch);
    }
    setRunNeedsInput(false);
  }, []);

  const stop = useCallback(() => {
    setQueued((pending) => pending ? { ...pending, paused: true } : null);
    abortRef.current?.abort();
    batchAbortRef.current?.();
    resolvePendingUi();
  }, [resolvePendingUi]);

  /**
   * 离开一个还没发过消息的新对话时，它不会被保存、也回不去；在它上面排队的指令要放回输入框，
   * 否则会永远等在一个看不见的会话里。
   */
  const reclaimOrphanedQueue = useCallback(() => {
    const pending = queued;
    if (!pending || pending.convId !== currentConvIdRef.current || hasUserMessage(itemsRef.current)) return;
    setQueued(null);
    setInput((draft) => [pending.text, draft].filter((t) => t.trim()).join("\n"));
    setPendingImages((draft) => [...pending.images, ...draft]);
    flashComposer("排队的指令已放回输入框。");
  }, [queued, flashComposer]);

  const newChat = useCallback(() => {
    if (switchingRef.current || batchRunningRef.current || workbookBusyRef.current) return;
    reclaimOrphanedQueue();
    switchingRef.current = true;
    try {
      flushPersist();
      persistRunSession(true);
      const run = runSessionRef.current;
      setRunNeedsInput(!!run && hasPendingRunInput(run.items));
      // 排队中的指令属于后台那条会话，切走时保留（回到该会话后照常发送），和切换到其他历史会话行为一致
      historyRef.current = [];
      createdAtRef.current = Date.now();
      setItems([]);
      setCurrentId(null);
      setShowHistory(false);
      setRunProgress(null);
      setCurrentToolLabel(null);
      stickToBottomRef.current = true;
      setShowJumpLatest(false);
    } finally {
      switchingRef.current = false;
    }
  }, [flushPersist, persistRunSession, reclaimOrphanedQueue, setCurrentId]);

  const openConversation = useCallback(
    (id: string) => {
      if (batchRunningRef.current || workbookBusyRef.current) return;
      if (id === currentConvIdRef.current) {
        setShowHistory(false);
        return;
      }
      if (switchingRef.current) return;
      reclaimOrphanedQueue();
      switchingRef.current = true;
      try {
        flushPersist();
        persistRunSession(true);
        const run = runSessionRef.current;
        if (run && run.convId === id) {
          createdAtRef.current = run.createdAt;
          historyRef.current = run.history;
          setItems(run.items.slice());
          setRunProgress(run.progress);
          setCurrentToolLabel(run.toolLabel);
          setRunNeedsInput(hasPendingRunInput(run.items));
        } else {
          const data = loadConversation(id);
          if (!data) {
            setConversations(listConversations());
            return;
          }
          createdAtRef.current = data.createdAt;
          historyRef.current = data.messages;
          setItems(sanitizeItems(data.items));
          setRunProgress(null);
          setCurrentToolLabel(null);
          if (run) setRunNeedsInput(hasPendingRunInput(run.items));
        }
        setCurrentId(id);
        setShowHistory(false);
        stickToBottomRef.current = true;
        setShowJumpLatest(false);
      } finally {
        switchingRef.current = false;
      }
    },
    [flushPersist, persistRunSession, reclaimOrphanedQueue, setCurrentId]
  );

  const removeConversation = useCallback(
    async (id: string) => {
      if (batchRunningRef.current || workbookBusyRef.current) return;
      const run = runSessionRef.current;
      if (run && run.convId === id) {
        stop();
        await runDoneRef.current;
      }
      setQueued((pending) => (pending && pending.convId === id ? null : pending));
      if (currentConvIdRef.current === id) {
        if (switchingRef.current) return;
        switchingRef.current = true;
        try {
          historyRef.current = [];
          createdAtRef.current = Date.now();
          setItems([]);
          setCurrentId(null);
        } finally {
          switchingRef.current = false;
        }
      }
      try {
        deleteConversation(id);
      } catch {
        setStorageWarning(true);
      }
      setConversations(listConversations());
      setQueued((q) => (q?.convId === id ? null : q));
    },
    [setCurrentId, stop]
  );

  const registerBatchAbort = useCallback((fn: (() => void) | null) => {
    batchAbortRef.current = fn;
  }, []);

  const handleBatchRunningChange = useCallback((v: boolean) => {
    batchRunningRef.current = v;
    setBatchRunning(v);
    // 批量任务算一次任务：「同类操作不再询问」对所有文件生效，结束即失效
    autoApproveRef.current = new Set();
    if (v) {
      batchDoneRef.current = new Promise<void>((resolve) => {
        batchSettleRef.current = resolve;
      });
    } else {
      batchSettleRef.current?.();
      batchSettleRef.current = null;
    }
  }, []);

  const loadChatModels = useCallback(async (opts?: { silent?: boolean }) => {
    if (!settings.apiKey.trim() && !isLoopbackEndpoint(settings.baseUrl)) {
      setAvailableModels(mergeModelLists([settings.model]));
      setModelMenuStatus("error");
      setModelMenuError("当前 API 尚未配置 Key，可先手动输入模型名，或打开模型设置。");
      return;
    }
    const req = ++modelsReqRef.current;
    if (!opts?.silent) setModelMenuStatus("loading");
    setModelMenuError("");
    try {
      const models = await listModels(settings);
      if (req !== modelsReqRef.current) return;
      setAvailableModels(mergeModelLists(models, [settings.model]));
      setModelMenuStatus("idle");
      if (!models.length) setModelMenuError("当前 API 没有返回可选择的模型。可在下方手动输入模型名。");
    } catch (e: any) {
      if (req !== modelsReqRef.current) return;
      setAvailableModels((prev) => mergeModelLists(prev, getCachedModels(settings), [settings.model]));
      setModelMenuStatus("error");
      setModelMenuError(e?.message || String(e));
    }
  }, [settings]);

  const toggleModelMenu = useCallback(() => {
    if (showModelMenu) {
      setShowModelMenu(false);
      return;
    }
    const cached = mergeModelLists(getCachedModels(settings), availableModels, [settings.model]);
    if (cached.length) setAvailableModels(cached);
    setShowModelMenu(true);
    void loadChatModels({ silent: cached.length > 0 });
  }, [availableModels, loadChatModels, settings, showModelMenu]);

  const selectChatModel = useCallback(
    (model: string) => {
      const ctx = getCachedContextWindow(settings.baseUrl, model, settings.apiKey) || 0;
      const next = { ...settings, model, contextWindowTokens: 0 };
      try {
        saveSettings(next);
      } catch {
        addItem({ kind: "error", id: uid(), text: "模型未切换：本机存储不可用或空间不足，请稍后重试。" });
        return;
      }
      setSettings(next);
      setAvailableModels((prev) => mergeModelLists(prev, [model]));
      setShowModelMenu(false);
      setModelFilter("");
      addItem({
        kind: "notice",
        id: uid(),
        text: ctx
          ? `已切换到模型 ${model}（上下文 ${formatContextWindow(ctx)}）`
          : `已切换到模型 ${model}`,
      });
      // 未缓存的上下文窗口由上面监听 settings.model 的 effect 统一探测（带取消保护），这里不再重复发请求
    },
    [addItem, settings]
  );

  const handleUndo = useCallback(
    async (entry: UndoEntry) => {
      if (running || batchRunning || undoBusyRef.current) return;
      const stack = getUndoStack();
      const idx = stack.findIndex((item) => item.id === entry.id);
      if (idx < 0) {
        addItem({ kind: "error", id: uid(), text: "撤销失败: 该操作已不在撤销记录中" });
        return;
      }
      undoBusyRef.current = true;
      setUndoBusy(true);
      try {
        const msg = await withWorkbookAction(() => applyUndoTo(entry.id));
        addItem({ kind: "notice", id: uid(), text: msg });
        refreshExcelCtx();
      } catch (e: any) {
        const fallback = checkpoint ? " 如需整段恢复，可使用下方“恢复到任务开始”。" : "";
        addItem({ kind: "error", id: uid(), text: `撤销失败: ${e?.message || String(e)}${fallback}` });
      } finally {
        undoBusyRef.current = false;
        setUndoBusy(false);
      }
    },
    [addItem, batchRunning, checkpoint, refreshExcelCtx, running, withWorkbookAction]
  );

  const handleExportWorkbook = useCallback(async () => {
    try {
      const result = await withWorkbookAction(() => toolExportFile({ format: "xlsx" }));
      addItem({
        kind: result.ok ? "notice" : "error",
        id: uid(),
        text: result.ok ? result.summary || "已开始下载" : result.error || "导出失败",
      });
    } catch (e: any) {
      addItem({ kind: "error", id: uid(), text: `导出失败: ${e?.message || String(e)}` });
    }
  }, [addItem, withWorkbookAction]);

  const handleDownloadAttachment = useCallback(
    async (name: string) => {
      try {
        const msg = await downloadAttachedFile(name);
        addItem({ kind: "notice", id: uid(), text: msg });
      } catch (e: any) {
        addItem({ kind: "error", id: uid(), text: e?.message || String(e) });
      }
    },
    [addItem]
  );

  /** 体检卡片 → 修复勾选的问题：组装带问题明细 JSON 的指令发给 agent */
  const handleAuditFix = useCallback(
    (item: Extract<UiItem, { kind: "audit" }>, selected: AuditIssue[]) => {
      if (selected.length === 0 || runningRef.current || batchRunningRef.current) return;
      const patched = itemsRef.current.map((it) =>
        it.id === item.id ? ({ ...it, dispatched: true } as UiItem) : it
      );
      itemsRef.current = patched;
      setItems(patched);
      const brief = selected
        .map((i) => `- [${i.id}] ${i.type}${i.column ? `（${i.column}）` : ""} × ${i.count}`)
        .join("\n");
      void sendPreset(
        [
          `请修复数据体检发现并已勾选的 ${selected.length} 个问题（目标 Sheet「${item.sheet}」，检查范围 ${item.range}）。`,
          "",
          "勾选的问题：",
          brief,
          "",
          "问题明细（JSON，含精确定位）：",
          JSON.stringify(selected),
          "",
          "要求：",
          "- 按每条 suggestion 与 detail 逐项修复（缺失值、重复行、类型、离群值、空格各自按建议处理），行号列号都是绝对位置。",
          "- 在原表上修复即可；破坏性操作照常等确认。",
          "- 完成后按问题类型逐项汇报处理了多少处、跳过了哪些及原因。",
        ].join("\n")
      );
    },
    [sendPreset]
  );

  const handleBatchFinish = useCallback(
    (summary: string) => {
      addItem({ kind: "notice", id: uid(), text: summary });
      refreshExcelCtx();
    },
    [addItem, refreshExcelCtx]
  );

  const handleRestoreCheckpoint = useCallback(async () => {
    if (running || batchRunning || cpBusy || !checkpoint) return;
    setCpBusy(true);
    try {
      const msg = await withWorkbookAction(() => restoreCheckpoint());
      clearUndoStack();
      addItem({ kind: "notice", id: uid(), text: msg });
      refreshExcelCtx();
    } catch (e: any) {
      const location = e?.debugInfo?.errorLocation;
      addItem({ kind: "error", id: uid(), text: `恢复检查点失败: ${e?.message || String(e)}${location ? `（${location}）` : ""}` });
    } finally {
      setCpBusy(false);
    }
  }, [addItem, batchRunning, checkpoint, cpBusy, refreshExcelCtx, running, withWorkbookAction]);

  const handleDiscardCheckpoint = useCallback(async () => {
    if (running || batchRunning || cpBusy || !checkpoint) return;
    setCpBusy(true);
    try {
      await withWorkbookAction(() => discardCheckpoint());
      addItem({ kind: "notice", id: uid(), text: "已丢弃任务检查点备份。" });
    } catch (e: any) {
      addItem({ kind: "error", id: uid(), text: `丢弃检查点失败: ${e?.message || String(e)}` });
    } finally {
      setCpBusy(false);
    }
  }, [addItem, batchRunning, checkpoint, cpBusy, running, withWorkbookAction]);

  /**
   * 排队：等正在跑的任务（本会话、其他会话或批量）结束后，在当前会话发送。
   * 排不进去时要说明原因，不能静默吞掉回车。
   */
  const queueNextPrompt = () => {
    const text = input.trim();
    if (!text && pendingImages.length === 0) return;
    if (queued) {
      flashComposer("已有一条排队指令，取消或编辑它后再添加。当前输入已保留。");
      return;
    }
    let convId = currentConvIdRef.current;
    if (!convId) {
      convId = uid();
      createdAtRef.current = Date.now();
      setCurrentId(convId);
    }
    setQueued({ convId, text, images: pendingImages.slice() });
    setInput("");
    setPendingImages([]);
    if (runningRef.current && runningConvId !== convId) {
      const title = conversations.find((c) => c.id === runningConvId)?.title || "另一会话";
      flashComposer(`已排队：「${title}」完成后在本会话发送。`);
    } else if (batchRunningRef.current) {
      flashComposer("已排队：批量任务完成后发送。");
    }
  };

  /** 把输入作为补充说明插入正在查看的任务，模型在下一步之前就能看到 */
  const steerCurrentRun = () => {
    const text = input.trim();
    if (!text || !runningRef.current || runningConvId !== currentConvIdRef.current) return;
    const id = uid();
    steerRef.current.push({ id, text });
    addRunItem({ kind: "user", id, text, steer: true, pending: true });
    pushPromptHistory(text);
    promptHistIdxRef.current = -1;
    setInput("");
    stickToBottomRef.current = true;
    setShowJumpLatest(false);
    if (runNeedsInput) flashComposer("上方卡片还在等你处理；这条补充会在它之后送达。");
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.nativeEvent.isComposing || e.nativeEvent.keyCode === 229) return;
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      if (running && runningConvId === currentConvIdRef.current) {
        // 带图片的补充需要作为完整消息发送，只能排队
        if (pendingImages.length) queueNextPrompt();
        else steerCurrentRun();
        return;
      }
      if (running || batchRunning) {
        queueNextPrompt();
        return;
      }
      if (workbookBusy) {
        if (input.trim()) flashComposer("正在处理工作簿，完成后再发送。当前输入已保留。");
        return;
      }
      send(input);
      return;
    }
    if (e.altKey || e.ctrlKey || e.metaKey) return;
    const el = inputRef.current;
    if (!el) return;
    const hist = loadPromptHistory();
    if (!hist.length) return;
    const atStart = el.selectionStart === 0 && el.selectionEnd === 0;
    const atEnd = el.selectionStart === input.length && el.selectionEnd === input.length;
    if (e.key === "ArrowUp" && (input === "" || atStart)) {
      e.preventDefault();
      if (promptHistIdxRef.current === -1) promptDraftRef.current = input;
      const next = Math.min(hist.length - 1, promptHistIdxRef.current + 1);
      promptHistIdxRef.current = next;
      setInput(hist[hist.length - 1 - next]);
    } else if (e.key === "ArrowDown" && promptHistIdxRef.current >= 0 && atEnd) {
      e.preventDefault();
      const next = promptHistIdxRef.current - 1;
      promptHistIdxRef.current = next;
      setInput(next < 0 ? promptDraftRef.current : hist[hist.length - 1 - next]);
    }
  };

  const viewingActiveRun = running && runningConvId === currentConvId;
  const backgroundRun = running && !!runningConvId && runningConvId !== currentConvId;
  const runningConvTitle =
    conversations.find((c) => c.id === runningConvId)?.title || "另一会话";
  // 模型在思考（没有流式文本、没有工具在跑、也没有等待用户操作的卡片）时显示指示器
  const lastItem = items[items.length - 1];
  const latestRetryItem = [...items].reverse().find((item) =>
    (item.kind === "user" && !item.steer) || (item.kind === "error" && item.retryable)
  );
  const taskStart = lastTaskStart(items);
  // 正在看的运行中任务保持展开；其余任务的执行过程折叠
  const chatNodes = groupChatItems(items, viewingActiveRun ? Math.max(0, taskStart) : items.length);
  const stepsDone = viewingActiveRun
    ? items.slice(taskStart + 1).filter((it) => it.kind === "tool" && it.status !== "running").length
    : 0;
  const showRunStatus = viewingActiveRun && (runNeedsInput || showJumpLatest || runProgress?.phase === "retry");
  const hasDraft = !!input.trim() || pendingImages.length > 0;
  const hasAttachStrip =
    extFiles.length > 0 || !!extSelection || pendingImages.length > 0 || artifacts.length > 0;
  const showThinking =
    viewingActiveRun &&
    !(
      lastItem &&
      ((lastItem.kind === "assistant" && lastItem.streaming) ||
        (lastItem.kind === "tool" && lastItem.status === "running") ||
        (lastItem.kind === "confirm" && !lastItem.resolved) ||
        (lastItem.kind === "choice" && !lastItem.resolved))
    );
  // 每次重新进入「思考中」都从零计时，等久了指示器上会显示已用时间
  if (showThinking && !wasThinkingRef.current) thinkingSinceRef.current = Date.now();
  wasThinkingRef.current = showThinking;
  // 末尾几项会在 items 不变时继续长高（思考指示器展开、回复结束后出现操作栏、追问建议入场、图片加载）；
  // 观察它们的尺寸并跟住底部，否则最后一项会被输入区挡住一截。只观察末尾，长对话不增加开销。
  // Presence 会晚一次渲染才插入节点，所以随列表子节点变化重新挑选末尾，而不是依赖 React 状态。
  useEffect(() => {
    const list = listRef.current;
    if (!list) return;
    const resize = new ResizeObserver(() => {
      if (!stickToBottomRef.current || scrollRafRef.current != null) return;
      scrollRafRef.current = requestAnimationFrame(() => { scrollRafRef.current = null; scrollChatToBottom(); });
    });
    const observeTail = () => {
      resize.disconnect();
      for (const child of Array.from(list.children).slice(-3)) {
        // Presence / Collapse 外层是 display: contents，没有盒子，观察其内容
        const boxes = getComputedStyle(child).display === "contents" ? Array.from(child.children) : [child];
        boxes.forEach((box) => resize.observe(box));
      }
    };
    observeTail();
    const mutations = new MutationObserver(observeTail);
    mutations.observe(list, { childList: true });
    return () => { mutations.disconnect(); resize.disconnect(); };
  }, [scrollChatToBottom]);

  const hasItems = items.length > 0;
  // 输入框上方的情境操作：随选区内容变化（有错误 → 修复；有公式 → 解释 / 检查；普通数据 → 解释选区）。
  // 只读的直接执行；会改表的只填入输入框，由用户确认后发送。
  const contextActions = useMemo(() => {
    const { activeSheet, selection, hasError, hasFormula } = excelCtx;
    const actions: ContextAction[] = [];
    if (activeSheet) {
      const where = `${activeSheet}!${selection}`;
      if (hasError) {
        actions.push({ label: "修复错误", mode: "fill", tone: "danger", prompt: `当前选区 ${where} 里有公式错误。请先 inspect_range 查看公式和错误类型，找出原因并直接修好（优先原地修正，修之前用一两句话说明）。` });
      } else if (hasFormula) {
        actions.push(
          { label: "解释公式", mode: "run", prompt: `请用 inspect_range 查看当前选区 ${where} 的公式，解释它在算什么、引用了哪些单元格、有没有风险。只解释，不要修改。` },
          { label: "检查公式", mode: "fill", prompt: `请检查当前选区 ${where} 的公式是否正确、能否简化或写得更稳妥；有问题就直接修好。` },
        );
      } else if (selection) {
        actions.push({ label: "解释选区", mode: "run", prompt: `请查看当前选区 ${where} 的内容和结构，用几句话说明它是什么、有没有明显问题。只说明，不要修改。` });
      }
      if (!hasError) {
        actions.push({ label: "诊断全表", mode: "run", prompt: `请用 diagnose_errors 扫描 ${activeSheet} 整个已用区域的公式错误，把错误按类型和公式模式聚类，解释每一类错误的根因并给出修复建议。先不要直接修改，等我确认后再修。` });
      }
    }
    // 对话进行中时，把选区匹配到的记忆/技能建议并进同一行（欢迎页里已经单独展示过）
    if (hasItems && suggestionsFromContext) {
      for (const p of suggestedPrompts) {
        if (p.source === "selection") continue;
        actions.push({ label: p.label, prompt: p.prompt, mode: "fill", tag: SUGGESTION_TAGS[p.source] });
      }
    }
    return actions;
  }, [excelCtx, hasItems, suggestionsFromContext, suggestedPrompts]);

  /** 预设指令只填入输入框（替换当前内容），由用户检查、补充后再发送 */
  const fillComposer = (prompt: string) => {
    setInput(prompt);
    inputRef.current?.focus();
  };

  /**
   * 把引用（文件编号、选区地址、最近产出）插到光标处，而不是一律接在末尾；前后不是空白时补一个空格。
   * 点胶囊时输入框已失焦，但光标位置还保留在 selectionStart/End 上。
   */
  const insertIntoComposer = (snippet: string) => {
    const el = inputRef.current;
    const value = el?.value ?? input;
    const start = el?.selectionStart ?? value.length;
    const end = el?.selectionEnd ?? value.length;
    const before = value.slice(0, start);
    const after = value.slice(end);
    const head = before + (before && !/\s$/.test(before) ? " " : "") + snippet + (after && !/^\s/.test(after) ? " " : "");
    pendingCaretRef.current = head.length;
    setInput(head + after);
    promptHistIdxRef.current = -1;
  };

  const renderItem = (item: UiItem) => {
    switch (item.kind) {
      case "user":
        return <UserBubble key={item.id} text={item.text} images={item.images} steer={item.steer} pending={item.pending} />;
      case "assistant":
        return <AssistantBubble key={item.id} text={item.text} streaming={item.streaming} interim={item.interim} />;
      case "tool":
        return <ToolCard key={item.id} item={item} />;
      case "confirm":
        return <ConfirmCard key={item.id} item={item} onResolve={resolveConfirm} />;
      case "choice":
        return <ChoiceCard key={item.id} item={item} onResolve={resolveChoice} />;
      case "error": {
        // 只有最近一次失败还能接着处理；更早的报错只留文字，免得旧按钮把新任务重跑一遍
        const actionable = item.retryable && item.id === latestRetryItem?.id && !running && !batchRunning && !workbookBusy;
        return (
          <ErrorBubble
            key={item.id}
            text={item.text}
            detail={item.detail}
            action={actionable ? item.action : undefined}
            retryable={actionable}
            onRetry={() => retryLast(item.id)}
            onAction={(action) => (action === "settings" ? setShowSettings(true) : newChat())}
          />
        );
      }
      case "notice":
        return <NoticeBubble key={item.id} text={item.text} />;
      case "audit":
        return (
          <AuditCard
            key={item.id}
            sheet={item.sheet}
            range={item.range}
            reportSheet={item.reportSheet}
            issues={item.issues}
            dispatched={item.dispatched}
            running={running || batchRunning || workbookBusy}
            onFix={(selected) => handleAuditFix(item, selected)}
          />
        );
      case "followups":
        return (
          <FollowupBar
            key={item.id}
            actions={item.actions}
            disabled={running || batchRunning || workbookBusy}
            onPick={(action) => (action.run ? void sendPreset(action.prompt) : fillComposer(action.prompt))}
          />
        );
    }
  };

  const extSelFile = extSelection ? extFiles.find((f) => f.name === extSelection.file) : undefined;
  const selectionRef = excelCtx.selection
    ? `${sheetRefName(excelCtx.activeSheet)}!${excelCtx.selection}`
    : `「${excelCtx.activeSheet}」`;
  const filteredChatModels = modelFilter
    ? availableModels.filter((model) => model.toLowerCase().includes(modelFilter.toLowerCase()))
    : availableModels;
  const typedModel = modelFilter.trim();
  const canApplyTypedModel = !!typedModel && !availableModels.some((model) => model === typedModel);

  return (
    <div
      className={`app ${dragging ? "dragging" : ""}`}
      onDragEnter={(e) => {
        if (!e.dataTransfer.types.includes("Files")) return;
        e.preventDefault();
        dragDepthRef.current++;
        setDragging(true);
      }}
      onDragOver={(e) => { if (e.dataTransfer.types.includes("Files")) e.preventDefault(); }}
      onDragLeave={() => {
        dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
        if (dragDepthRef.current === 0) setDragging(false);
      }}
      onDrop={(e) => {
        if (!e.dataTransfer.types.includes("Files")) return;
        e.preventDefault();
        dragDepthRef.current = 0;
        setDragging(false);
        if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
      }}
      onPaste={(e) => {
        const dt = e.clipboardData;
        if (!dt) return;
        const images = collectClipboardImages(dt);
        if (images.length) {
          handleFiles(images);
          if (!dt.getData("text")) e.preventDefault();
        }
      }}
    >
      <header className="header">
        <div className="header-title">
          <span className="logo-badge"><IconSparkle size={15} /></span>
          <div className="header-text" ref={modelMenuRef}>
            <span className="header-name">Excel AI Agent</span>
            <button
              className={`header-model ${showModelMenu ? "active" : ""}`}
              title={`${settings.baseUrl} · ${settings.model}${displayedContextHint ? ` · 上下文 ${displayedContextHint}` : ""}\n点击切换当前 API 的模型`}
              onClick={toggleModelMenu}
              disabled={running || batchRunning || workbookBusy}
            >
              <span>{settings.model || "未配置模型"}</span>
              {displayedContextTokens > 0 && (
                <span className="header-ctx">{!settings.contextWindowTokens && detectedContext?.source === "estimate" ? "≈" : ""}{formatContextWindow(displayedContextTokens)}</span>
              )}
              <span className="header-model-arrow"><IconChevronDown size={12} /></span>
            </button>
            <Presence open={showModelMenu} duration={MOTION.duration.exit} focus="popover">
              <div className="chat-model-menu" onMouseDown={(e) => e.stopPropagation()}>
                <div className="chat-model-menu-head">
                  <span>当前 API 的模型</span>
                  <button disabled={modelMenuStatus === "loading"} onClick={() => void loadChatModels()}>
                    {modelMenuStatus === "loading" ? "获取中…" : "重新获取"}
                  </button>
                </div>
                <input
                  className="chat-model-search"
                  value={modelFilter}
                  onChange={(e) => setModelFilter(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key !== "Enter") return;
                    e.preventDefault();
                    const typed = modelFilter.trim();
                    if (!typed) return;
                    const exact = availableModels.find((model) => model.toLowerCase() === typed.toLowerCase());
                    selectChatModel(exact || typed);
                  }}
                  placeholder="搜索或输入模型名后回车"
                  autoFocus
                />
                {modelMenuStatus === "loading" && availableModels.length === 0 && (
                  <div className="chat-model-message">正在从当前 API 获取模型…</div>
                )}
                {modelMenuError && <div className="chat-model-message error">{modelMenuError}</div>}
                {filteredChatModels.length > 0 && (
                  <div className="chat-model-list">
                    {filteredChatModels.slice(0, 200).map((model) => (
                      <button
                        key={model}
                        className={settings.model === model ? "active" : ""}
                        onClick={() => selectChatModel(model)}
                        title={model}
                      >
                        <span>{model}</span>
                        {!!getCachedContextWindow(settings.baseUrl, model, settings.apiKey) && (
                          <span className="model-ctx">
                            {formatContextWindow(getCachedContextWindow(settings.baseUrl, model, settings.apiKey))}
                          </span>
                        )}
                      </button>
                    ))}
                  </div>
                )}
                {canApplyTypedModel && (
                  <button className="chat-model-settings" onClick={() => selectChatModel(typedModel)}>
                    使用模型 {typedModel}
                  </button>
                )}
                {availableModels.length > 0 && filteredChatModels.length === 0 && !canApplyTypedModel && (
                  <div className="chat-model-message">没有匹配的模型</div>
                )}
                <button
                  className="chat-model-settings"
                  onClick={() => {
                    setShowModelMenu(false);
                    setShowSettings(true);
                  }}
                >
                  打开完整模型设置
                </button>
              </div>
            </Presence>
          </div>
        </div>
        <div className="header-actions">
          <button className="icon-btn" onClick={newChat} disabled={batchRunning || workbookBusy} title="新对话">
            <IconPlus />
          </button>
          <button
            className="icon-btn"
            onClick={() => {
              flushPersist();
              setConversations(listConversations());
              setShowHistory(true);
            }}
            title="历史对话"
            disabled={batchRunning || workbookBusy}
          >
            <IconChats />
          </button>
          <div className="menu-anchor" ref={menuRef}>
            <button
              className={`icon-btn ${showMenu ? "active" : ""}`}
              onClick={() => setShowMenu((v) => !v)}
              title="更多功能"
            >
              <IconDots />
            </button>
            <Presence open={showMenu} duration={MOTION.duration.exit} focus="popover">
              <div className="header-menu">
                <div className="menu-heading">工作流</div>
                {([
                  { label: "常用指令", desc: "保存并复用常做的指令", Icon: IconTemplate, open: () => setShowTemplates(true) },
                  { label: "批量任务", desc: batchRunning ? "执行中，点击查看进度" : "对多个文件执行同一指令", Icon: IconLayers, open: () => setShowBatch(true) },
                  { label: "操作记录", desc: "查看并撤回 AI 做过的改动", Icon: IconHistory, open: () => setShowTimeline(true) },
                ] as const).map(({ label, desc, Icon, open }) => (
                  <button
                    key={label}
                    className="menu-item"
                    aria-label={label}
                    onClick={() => {
                      setShowMenu(false);
                      open();
                    }}
                  >
                    <span className="menu-item-icon"><Icon /></span>
                    <span className="menu-item-text" aria-hidden="true">
                      <span className="menu-item-label">{label}</span>
                      <span className="menu-item-desc">{desc}</span>
                    </span>
                    {label === "批量任务" && batchRunning && <span className="menu-item-badge" aria-hidden="true" />}
                  </button>
                ))}
                <div className="menu-sep" />
                <div className="menu-heading">工作簿</div>
                <button
                  className="menu-item"
                  disabled={running || batchRunning || workbookBusy}
                  onClick={() => {
                    setShowMenu(false);
                    void handleExportWorkbook();
                  }}
                >
                  <span className="menu-item-icon"><IconDownload /></span>
                  <span className="menu-item-text">
                    <span className="menu-item-label">导出工作簿</span>
                    <span className="menu-item-desc">下载当前文件副本（xlsx）</span>
                  </span>
                </button>
                <div className="menu-sep" />
                <div className="menu-heading">个性化</div>
                <button
                  className="menu-item"
                  onClick={() => {
                    setShowMenu(false);
                    setShowMemory(true);
                  }}
                >
                  <span className="menu-item-icon"><IconBookmark /></span>
                  <span className="menu-item-text">
                    <span className="menu-item-label">长期记忆</span>
                    <span className="menu-item-desc">AI 记住的偏好与事实</span>
                  </span>
                </button>
                <button
                  className="menu-item"
                  onClick={() => {
                    setShowMenu(false);
                    setShowSkills(true);
                  }}
                >
                  <span className="menu-item-icon"><IconZap /></span>
                  <span className="menu-item-text">
                    <span className="menu-item-label">技能库</span>
                    <span className="menu-item-desc">沉淀下来的可复用操作流程</span>
                  </span>
                </button>
              </div>
            </Presence>
          </div>
          <button className="icon-btn" onClick={() => setShowSettings(true)} title="模型设置">
            <IconGear />
          </button>
        </div>
      </header>

      <div className="chat-list-wrap">
      {storageWarning && <div className="storage-warning" role="status">对话暂未保存，本机存储不可用或空间不足。关闭窗格前请复制需要保留的内容。</div>}
      {batchRunning && !showBatch && (
        <div className="bg-run-banner" role="status">
          <span>{hasPendingRunInput(items) ? "批量任务等待确认，请查看下方卡片" : "批量任务执行中"}</span>
          <div className="bg-run-banner-actions">
            <button className="btn tiny" onClick={() => setShowBatch(true)}>查看进度</button>
            <button className="btn tiny danger" onClick={stop}>停止</button>
          </div>
        </div>
      )}
      {backgroundRun && (
        <div className="bg-run-banner">
          <span>
            {runNeedsInput
              ? `「${runningConvTitle}」在等你确认后才能继续`
              : `「${runningConvTitle}」仍在后台执行`}
          </span>
          <div className="bg-run-banner-actions">
            {runningConvId && (
              <button type="button" className="btn tiny" onClick={() => openConversation(runningConvId)}>
                返回
              </button>
            )}
            <button type="button" className="btn tiny danger" onClick={stop}>
              停止
            </button>
          </div>
        </div>
      )}
      <WorkbookLinksContext.Provider value={workbookLinks}>
      <div className="chat-list" ref={listRef} onScroll={onListScroll}>
        {/* minOpenMs：启动时恢复历史会话会先渲染一帧欢迎页，这种情况直接切换、不播退场 */}
        <Presence open={items.length === 0} duration={MOTION.duration.standard} minOpenMs={MOTION.duration.welcome}>
          <WelcomePanel
            needsSetup={!settings.apiKey.trim() && !isLoopbackEndpoint(settings.baseUrl)}
            onSetup={() => setShowSettings(true)}
            suggestions={suggestedPrompts}
            contextLabel={suggestionsFromContext ? `根据选区 ${selProfile?.sheet}!${selProfile?.address}` : "推荐任务"}
            disabled={running || batchRunning || workbookBusy}
            onPick={(prompt) => { setInput(prompt); inputRef.current?.focus(); }}
            onTemplates={() => setShowTemplates(true)}
          />
        </Presence>
        {chatNodes.map((node) => {
          if (node.kind === "item") return renderItem(node.item);
          const tools = node.items.filter((it) => it.kind === "tool");
          const failed = tools.filter((it) => it.kind === "tool" && it.status === "error");
          return (
            <ProcessGroup
              key={`group-${node.items[0].id}`}
              steps={tools.length}
              failed={failed.length}
              failedContent={failed.map(renderItem)}
            >
              {node.items.map(renderItem)}
            </ProcessGroup>
          );
        })}
        <Collapse open={showThinking} duration={MOTION.duration.micro}>
          <ThinkingIndicator
            since={thinkingSinceRef.current}
            label={
              runProgress?.phase === "retry"
                ? "网络不稳，正在重试模型请求…"
                : stepsDone > 0
                  ? "思考下一步…"
                  : undefined
            }
          />
        </Collapse>
      </div>
      </WorkbookLinksContext.Provider>
      <Presence open={showJumpLatest}>
        <button type="button" className="jump-latest" onClick={jumpToLatest}>
          <IconArrowDown size={13} /> 回到最新
        </button>
      </Presence>
      </div>

      <RollbackBar
        undoEntries={undoEntries}
        checkpoint={checkpoint}
        cpBusy={cpBusy}
        disabled={running || batchRunning || workbookBusy}
        undoBusy={undoBusy}
        onUndo={handleUndo}
        onCheckpointAction={setCheckpointAction}
        onOpenTimeline={() => setShowTimeline(true)}
      />

      <div className="composer">
        {/* 附件、框选、图片和最近产出合成一行，横向滚动，不再各占一行挤压对话区 */}
        <div className={`attach-strip${hasAttachStrip ? "" : " empty"}`}>
          {restoredFiles > 0 && extFiles.length > 0 && (
            <span className="restored-tag" title="上次附加的文件已恢复，关闭任务窗格后也会保留">
              已恢复 {restoredFiles} 个附件
              <button className="file-chip-x" onClick={() => setRestoredFiles(0)} title="知道了">
                <IconClose size={10} />
              </button>
            </span>
          )}
          <MotionList className="file-chips">
            {extFiles.map((f) => (
              <span
                key={f.name}
                className="file-chip clickable"
                title={`${f.caption}\n${f.sheets.map((s) => `${s.name} (${s.rows.length}行)`).join("\n")}\n点击编号可插入「${f.alias}」，点击文件名可预览`}
              >
                <button
                  className="file-chip-alias"
                  title={`插入「${f.alias}」，对 AI 说这个即可`}
                  onClick={() => insertIntoComposer(f.alias)}
                >
                  {f.alias}
                </button>
                <button className="file-chip-name" onClick={() => setPreviewFile(f.name)}>
                  <IconFile size={12} /> {f.name}
                </button>
                <button className="file-chip-x" onClick={() => void handleDownloadAttachment(f.name)} title="下载原文件">
                  <IconDownload size={11} />
                </button>
                <button className="file-chip-x" onClick={() => removeExternalFile(f.name)} title="移除"><IconClose size={11} /></button>
              </span>
            ))}
            {extSelection && (
              <span key="external-selection" className="file-chip selection" title="已框选的外部文件区域，会随消息发给 AI">
                {extSelFile && <span className="file-chip-alias">{extSelFile.alias}</span>}
                <button className="file-chip-name" onClick={() => setPreviewFile(extSelection.file)}>
                  <IconSelection size={12} /> {extSelection.file} / {extSelection.sheet}!{extSelection.range}
                </button>
                <button className="file-chip-x" onClick={() => setExternalSelection(null)} title="取消引用"><IconClose size={11} /></button>
              </span>
            )}
            {pendingImages.map((img) => (
              <span key={img.id} className="image-chip" title={img.name}>
                <img src={img.dataUrl} alt={img.name} />
                <button
                  className="file-chip-x"
                  onClick={() => setPendingImages((prev) => prev.filter((x) => x.id !== img.id))}
                  title="移除"
                >
                  <IconClose size={11} />
                </button>
              </span>
            ))}
          </MotionList>
          {artifacts.slice(0, 6).map((a) => (
            <button
              key={a.id}
              className="artifact-chip"
              title={`AI 最近生成的产出，点击把引用插入输入框\n${a.sheet}${a.address ? "!" + a.address : ""} · ${a.brief}`}
              onClick={() => insertIntoComposer(input.trim() ? `「${a.name}」` : `把「${a.name}」`)}
            >
              {a.kind === "chart" ? <IconChart size={12} /> : a.kind === "pivot" ? <IconPivot size={12} /> : a.kind === "table" ? <IconSheet size={12} /> : <IconFile size={12} />} {a.name}
            </button>
          ))}
        </div>
        {queued && queued.convId === currentConvId && (
          <div className="queue-chip">
            <span>
              {queued.paused ? "上一任务未完成，已暂停排队：" : "排队下一条："}
              {queued.text.length > 48 ? queued.text.slice(0, 48) + "…" : queued.text || `（${queued.images.length} 张图片）`}
            </span>
            <button className="btn tiny" onClick={() => {
              setInput((draft) => [queued.text, draft].filter(Boolean).join("\n"));
              setPendingImages((draft) => [...queued.images, ...draft]);
              setQueued(null);
              inputRef.current?.focus();
            }} title="取回指令">编辑</button>
            <button className="file-chip-x" onClick={() => setQueued(null)} title="取消排队">
              <IconClose size={11} />
            </button>
          </div>
        )}
        {workbookBusy && <div className="run-status" role="status"><span className="spinner" />正在处理工作簿，请稍候…</div>}
        {/* 对话里已经有思考提示和工具卡片；这里只在需要你处理、在重试、或你翻看历史看不到底部时出现 */}
        {showRunStatus && (
          <div className="run-status" role="status">
            <span className={runNeedsInput ? "run-waiting-dot" : "spinner"} />
            <span>
              {runNeedsInput
                ? "等待你的确认或选择"
                : `${currentToolLabel || (runProgress?.phase === "retry" ? "网络不稳，正在重试" : "思考中")}${stepsDone ? ` · 已完成 ${stepsDone} 步` : ""}`}
            </span>
          </div>
        )}
        <TaskLauncher
          disabled={running || batchRunning || workbookBusy}
          contextActions={contextActions}
          onRun={(prompt) => void sendPreset(prompt)}
          onTemplates={() => setShowTemplates(true)}
          onPick={fillComposer}
        />
        <div className={`input-row${(running || batchRunning) && hasDraft ? " has-run-actions" : ""}`}>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.xlsm,.csv,.png,.jpg,.jpeg,.gif,.webp,.bmp,image/*"
            multiple
            style={{ display: "none" }}
            onChange={(e) => {
              if (e.target.files?.length) handleFiles(e.target.files);
              e.target.value = "";
            }}
          />
          <button
            className="attach-btn"
            onClick={() => fileInputRef.current?.click()}
            title="附加 Excel/CSV 或图片（也可拖入窗口，或 Ctrl+V 粘贴截图）"
          >
            <IconClip />
          </button>
          <button
            type="button"
            className={`ctx-chip sheet${excelCtx.activeSheet ? "" : " pending"}`}
            disabled={!excelCtx.activeSheet}
            onClick={() => insertIntoComposer(selectionRef)}
            title={excelCtx.activeSheet ? `当前上下文：${selectionRef}，会随消息发给 AI\n点击把这个地址插入输入框` : "正在读取当前工作表…"}
          >
            <IconSheet size={12} />
            <span className="ctx-chip-text">
              {excelCtx.activeSheet || "读取工作表…"}
              {excelCtx.selection && <span className="ctx-chip-range">{excelCtx.selection}</span>}
            </span>
          </button>
          <textarea
            ref={inputRef}
            value={input}
            rows={1}
            aria-label="任务指令"
            placeholder={
              viewingActiveRun
                ? "任务执行中…输入补充或纠正，Enter 后 AI 在下一步前就能看到"
                : running
                  ? "另一会话仍在执行。这里输入的指令会排队，等它完成后在本会话发送"
                  : batchRunning
                    ? "批量任务执行中…这里输入的指令会排队，完成后发送"
                    : "描述任务，例如：按月份汇总销售额…"
            }
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={onKeyDown}
          />
          {viewingActiveRun && input.trim() && pendingImages.length === 0 && (
            <button className="queue-next-btn steer" onClick={steerCurrentRun} aria-label="补充到当前任务" title="插入正在执行的任务，AI 在下一步之前就能看到">补充</button>
          )}
          {(running || batchRunning) && hasDraft && (
            <button
              className="queue-next-btn"
              disabled={!!queued}
              onClick={queueNextPrompt}
              aria-label="排队发送"
              title={queued ? "已有指令排队，取消后可重新添加" : "等正在执行的任务完成后再发送"}
            >
              排队
            </button>
          )}
          {viewingActiveRun ? (
            <button className="send-btn stop" onClick={stop} title="停止">
              <IconStop />
            </button>
          ) : !((running || batchRunning) && hasDraft) && (
            <button
              className="send-btn"
              onClick={() => send(input)}
              disabled={running || batchRunning || workbookBusy || !hasDraft}
              title="发送"
            >
              <IconSend size={15} />
            </button>
          )}
        </div>
        <div className="composer-hint" role="status">
          {composerFlash ? (
            <span className="composer-flash">{composerFlash}</span>
          ) : viewingActiveRun ? (
            <><span>Enter 补充到当前任务 · 排队 = 完成后再发</span><span>Shift+Enter 换行</span></>
          ) : (
            <><span>Enter 发送 · Shift+Enter 换行</span><span>↑ 历史指令</span></>
          )}
        </div>
      </div>

      <Presence open={dragging} duration={MOTION.duration.micro}>
        <div className="drop-overlay">
          <div className="drop-inner">
            <div className="drop-icon"><IconUpload size={26} /></div>
            <div>松开以附加 Excel / CSV / 图片</div>
          </div>
        </div>
      </Presence>

      <Presence open={!!previewFile} focus="dialog">
        {previewFile && (
          <FilePreview
            fileName={previewFile}
            key={previewFile}
            busy={running || batchRunning || workbookBusy}
            onImport={(args) => withWorkbookAction(() => toolImportExternalSheet(args))}
            onClose={() => setPreviewFile(null)}
            onNotice={(text, isError) =>
              addItem(isError ? { kind: "error", id: uid(), text } : { kind: "notice", id: uid(), text })
            }
          />
        )}
      </Presence>

      <Presence open={showHistory} focus="dialog">
        <HistoryPanel
          conversations={conversations}
          currentId={currentConvId}
          runningId={runningConvId}
          onSelect={openConversation}
          onDelete={removeConversation}
          onNew={newChat}
          onClose={() => setShowHistory(false)}
        />
      </Presence>

      <Presence open={showMemory} focus="dialog">
        <MemoryPanel onClose={() => setShowMemory(false)} />
      </Presence>

      <Presence open={showSkills} focus="dialog">
        <SkillsPanel
          onClose={() => setShowSkills(false)}
          onUse={(skill) => {
            setInput(`请按技能「${skill.name}」处理当前工作簿。\n\n执行流程：\n${skill.steps}\n\n本次具体要求：`);
            setShowSkills(false);
            window.setTimeout(() => inputRef.current?.focus(), 0);
          }}
        />
      </Presence>

      <Presence open={showTemplates} focus="dialog">
        <PromptTemplatesPanel
          currentPrompt={input}
          busy={running || batchRunning || workbookBusy}
          onClose={() => setShowTemplates(false)}
          onUse={(prompt) => {
            setInput(prompt);
            setShowTemplates(false);
            window.setTimeout(() => inputRef.current?.focus(), 0);
          }}
          onRun={(prompt) => {
            setShowTemplates(false);
            void sendPreset(prompt);
          }}
        />
      </Presence>

      <Presence open={showTimeline} focus="dialog">
        <TimelinePanel
          onClose={() => setShowTimeline(false)}
          onUndoEntry={handleUndo}
          undoBusy={running || batchRunning || workbookBusy || undoBusy}
        />
      </Presence>

      {batchOpenedRef.current && (
        <Presence open={showBatch} focus="dialog" keepMounted>
          <BatchPanel
            settings={settings}
            mainRunning={running || batchRunning || workbookBusy}
            buildExecCtx={buildExecCtx}
            onClose={() => setShowBatch(false)}
            onFinish={handleBatchFinish}
            onRunningChange={handleBatchRunningChange}
            onAbortCleanup={resolvePendingUi}
            registerAbort={registerBatchAbort}
          />
        </Presence>
      )}

      <Presence open={!!checkpointAction} focus="dialog">
        {checkpointAction && (
        <div className="settings-overlay" onClick={() => setCheckpointAction(null)}>
          <section className="settings-panel" role="alertdialog" aria-modal="true"
            aria-labelledby="checkpoint-confirm-title" aria-describedby="checkpoint-confirm-detail"
            onClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => {
              if (event.key === "Escape") { event.stopPropagation(); setCheckpointAction(null); }
              if (event.key === "Tab") {
                const buttons = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>("button"));
                const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
                event.preventDefault();
                buttons[(index + (event.shiftKey ? -1 : 1) + buttons.length) % buttons.length]?.focus();
              }
            }}>
            <div className="settings-header"><h2 id="checkpoint-confirm-title">
              {checkpointAction === "restore" ? "恢复到任务开始？" : "丢弃任务备份？"}
            </h2></div>
            <p id="checkpoint-confirm-detail">
              {checkpointAction === "restore"
                ? "将恢复本次任务开始前的工作簿，撤销之后的改动，包括新建工作表和删除的行。"
                : "丢弃备份后，将无法整段恢复到本次任务开始。当前表格数据不会改变。"}
            </p>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
              <button className="btn" autoFocus onClick={() => setCheckpointAction(null)}>取消</button>
              <button className="btn primary" disabled={running || batchRunning || workbookBusy || cpBusy} onClick={() => {
                const action = checkpointAction;
                setCheckpointAction(null);
                void (action === "restore" ? handleRestoreCheckpoint() : handleDiscardCheckpoint());
              }}>{checkpointAction === "restore" ? "确认恢复" : "确认丢弃"}</button>
            </div>
          </section>
        </div>
        )}
      </Presence>

      <Presence open={showSettings} focus="dialog">
        <SettingsPanel
          settings={settings}
          onClose={() => setShowSettings(false)}
          onSave={(s) => {
            saveSettings(s);
            setDetectedContext(getCachedContextInfo(s.baseUrl, s.model, s.apiKey) || null);
            setSettings(s);
            setAvailableModels(mergeModelLists(getCachedModels(s), [s.model]));
          }}
        />
      </Presence>
    </div>
  );
}
