/** OpenAI 兼容的消息与工具调用类型 */

export interface ToolCallFn {
  name: string;
  arguments: string; // JSON 字符串
}

export interface ToolCall {
  id: string;
  type: "function";
  function: ToolCallFn;
}

/** OpenAI 兼容的多模态内容块（文本 + 图片） */
export type ContentPart =
  | { type: "text"; text: string }
  | { type: "image_url"; image_url: { url: string; detail?: "auto" | "low" | "high" } };

export interface ChatMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string | null | ContentPart[];
  tool_calls?: ToolCall[];
  tool_call_id?: string;
}

export interface ChatImage {
  id: string;
  name: string;
  dataUrl: string;
}

export interface ToolDefinition {
  type: "function";
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

/** 工具执行结果 */
export interface ToolResult {
  ok: boolean;
  /** 返回给模型的内容（会被 JSON 序列化） */
  data?: unknown;
  error?: string;
  /** 展示给用户的一句话摘要 */
  summary?: string;
  /** 本次操作是否可撤销 */
  undoable?: boolean;
  /** 任务完成后展示给用户的后续操作 */
  followups?: FollowupAction[];
  /** 数据体检结果：UI 据此渲染可勾选修复的问题卡片 */
  audit?: AuditReport;
}

/** 数据体检报告 */
export interface AuditReport {
  sheet: string;
  range: string;
  reportSheet?: string;
  issues: AuditIssue[];
}

/** 数据体检发现的一类问题 */
export interface AuditIssue {
  id: string;
  type: "missing" | "duplicate" | "type" | "outlier" | "whitespace";
  severity: "high" | "medium" | "low";
  /** 涉及列（表头名） */
  column?: string;
  /** 问题单元格/行数（绝对行号，供修复定位） */
  count: number;
  /** 给用户看的一句示例 */
  sample: string;
  /** 修复建议 */
  suggestion: string;
  /** 机器可读的定位与上下文，随修复指令发回模型 */
  detail: Record<string, unknown>;
}

export interface FollowupAction {
  label: string;
  prompt: string;
  /** 点击后直接执行（只用于接着跑刚才的任务，如「继续」）；默认只填入输入框 */
  run?: boolean;
}

/** 确认卡片的回答：拒绝时用户可以附一句说明，交给模型调整做法 */
export interface ConfirmAnswer {
  accepted: boolean;
  note?: string;
}

/** UI 中的消息条目 */
export type UiItem =
  | {
      kind: "user";
      id: string;
      text: string;
      images?: ChatImage[];
      /** 任务执行中插入的补充说明（不是新任务） */
      steer?: boolean;
      /** 补充说明还没送到模型（等当前步骤结束） */
      pending?: boolean;
    }
  | {
      kind: "assistant";
      id: string;
      text: string;
      streaming?: boolean;
      /** 调用工具前的过程叙述，不是最终结论；界面弱化显示 */
      interim?: boolean;
    }
  | {
      kind: "tool";
      id: string;
      name: string;
      label: string;
      argsSummary: string;
      status: "running" | "ok" | "error";
      detail?: string;
      undoable?: boolean;
      /** 开始执行的时间，跑得久时卡片上显示已用时间 */
      startedAt?: number;
    }
  | {
      kind: "confirm";
      id: string;
      title: string;
      detail: string;
      /** 请求确认的工具名，用于「本次任务同类操作不再询问」 */
      tool?: string;
      resolved?: "accepted" | "rejected";
      /** 因本次任务已允许同类操作而自动通过 */
      auto?: boolean;
      /** 拒绝时用户给的说明 */
      note?: string;
    }
  | {
      kind: "choice";
      id: string;
      title: string;
      question?: string;
      options: ChoiceOption[];
      /** custom=true 表示用户没选给定选项，label 是用户自己写的做法 */
      resolved?: { id: string; label: string; custom?: boolean } | "skipped" | "cancelled";
    }
  | {
      kind: "error";
      id: string;
      text: string;
      retryable?: boolean;
      /** 服务商原话等次要信息，显示在结论下面 */
      detail?: string;
      /** 除重试之外的出路：去设置里改 / 新建对话 */
      action?: "settings" | "new-chat";
    }
  | { kind: "notice"; id: string; text: string }
  | { kind: "followups"; id: string; actions: FollowupAction[] }
  | {
      kind: "audit";
      id: string;
      sheet: string;
      range: string;
      reportSheet?: string;
      issues: AuditIssue[];
      /** 是否已点过「修复选中项」 */
      dispatched?: boolean;
    };

export interface ChoiceOption {
  id: string;
  label: string;
  description?: string;
  recommended?: boolean;
}

export type ChoiceResolution =
  | { kind: "selected"; id: string; label: string }
  | { kind: "custom"; text: string }
  | { kind: "skipped" }
  | { kind: "cancelled" };

export interface AppSettings {
  baseUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxToolRounds: number;
  /** 单次模型回复的 max_tokens 上限；0 表示不传，由服务端决定 */
  maxOutputTokens: number;
  confirmDestructive: boolean;
  sendSelectionContext: boolean;
  /** 任务有多种做法时先让用户选方向 */
  askDirection: boolean;
  /** 把长期记忆注入上下文，并允许模型读写记忆 */
  useMemory: boolean;
  /** 任务完成后自动复盘，沉淀/更新技能库（自我进化） */
  autoLearnSkills: boolean;
  /** 动手改表前自动备份工作簿，支持整段回退（含删行/删表） */
  autoCheckpoint: boolean;
  /** 当前模型上下文窗口（token）。0 表示尚未探测 */
  contextWindowTokens: number;
}

export interface UndoEntry {
  id: string;
  description: string;
  sheet: string;
  address: string;
  /** 覆盖前的公式快照（含普通值） */
  formulas: (string | number | boolean)[][];
  /** 操作完成后的值/公式，用于撤销前检测用户后续编辑冲突 */
  expectedValues: (string | number | boolean)[][];
  expectedFormulas: (string | number | boolean)[][];
  /** 覆盖前的数字格式（清空「全部」时一并恢复） */
  numberFormats?: string[][];
  /** 操作完成后的数字格式，用于检测格式冲突 */
  expectedNumberFormats?: string[][];
  timestamp: number;
}

/** 操作时间线的条目类型 */
export type OpKind = "write" | "structure" | "format" | "import" | "chart" | "table" | "undo" | "other";

/** 操作时间线条目：记录每次改动的描述与前后对比 */
export interface OpLogEntry {
  id: string;
  timestamp: number;
  description: string;
  kind: OpKind;
  sheet: string;
  address?: string;
  /** 修改前的内容预览（文本网格） */
  before?: string;
  /** 修改后的内容预览（文本网格） */
  after?: string;
  /** 关联的撤销快照 id（可撤销的操作才有） */
  undoEntryId?: string;
}

/** 本会话 AI 产出的可引用结果（新表/区域/图表/透视等） */
export interface Artifact {
  id: string;
  name: string;
  kind: "sheet" | "range" | "chart" | "table" | "pivot";
  sheet: string;
  address?: string;
  brief: string;
  createdAt: number;
}
