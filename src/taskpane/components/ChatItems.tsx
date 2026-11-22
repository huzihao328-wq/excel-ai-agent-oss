import { KeyboardEvent, memo, ReactNode, useEffect, useId, useRef, useState } from "react";
import { AuditIssue, FollowupAction, UiItem } from "../../core/types";
import Markdown from "./Markdown";
import Collapse from "./Collapse";
import { animate, SPRING } from "../motion";
import { copyText } from "../clipboard";
import { IconAlert, IconCheck, IconChevronRight, IconClose, IconCompass, IconCopy, IconInfo, IconPlay, IconPulse, IconSparkle, splitLeadingGlyph } from "./Icons";

export const UserBubble = memo(function UserBubble({
  text,
  images,
  steer,
  pending,
}: {
  text: string;
  images?: { name: string; dataUrl: string }[];
  steer?: boolean;
  pending?: boolean;
}) {
  return (
    <div className={`msg-row user${steer ? " steer" : ""}`}>
      <div className={`bubble user-bubble${steer ? " steer" : ""}${pending ? " pending" : ""}`}>
        {steer && (
          <div className="steer-label">{pending ? "补充说明 · 当前步骤完成后送达" : "补充说明"}</div>
        )}
        {images && images.length > 0 && (
          <div className="msg-images">
            {images.map((img, i) => (
              <img key={`${i}-${img.name}`} className="msg-image" src={img.dataUrl} alt={img.name} title={img.name} />
            ))}
          </div>
        )}
        {text}
      </div>
    </div>
  );
});

export const AssistantBubble = memo(function AssistantBubble({ text, streaming, interim }: { text: string; streaming?: boolean; interim?: boolean }) {
  const [copied, setCopied] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const bubble = useRef<HTMLDivElement>(null);
  const resizeAnimation = useRef<Animation>();
  useEffect(() => () => resizeAnimation.current?.cancel(), []);
  // 流式阶段也要限制高度；否则某些啰嗦模型会让气泡持续增高并扰动滚动位置。
  // 过程叙述（interim）更早折叠，避免模型的「我将……」占满屏幕。
  const tooLong = text.length > (interim ? 240 : 1600);
  const collapsed = tooLong && !expanded;
  return (
    <div className={`msg-row assistant ${interim ? "interim" : ""}`}>
      <div ref={bubble} className={`bubble assistant-bubble ${collapsed ? "collapsed" : ""} ${interim ? "interim" : ""}`}>
        {!interim && <div className="assistant-label"><IconSparkle size={12} />表格助手</div>}
        {streaming ? <div className="markdown markdown-stream">{text}</div> : <Markdown text={text} />}
        {streaming && <span className="cursor-blink">▍</span>}
        {tooLong && (
          <button className="bubble-expand" aria-expanded={!collapsed} onClick={() => {
            const element = bubble.current;
            if (element) {
              const content = element.querySelector<HTMLElement>(".markdown");
              if (!content) return;
              const from = content.getBoundingClientRect().height;
              resizeAnimation.current?.cancel();
              element.classList.toggle("collapsed", expanded);
              const to = content.getBoundingClientRect().height;
              resizeAnimation.current = animate(content, [{ maxHeight: `${from}px` }, { maxHeight: `${to}px` }], SPRING.smooth.duration, SPRING.smooth.easing);
            }
            setExpanded((v) => !v);
          }}>
            {expanded ? "收起" : "展开全文"}
          </button>
        )}
        {!streaming && text.trim() && (
          <div className="bubble-actions">
            <button
              className={`bubble-copy ${copied ? "copied" : ""}`}
              title="复制回复内容"
              onClick={async () => {
                if (await copyText(text)) {
                  setCopied(true);
                  window.setTimeout(() => setCopied(false), 1500);
                }
              }}
            >
              {copied ? <IconCheck size={13} /> : <IconCopy size={13} />}
              <span>{copied ? "已复制" : "复制"}</span>
            </button>
          </div>
        )}
      </div>
    </div>
  );
});

/** 从 since 起已过去的整秒数，每秒刷新；active 为假时不计时 */
function useElapsedSeconds(since: number | undefined, active = true): number {
  const [now, setNow] = useState(() => Date.now());
  useEffect(() => {
    if (!since || !active) return;
    setNow(Date.now());
    const timer = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, [since, active]);
  return since && active ? Math.max(0, Math.floor((now - since) / 1000)) : 0;
}

function formatElapsed(seconds: number): string {
  return seconds < 60 ? `${seconds} 秒` : `${Math.floor(seconds / 60)} 分 ${String(seconds % 60).padStart(2, "0")} 秒`;
}

/** 工具报错给用户看的一句话：去掉「执行 xxx 出错:」前缀和调用栈，完整原文仍可展开查看 */
function toolErrorReason(detail?: string): string {
  const first = (detail || "").split("\n")[0].replace(/^执行\s+\S+\s+出错[:：]\s*/, "").trim();
  return first.length > 90 ? first.slice(0, 90) + "…" : first || "这一步没有成功";
}

export const ToolCard = memo(function ToolCard({ item }: { item: Extract<UiItem, { kind: "tool" }> }) {
  const [expanded, setExpanded] = useState(false);
  const detailId = useId();
  const icon = item.status === "running" ? <span className="spinner" /> : item.status === "ok" ? <IconCheck size={13} /> : <IconClose size={12} />;
  const open = expanded && !!item.detail;
  // 大表汇总、导入这类步骤可能要跑十几秒：超过 3 秒显示已用时间，看得出没卡住
  const elapsed = useElapsedSeconds(item.startedAt, item.status === "running");
  return (
    <div className={`tool-card ${item.status}`}>
      <button type="button" className="tool-card-head disclosure-button" aria-expanded={open} aria-controls={detailId} disabled={!item.detail} onClick={() => item.detail && setExpanded((v) => !v)}>
        <span className={`tool-status ${item.status}`}>{icon}</span>
        <span className="tool-label">{item.label}</span>
        <span className="tool-args" title={item.argsSummary}>{item.argsSummary}</span>
        {elapsed >= 3 && <span className="tool-elapsed">{formatElapsed(elapsed)}</span>}
        {item.detail && <span className={`tool-expand${open ? " open" : ""}`} aria-hidden="true"><IconChevronRight size={13} /></span>}
      </button>
      {item.status === "error" && !open && (
        <div className="tool-error-line">未成功：{toolErrorReason(item.detail)}</div>
      )}
      <Collapse open={open} id={detailId}><pre className="tool-detail">{item.detail}</pre></Collapse>
    </div>
  );
});

/**
 * 卡片里的一行补充输入（拒绝并说明 / 其他做法）：Enter 提交，Shift+Enter 换行，Esc 返回。
 * 输入法组字时的 Enter 不算提交。
 */
function InlineNote({
  label,
  placeholder,
  submitLabel,
  onSubmit,
  onCancel,
}: {
  label: string;
  placeholder: string;
  submitLabel: string;
  onSubmit: (text: string) => void;
  onCancel: () => void;
}) {
  const [text, setText] = useState("");
  const field = useRef<HTMLTextAreaElement>(null);
  // 不用 autoFocus：浏览器会为聚焦滚动对话列表，被当成用户上翻而停止跟随底部
  useEffect(() => field.current?.focus({ preventScroll: true }), []);
  const submit = () => {
    if (text.trim()) onSubmit(text.trim());
  };
  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.nativeEvent.isComposing || e.nativeEvent.keyCode === 229) return;
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    } else if (e.key === "Escape") {
      e.stopPropagation();
      onCancel();
    }
  };
  return (
    <div className="inline-note">
      <textarea
        ref={field}
        className="inline-note-input"
        aria-label={label}
        placeholder={placeholder}
        rows={2}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={onKeyDown}
      />
      <div className="inline-note-actions">
        <button className="btn primary tiny" disabled={!text.trim()} onClick={submit}>{submitLabel}</button>
        <button className="btn secondary tiny" onClick={onCancel}>返回</button>
      </div>
    </div>
  );
}

export interface ConfirmDecision {
  accepted: boolean;
  /** 拒绝时的说明 */
  note?: string;
  /** 允许，并在本次任务里自动允许同一工具的后续确认 */
  allowSimilar?: boolean;
}

export function ConfirmCard({
  item,
  onResolve,
}: {
  item: Extract<UiItem, { kind: "confirm" }>;
  onResolve: (id: string, decision: ConfirmDecision) => void;
}) {
  const [explaining, setExplaining] = useState(false);
  const [allowSimilar, setAllowSimilar] = useState(false);
  const [showDetail, setShowDetail] = useState(false);
  const detailId = useId();

  // 处理完的确认收成一行：预览已经没有可做的决定了，不该继续占半屏；想回看可以展开
  if (item.resolved) {
    const accepted = item.resolved === "accepted";
    return (
      <div className={`confirm-card settled ${item.resolved}${item.auto ? " auto" : ""}`}>
        <button
          type="button"
          className="confirm-auto-head disclosure-button"
          aria-expanded={showDetail}
          aria-controls={detailId}
          disabled={!item.detail}
          onClick={() => setShowDetail((v) => !v)}
        >
          <span className={`confirm-auto-icon${accepted ? "" : " rejected"}`}>
            {accepted ? <IconCheck size={12} /> : <IconClose size={11} />}
          </span>
          <span className="confirm-auto-title" title={item.title}>{item.title}</span>
          <span className="confirm-auto-tag">{item.auto ? "已自动允许" : accepted ? "已允许" : "已拒绝"}</span>
          {item.detail && <span className={`tool-expand${showDetail ? " open" : ""}`} aria-hidden="true"><IconChevronRight size={12} /></span>}
        </button>
        {!accepted && item.note && <div className="confirm-resolved">已拒绝：{item.note}</div>}
        <Collapse open={showDetail && !!item.detail} id={detailId}><pre className="confirm-detail">{item.detail}</pre></Collapse>
      </div>
    );
  }

  return (
    <div className="confirm-card">
      <div className="confirm-title">
        <span className="confirm-icon"><IconAlert size={15} /></span> {item.title}
      </div>
      {item.detail && <pre className="confirm-detail">{item.detail}</pre>}
      {explaining ? (
        <InlineNote
          label="拒绝原因或新的要求"
          placeholder="说明你想怎么做，例如：写到 F 列，不要覆盖原数据"
          submitLabel="拒绝并发送说明"
          onSubmit={(note) => onResolve(item.id, { accepted: false, note })}
          onCancel={() => setExplaining(false)}
        />
      ) : (
        <>
          <div className="confirm-actions">
            <button className="btn danger" onClick={() => onResolve(item.id, { accepted: true, allowSimilar })}>允许</button>
            <button className="btn secondary" onClick={() => onResolve(item.id, { accepted: false })}>拒绝</button>
            <button className="card-link" onClick={() => setExplaining(true)}>拒绝并说明…</button>
          </div>
          {item.tool && (
            <label className="confirm-similar">
              <input type="checkbox" checked={allowSimilar} onChange={(e) => setAllowSimilar(e.target.checked)} />
              本次任务中同类操作不再询问
            </label>
          )}
        </>
      )}
    </div>
  );
}

export type ChoiceDecision = { id: string; label: string } | { custom: string } | "skipped" | "cancelled";

export function ChoiceCard({
  item,
  onResolve,
}: {
  item: Extract<UiItem, { kind: "choice" }>;
  onResolve: (id: string, resolution: ChoiceDecision) => void;
}) {
  const [writing, setWriting] = useState(false);
  const resolved = item.resolved;
  return (
    <div className={`choice-card ${resolved ? "resolved" : ""}`}>
      <div className="choice-title">
        <span className="choice-icon"><IconCompass size={15} /></span> {item.title}
      </div>
      {item.question && <div className="choice-question">{item.question}</div>}
      <div className="choice-options">
        {item.options.map((opt) => {
          const selected = typeof resolved === "object" && resolved.id === opt.id;
          return (
            <button
              key={opt.id}
              className={`choice-option ${opt.recommended ? "recommended" : ""} ${selected ? "selected" : ""}`}
              disabled={!!resolved}
              onClick={() => onResolve(item.id, { id: opt.id, label: opt.label })}
            >
              <span className="choice-option-head">
                <span className="choice-option-label">{opt.label}</span>
                {opt.recommended && <span className="choice-badge">推荐</span>}
              </span>
              {opt.description && <span className="choice-option-desc">{opt.description}</span>}
            </button>
          );
        })}
      </div>
      {resolved ? (
        <div className="choice-resolved">
          {resolved === "skipped"
            ? "已交给 AI 自行判断"
            : resolved === "cancelled"
              ? "已取消"
              : resolved.custom
                ? `你的做法：「${resolved.label}」`
                : `已选择「${resolved.label}」`}
        </div>
      ) : writing ? (
        <InlineNote
          label="其他做法"
          placeholder="写下你想要的做法，例如：按客户汇总，只看 2024 年"
          submitLabel="按这个做"
          onSubmit={(text) => onResolve(item.id, { custom: text })}
          onCancel={() => setWriting(false)}
        />
      ) : (
        <div className="choice-footer">
          <button className="choice-skip" onClick={() => onResolve(item.id, "skipped")}>
            按你判断，不用我选
          </button>
          <button className="choice-skip" onClick={() => setWriting(true)}>
            其他做法…
          </button>
        </div>
      )}
    </div>
  );
}

const ERROR_ACTION_LABELS = { settings: "打开模型设置", "new-chat": "新建对话" } as const;

/**
 * 报错卡片：结论一句话，服务商原话作为次要信息；按钮给出下一步。
 * 需要先改设置或换对话才可能成功的错误，把那一步放在「重试」前面。
 */
export function ErrorBubble({
  text,
  detail,
  action,
  retryable,
  onRetry,
  onAction,
}: {
  text: string;
  detail?: string;
  action?: keyof typeof ERROR_ACTION_LABELS;
  retryable?: boolean;
  onRetry?: () => void;
  onAction?: (action: keyof typeof ERROR_ACTION_LABELS) => void;
}) {
  const showAction = !!action && !!onAction;
  const showRetry = !!retryable && !!onRetry;
  return (
    <div className="error-bubble" role="alert">
      <div className="error-text"><IconAlert size={15} className="error-icon" /><span>{text}</span></div>
      {detail && <div className="error-detail">{detail}</div>}
      {(showAction || showRetry) && (
        <div className="error-actions">
          {showAction && (
            <button className="error-retry primary" onClick={() => onAction!(action!)}>
              {ERROR_ACTION_LABELS[action!]}
            </button>
          )}
          {showRetry && (
            <button className="error-retry" onClick={onRetry}>
              重试这条指令
            </button>
          )}
        </div>
      )}
    </div>
  );
}

const SEV_LABELS: Record<AuditIssue["severity"], string> = { high: "高", medium: "中", low: "低" };
const TYPE_LABELS: Record<AuditIssue["type"], string> = {
  missing: "缺失值",
  duplicate: "重复行",
  type: "类型不一致",
  outlier: "离群值",
  whitespace: "首尾空格",
};

/** 数据体检结果卡片：勾选问题 → 让 AI 修复选中项 */
export function AuditCard({
  sheet,
  range,
  reportSheet,
  issues,
  dispatched,
  running,
  onFix,
}: {
  sheet: string;
  range: string;
  reportSheet?: string;
  issues: AuditIssue[];
  dispatched?: boolean;
  running?: boolean;
  onFix: (selected: AuditIssue[]) => void;
}) {
  const [checked, setChecked] = useState<Set<string>>(
    () => new Set(issues.filter((i) => i.severity === "high").map((i) => i.id))
  );
  useEffect(() => {
    if (dispatched) return;
    setChecked(new Set(issues.filter((i) => i.severity === "high").map((i) => i.id)));
  }, [issues, dispatched]);

  const counts = {
    high: issues.filter((i) => i.severity === "high").length,
    medium: issues.filter((i) => i.severity === "medium").length,
    low: issues.filter((i) => i.severity === "low").length,
  };
  const allChecked = checked.size === issues.length;
  const locked = dispatched || running;

  const toggle = (id: string) => {
    setChecked((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  if (issues.length === 0) {
    return (
      <div className="audit-card clean">
        <div className="audit-head">
          <span className="audit-icon"><IconPulse size={15} /></span> 数据体检
          <span className="audit-scope">{sheet}!{range}</span>
        </div>
        <div className="audit-done"><IconCheck size={13} /> 未发现问题，数据很干净</div>
      </div>
    );
  }

  return (
    <div className={`audit-card ${dispatched ? "dispatched" : ""}`}>
      <div className="audit-head">
        <span className="audit-icon"><IconPulse size={15} /></span> 数据体检 · {issues.length} 个问题
        <span className="audit-sev-summary">
          高 {counts.high} / 中 {counts.medium} / 低 {counts.low}
        </span>
        <span className="audit-scope" title="检查范围">
          {sheet}!{range}
        </span>
      </div>
      {reportSheet && <div className="audit-report-hint">完整报告已写入 Sheet「{reportSheet}」</div>}
      <div className="audit-issues">
        {issues.map((issue) => (
          <label key={issue.id} className={`audit-issue ${checked.has(issue.id) ? "checked" : ""}`}>
            <input
              type="checkbox"
              checked={checked.has(issue.id)}
              disabled={locked}
              onChange={() => toggle(issue.id)}
            />
            <span className={`audit-sev-dot ${issue.severity}`} title={`严重度：${SEV_LABELS[issue.severity]}`} />
            <span className="audit-issue-main">
              <span className="audit-issue-title">
                {TYPE_LABELS[issue.type]}
                {issue.column ? ` · ${issue.column}` : ""} × {issue.count}
              </span>
              <span className="audit-issue-sample" title={issue.suggestion}>{issue.sample}</span>
            </span>
          </label>
        ))}
      </div>
      {dispatched ? (
        <div className="audit-done"><IconCheck size={13} /> 已发起修复，进展见上方对话</div>
      ) : (
        <div className="audit-actions">
          <button
            className="btn tiny secondary"
            disabled={running}
            onClick={() => setChecked(allChecked ? new Set() : new Set(issues.map((i) => i.id)))}
          >
            {allChecked ? "全不选" : "全选"}
          </button>
          <button
            className="btn primary tiny"
            disabled={running || checked.size === 0}
            onClick={() => onFix(issues.filter((i) => checked.has(i.id)))}
          >
            修复选中项{checked.size > 0 ? ` (${checked.size})` : ""}
          </button>
        </div>
      )}
    </div>
  );
}

export function NoticeBubble({ text }: { text: string }) {
  // 旧对话里的提示以 emoji 开头，渲染时换成统一的线性图标
  const { Icon = IconInfo, rest } = splitLeadingGlyph(text);
  return <div className="notice-bubble"><Icon size={13} className="notice-icon" /><span>{rest}</span></div>;
}

/** 模型思考中（还没有流式输出、也没有工具在跑）时的指示器：流光文字，等久了附上已用时间 */
export function ThinkingIndicator({ label, since }: { label?: string; since?: number }) {
  const elapsed = useElapsedSeconds(since);
  return (
    <div className="msg-row assistant thinking-row">
      <div className="thinking-bubble">
        <span className="thinking-spark"><IconSparkle size={13} /></span>
        <span className="shimmer-text">{label || "思考中…"}</span>
        {elapsed >= 5 && <span className="thinking-elapsed">{formatElapsed(elapsed)}</span>}
      </div>
    </div>
  );
}

export function FollowupBar({
  actions,
  disabled,
  onPick,
}: {
  actions: FollowupAction[];
  disabled?: boolean;
  onPick: (action: FollowupAction) => void;
}) {
  return (
    <div className="followup-bar">
      <div className="followup-label">接下来可以</div>
      <div className="followup-actions">
        {actions.map((a) => (
          <button
            key={a.label}
            className={`followup-chip${a.run ? " run" : ""}`}
            disabled={disabled}
            title={a.run ? `直接执行：${a.prompt}` : `填入输入框，可修改后发送：${a.prompt}`}
            onClick={() => onPick(a)}
          >
            {a.run && <IconPlay size={10} />}
            {a.label}
          </button>
        ))}
      </div>
    </div>
  );
}

/**
 * 一段已结束的执行过程（工具调用 + 过程叙述）折叠成一行，结论留在外面。
 * 折叠时仍把失败的步骤露出来，出了问题不用展开找。
 */
export function ProcessGroup({
  steps,
  failed,
  failedContent,
  children,
}: {
  steps: number;
  failed: number;
  failedContent: ReactNode;
  children: ReactNode;
}) {
  const [open, setOpen] = useState(false);
  const bodyId = useId();
  return (
    <div className={`process-group${failed ? " has-failed" : ""}`}>
      <button
        type="button"
        className="process-group-head disclosure-button"
        aria-expanded={open}
        aria-controls={bodyId}
        onClick={() => setOpen((v) => !v)}
      >
        <span className="process-group-icon">{failed ? <IconAlert size={12} /> : <IconCheck size={12} />}</span>
        <span className="process-group-text">
          已执行 {steps} 步{failed ? ` · ${failed} 步未成功` : ""}
        </span>
        <span className={`tool-expand${open ? " open" : ""}`} aria-hidden="true"><IconChevronRight size={13} /></span>
      </button>
      {/* 展开时失败步骤随正文一起过渡收起，而不是瞬间消失：总高度连续变化，不会把滚动位置夹到顶部 */}
      <Collapse open={!open && failed > 0} appear={false}>
        <div className="process-group-failed">{failedContent}</div>
      </Collapse>
      <Collapse open={open} id={bodyId}>
        <div className="process-group-body">{children}</div>
      </Collapse>
    </div>
  );
}
