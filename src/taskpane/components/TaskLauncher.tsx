import { useEffect, useRef, useState } from "react";
import { QUICK_ACTIONS } from "../quickActions";
import Presence from "./Presence";
import MotionTabs from "./MotionTabs";
import MotionList from "./MotionList";
import { MOTION } from "../motion";
import { GlyphIcon, IconArrowUpRight, IconChevronDown, IconClose, IconPlay, IconSearch, IconTemplate } from "./Icons";

const groups = [
  { label: "全部", names: [] },
  { label: "分析", names: ["分析当前表", "数据体检", "列统计", "生成图表", "做透视表"] },
  { label: "整理", names: ["数据清洗", "跨表合并", "整理格式", "冻结表头"] },
  { label: "公式与导出", names: ["写公式", "导出文件"] },
];

/**
 * 输入框上方的快捷操作。规则：只读的（解释、诊断）点了直接执行；
 * 可能改表的一律只填入输入框，用户看过、补充后再发。两种样式不同，点之前就能分辨。
 */
export interface ContextAction {
  label: string;
  prompt: string;
  mode: "run" | "fill";
  tone?: "danger";
  /** 来源标签，如「记忆」「技能」 */
  tag?: string;
}

function ShortcutChip({ action, disabled, onRun, onPick }: {
  action: ContextAction;
  disabled: boolean;
  onRun: (prompt: string) => void;
  onPick: (prompt: string) => void;
}) {
  const run = action.mode === "run";
  return (
    <button
      className={`chip context ${run ? "run" : "fill"}${action.tone === "danger" ? " danger" : ""}`}
      disabled={disabled}
      title={run ? `直接执行（只读，不改表）：${action.prompt}` : `填入输入框，可修改后发送：${action.prompt}`}
      onClick={() => (run ? onRun(action.prompt) : onPick(action.prompt))}
    >
      {run && <IconPlay size={10} />}
      {action.label}
      {action.tag && <span className="chip-tag">{action.tag}</span>}
    </button>
  );
}

interface Props {
  disabled: boolean;
  /** 随当前选区变化的情境操作 */
  contextActions?: ContextAction[];
  onRun: (prompt: string) => void;
  /** 快捷任务只填入输入框，用户补充要求后再发 */
  onPick: (prompt: string) => void;
  onTemplates?: () => void;
}

export default function TaskLauncher({ disabled, contextActions = [], onRun, onPick, onTemplates }: Props) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [group, setGroup] = useState(0);
  const root = useRef<HTMLDivElement>(null);
  const toggle = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    if (!open) return;
    const dismiss = (event: PointerEvent) => {
      if (!root.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("pointerdown", dismiss);
    return () => document.removeEventListener("pointerdown", dismiss);
  }, [open]);
  const filtered = QUICK_ACTIONS.filter((action) =>
    (group === 0 || groups[group].names.includes(action.label)) &&
    `${action.label} ${action.prompt}`.toLowerCase().includes(query.trim().toLowerCase())
  );
  const pick = (prompt: string) => { setOpen(false); onPick(prompt); };
  const close = () => { setOpen(false); toggle.current?.focus(); };
  return (
    <div className="task-launcher" ref={root} onKeyDown={(event) => {
      if (event.key === "Escape" && open) { event.stopPropagation(); close(); }
    }}>
      <div className="launcher-shortcuts">
        <div className="launcher-scroll" role="group" aria-label={contextActions.length ? "针对当前选区" : "常用快捷任务"}>
          {contextActions.length > 0
            ? contextActions.map((action) => (
              <ShortcutChip key={action.label} action={action} disabled={disabled} onRun={onRun} onPick={pick} />
            ))
            : QUICK_ACTIONS.slice(0, 2).map((action) => (
              <button className="chip fill" key={action.label} disabled={disabled} title={`填入输入框，可修改后发送：${action.prompt}`} onClick={() => pick(action.prompt)}>{action.label}</button>
            ))}
        </div>
        <button className={`chip launcher-toggle${open ? " active" : ""}`} ref={toggle} aria-expanded={open} aria-controls="task-catalog" onClick={() => setOpen(!open)}>
          全部任务 <span className={`launcher-caret${open ? " open" : ""}`} aria-hidden="true"><IconChevronDown size={12} /></span>
        </button>
      </div>
      <Presence open={open} duration={MOTION.duration.exit} focus="popover">
        <section className="task-catalog" id="task-catalog" aria-label="快捷任务">
          <div className="task-catalog-head">
            <strong>你想做什么？</strong>
            <button className="icon-btn" aria-label="关闭快捷任务" onClick={close}><IconClose size={14} /></button>
          </div>
          <label className="search-field">
            <IconSearch size={14} />
            <input autoFocus aria-label="搜索快捷任务" placeholder="搜索任务，例如：公式、合并…" value={query} onChange={(event) => setQuery(event.target.value)} />
          </label>
          <MotionTabs className="task-categories" label="任务分类" options={groups.map((item, index) => ({ value: String(index), label: item.label }))} value={String(group)} onChange={value => setGroup(Number(value))} />
          <MotionList className="task-results">{filtered.map((action) => (
            <button key={action.label} disabled={disabled} onClick={() => pick(action.prompt)}>
              <span className="task-result-icon" aria-hidden="true"><GlyphIcon glyph={action.icon} size={15} /></span>
              <span>{action.label}</span>
              <span className="task-result-arrow" aria-hidden="true"><IconArrowUpRight size={12} /></span>
            </button>
          ))}</MotionList>
          {filtered.length === 0 && <p className="task-no-results">没有匹配的任务，试试其他关键词。</p>}
          {onTemplates && (
            <button className="task-catalog-link" onClick={() => { setOpen(false); onTemplates(); }}>
              <IconTemplate size={14} /> 我的常用指令
            </button>
          )}
        </section>
      </Presence>
    </div>
  );
}
