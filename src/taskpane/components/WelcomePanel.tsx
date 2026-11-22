import { SuggestedPrompt } from "../../core/recommendations";
import { IconArrowRight, IconArrowUpRight, IconBulb, IconSheet, IconSparkle, IconZap } from "./Icons";

interface Props {
  needsSetup: boolean;
  onSetup: () => void;
  suggestions: SuggestedPrompt[];
  contextLabel: string;
  disabled: boolean;
  onPick: (prompt: string) => void;
  onTemplates: () => void;
}

export default function WelcomePanel({ needsSetup, onSetup, suggestions, contextLabel, disabled, onPick, onTemplates }: Props) {
  return (
    <section className="welcome-panel" aria-label="任务工作台">
      <div className="welcome-mark" aria-hidden="true"><IconSparkle size={22} /></div>
      <div className="welcome-eyebrow">你的 Excel 工作台</div>
      <h1>把表格工作，<br />交给一句话。</h1>
      <p className="welcome-description">分析数据、整理表格、编写公式。选中单元格，再告诉我你想完成什么。</p>
      {needsSetup && (
        <button className="welcome-setup" onClick={onSetup}>
          <strong>连接你的模型 <IconArrowRight size={14} /></strong>
          <span>配置 API 后，即可开始处理表格</span>
        </button>
      )}
      <div className="welcome-section-head"><h2>{contextLabel}</h2><button onClick={onTemplates}>常用指令 <IconArrowUpRight size={12} /></button></div>
      <div className="welcome-tasks">
        {suggestions.slice(0, 4).map((item) => {
          const Icon = item.source === "selection" ? IconSheet : item.source === "skill" ? IconZap : IconBulb;
          return (
            <button key={item.id} disabled={disabled} onClick={() => onPick(item.prompt)} title={item.prompt}>
              <span className="welcome-task-icon" aria-hidden="true"><Icon size={15} /></span>
              <span className="welcome-task-copy"><strong>{item.label}</strong><small>{item.source === "selection" ? "基于当前表格" : "结合你的记忆与技能"}</small></span>
              <span className="welcome-task-arrow" aria-hidden="true"><IconArrowUpRight size={13} /></span>
            </button>
          );
        })}
      </div>
      <p className="welcome-tip">点击任务填入指令，补充要求后发送。也可拖入 Excel、CSV 或截图。</p>
    </section>
  );
}
