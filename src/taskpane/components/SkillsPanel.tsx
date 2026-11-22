import { useEffect, useState } from "react";
import {
  Skill,
  approveSkill,
  clearSkills,
  loadSkills,
  onSkillsChange,
  rejectSkill,
  removeSkill,
  upsertSkill,
} from "../../core/skills";
import Collapse from "./Collapse";
import MotionList from "./MotionList";
import { EmptyState, PanelHeader } from "./Panel";
import { IconSearch, IconTrash, IconZap } from "./Icons";

interface Props {
  onClose: () => void;
  onUse: (skill: Skill) => void;
}

export default function SkillsPanel({ onClose, onUse }: Props) {
  const [items, setItems] = useState<Skill[]>(loadSkills);
  const [expanded, setExpanded] = useState<string | null>(null);
  const [query, setQuery] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [draftName, setDraftName] = useState("");
  const [draftSteps, setDraftSteps] = useState("");
  const [draftTriggers, setDraftTriggers] = useState("");
  const [error, setError] = useState("");
  // 面板内二次确认，不用 window.confirm（任务窗格里原生弹窗可能被宿主吞掉）
  const [confirmClear, setConfirmClear] = useState(false);

  useEffect(() => {
    setItems(loadSkills());
    return onSkillsChange(() => setItems(loadSkills()));
  }, []);

  const add = () => {
    setError("");
    try {
      upsertSkill({
        name: draftName,
        steps: draftSteps,
        triggers: draftTriggers.split(/[,，、\s]+/).filter(Boolean),
        source: "manual",
      });
      setDraftName("");
      setDraftSteps("");
      setDraftTriggers("");
      setShowAdd(false);
    } catch (e: any) {
      setError(e?.message || String(e));
    }
  };

  const needsReview = (s: Skill) => !!(s.pending || s.proposal);
  const visibleItems = items
    .filter((item) => {
      const q = query.trim().toLowerCase();
      if (!q) return true;
      return `${item.name} ${item.description} ${item.triggers.join(" ")} ${item.steps}`.toLowerCase().includes(q);
    })
    // 待确认的排在前面
    .sort((a, b) => Number(needsReview(b)) - Number(needsReview(a)));

  return (
    <div className="history-overlay" onClick={onClose}>
      <div className="history-panel memory-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconZap size={16} />} title="技能库" onClose={onClose} />

        <p className="memory-hint">
          完成任务后 AI 会自动复盘，把可复用的流程总结成技能；你确认后，下次遇到相似任务会自动调取。升级插件后技能库会自动继承，也可以手动添加或删除。
        </p>

        <div className="skill-toolbar">
          <button className="btn secondary" aria-expanded={showAdd} aria-controls="skill-editor" onClick={() => setShowAdd((v) => !v)}>
            {showAdd ? "收起" : "手动添加"}
          </button>
          {items.length > 0 && (confirmClear ? (
            <div className="history-item-confirm">
              <button
                className="btn tiny danger"
                onClick={() => {
                  clearSkills();
                  setConfirmClear(false);
                }}
              >
                确定清空全部技能
              </button>
              <button className="btn tiny" onClick={() => setConfirmClear(false)}>取消</button>
            </div>
          ) : (
            <button className="memory-clear" onClick={() => setConfirmClear(true)}>
              清空技能库
            </button>
          ))}
        </div>

        <Collapse open={showAdd} id="skill-editor">
          <div className="memory-add">
            <input
              className="skill-add-input"
              value={draftName}
              placeholder="技能名，如：跨表订单合并"
              onChange={(e) => setDraftName(e.target.value)}
            />
            <input
              className="skill-add-input"
              value={draftTriggers}
              placeholder="触发关键词（逗号分隔），如：合并, 订单, 关联"
              onChange={(e) => setDraftTriggers(e.target.value)}
            />
            <textarea
              value={draftSteps}
              rows={4}
              placeholder={"步骤，如：\n1. 查看两表结构，确认关联键\n2. 用 XLOOKUP 把 B 表字段并入 A 表副本…"}
              onChange={(e) => setDraftSteps(e.target.value)}
            />
            <div className="memory-add-row">
              <span />
              <button className="btn primary" disabled={!draftName.trim() || !draftSteps.trim()} onClick={add}>
                保存技能
              </button>
            </div>
            {error && <div className="memory-error">{error}</div>}
          </div>
        </Collapse>

        <div className="history-search skill-search">
          <IconSearch size={14} />
          <input value={query} placeholder="搜索技能名、触发词或步骤…" onChange={(e) => setQuery(e.target.value)} />
        </div>

        <div className="history-list">
          {items.length === 0 && (
            <EmptyState icon={<IconZap size={20} />} title="技能库还是空的" hint="完成几个多步任务后，AI 会自动学会第一批技能；也可以点「手动添加」。" />
          )}
          {items.length > 0 && visibleItems.length === 0 && <EmptyState icon={<IconSearch size={20} />} title="没有匹配的技能" hint="换个关键词试试" />}
          <MotionList>{visibleItems.map((s) => (
            <div key={s.id} className="skill-item">
              <div className="skill-item-head">
                <button type="button" className="skill-item-main disclosure-button" aria-expanded={expanded === s.id} aria-controls={`skill-${s.id}`} onClick={() => setExpanded(expanded === s.id ? null : s.id)}>
                  <div className="skill-item-name">
                    {s.name}
                    <span className={`skill-badge ${s.source}`}>{s.source === "learned" ? "自学" : "手动"}</span>
                    {s.pending && <span className="skill-badge review">待确认</span>}
                    {!s.pending && s.proposal && <span className="skill-badge review">有修改建议</span>}
                    {s.useCount > 0 && <span className="skill-use">调用 {s.useCount} 次</span>}
                  </div>
                  {s.description && <div className="skill-item-desc">{s.description}</div>}
                  {s.triggers.length > 0 && (
                    <div className="skill-item-triggers">
                      {s.triggers.map((t) => (
                        <span key={t} className="skill-trigger">{t}</span>
                      ))}
                    </div>
                  )}
                </button>
                <button
                  className="history-item-del"
                  title="删除技能"
                  onClick={(e) => {
                    e.stopPropagation();
                    removeSkill(s.id);
                  }}
                >
                  <IconTrash size={14} />
                </button>
              </div>
              <Collapse open={expanded === s.id} id={`skill-${s.id}`}>
                <div className="skill-expanded">
                  {needsReview(s) && (
                    <div className="skill-review">
                      <span>
                        {s.pending
                          ? "这是 AI 从任务里总结的新技能，确认前不会被自动使用。请检查步骤里没有你不认可的操作。"
                          : "AI 建议把步骤改成下面的内容；确认前仍使用原步骤。"}
                      </span>
                      {s.proposal && <pre className="skill-steps">{s.proposal.steps}</pre>}
                      <div className="skill-review-actions">
                        <button className="btn tiny primary" onClick={() => approveSkill(s.id)}>
                          {s.pending ? "采用" : "采用修改"}
                        </button>
                        <button className="btn tiny" onClick={() => rejectSkill(s.id)}>
                          {s.pending ? "不要" : "保留原样"}
                        </button>
                      </div>
                    </div>
                  )}
                  {s.proposal && <div className="skill-item-desc">当前步骤：</div>}
                  <pre className="skill-steps">{s.steps}</pre>
                  <button className="btn primary skill-use-btn" onClick={() => onUse(s)}>
                    套用到当前任务
                  </button>
                </div>
              </Collapse>
            </div>
          ))}</MotionList>
        </div>
      </div>
    </div>
  );
}
