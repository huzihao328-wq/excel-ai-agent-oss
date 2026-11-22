import { useEffect, useMemo, useState } from "react";
import Collapse from "./Collapse";
import MotionList from "./MotionList";
import { EmptyState, PanelHeader } from "./Panel";
import { GlyphIcon, IconPlus, IconSearch, IconTemplate } from "./Icons";
import {
  listPromptTemplates,
  onPromptTemplatesChange,
  PromptTemplate,
  removePromptTemplate,
  upsertPromptTemplate,
} from "../../core/promptTemplates";

interface Props {
  currentPrompt: string;
  busy: boolean;
  onUse: (prompt: string) => void;
  onRun: (prompt: string) => void;
  onClose: () => void;
}

export default function PromptTemplatesPanel({ currentPrompt, busy, onUse, onRun, onClose }: Props) {
  const [items, setItems] = useState<PromptTemplate[]>(listPromptTemplates);
  const [query, setQuery] = useState("");
  const [showEditor, setShowEditor] = useState(false);
  const [editId, setEditId] = useState<string | undefined>();
  const [draftName, setDraftName] = useState("");
  const [draftPrompt, setDraftPrompt] = useState("");
  const [error, setError] = useState("");
  const [pendingDeleteId, setPendingDeleteId] = useState<string | null>(null);

  useEffect(() => onPromptTemplatesChange(() => setItems(listPromptTemplates())), []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return items;
    return items.filter((item) => `${item.name} ${item.category} ${item.prompt}`.toLowerCase().includes(q));
  }, [items, query]);

  const groups = useMemo(() => {
    const categories = ["我的", "分析", "汇总", "核对", "清洗", "整理"];
    return categories
      .map((category) => ({ category, items: filtered.filter((item) => item.category === category) }))
      .filter((group) => group.items.length > 0);
  }, [filtered]);

  const openCreate = () => {
    setEditId(undefined);
    setDraftName("");
    setDraftPrompt(currentPrompt.trim());
    setError("");
    setShowEditor(true);
  };

  const openEdit = (item: PromptTemplate) => {
    setEditId(item.id);
    setDraftName(item.name);
    setDraftPrompt(item.prompt);
    setError("");
    setShowEditor(true);
  };

  const save = () => {
    setError("");
    try {
      upsertPromptTemplate({ id: editId, name: draftName, prompt: draftPrompt });
      setShowEditor(false);
      setEditId(undefined);
      setDraftName("");
      setDraftPrompt("");
    } catch (e: any) {
      setError(e?.message || String(e));
    }
  };

  return (
    <div className="history-overlay" onClick={onClose}>
      <div className="history-panel template-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconTemplate size={16} />} title="常用指令" onClose={onClose} />

        <p className="memory-hint">保存经常重复的指令，填入后可补充细节，也可以直接运行。</p>

        <button className="history-new-btn" onClick={openCreate}>
          <IconPlus size={14} /> {currentPrompt.trim() ? "保存当前输入" : "新建常用指令"}
        </button>

        <Collapse open={showEditor}>
          <div className="template-editor">
            <input
              value={draftName}
              placeholder="名称，如：每周销售周报"
              onChange={(e) => setDraftName(e.target.value)}
            />
            <textarea
              value={draftPrompt}
              rows={6}
              placeholder="写下希望 AI 重复执行的完整指令…"
              onChange={(e) => setDraftPrompt(e.target.value)}
            />
            <div className="template-editor-actions">
              <button className="btn tiny" onClick={() => setShowEditor(false)}>取消</button>
              <button className="btn primary" disabled={!draftName.trim() || !draftPrompt.trim()} onClick={save}>
                保存
              </button>
            </div>
            {error && <div className="memory-error">{error}</div>}
          </div>
        </Collapse>

        <div className="history-search">
          <IconSearch size={14} />
          <input value={query} placeholder="搜索名称、场景或指令内容…" onChange={(e) => setQuery(e.target.value)} />
        </div>

        <div className="history-list template-list">
          <MotionList>
          {groups.length === 0 && <div key="empty"><EmptyState icon={<IconSearch size={20} />} title="没有匹配的常用指令" hint="换个关键词，或新建一条" /></div>}
          {groups.map((group) => (
            <section key={group.category} className="template-group">
              <div className="history-group-label">{group.category}</div>
              <MotionList>{group.items.map((item) => (
                <article key={item.id} className="template-card">
                  <div className="template-card-head">
                    <span className="template-card-icon"><GlyphIcon glyph={item.icon} size={15} /></span>
                    <div className="template-card-title">{item.name}</div>
                    {!item.builtIn && (
                      <div className="template-card-manage">
                        {pendingDeleteId === item.id ? (
                          <>
                            <button
                              className="danger-text"
                              onClick={() => {
                                removePromptTemplate(item.id);
                                setPendingDeleteId(null);
                              }}
                            >
                              确定删除
                            </button>
                            <button onClick={() => setPendingDeleteId(null)}>取消</button>
                          </>
                        ) : (
                          <>
                            <button onClick={() => openEdit(item)}>编辑</button>
                            <button className="danger-text" onClick={() => setPendingDeleteId(item.id)}>删除</button>
                          </>
                        )}
                      </div>
                    )}
                  </div>
                  <p>{item.prompt}</p>
                  <div className="template-card-actions">
                    <button className="btn tiny" onClick={() => onUse(item.prompt)}>填入</button>
                    <button className="btn tiny template-run" disabled={busy} onClick={() => onRun(item.prompt)}>
                      直接运行
                    </button>
                  </div>
                </article>
              ))}</MotionList>
            </section>
          ))}
          </MotionList>
        </div>
      </div>
    </div>
  );
}
