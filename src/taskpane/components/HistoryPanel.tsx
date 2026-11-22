import { memo, useMemo, useState } from "react";
import { ConversationMeta } from "../../core/conversations";
import MotionList from "./MotionList";
import { EmptyState, PanelHeader } from "./Panel";
import { IconChats, IconPlus, IconSearch, IconTrash } from "./Icons";

interface Props {
  conversations: ConversationMeta[];
  currentId: string | null;
  runningId?: string | null;
  onSelect: (id: string) => void;
  onDelete: (id: string) => void;
  onNew: () => void;
  onClose: () => void;
}

function formatRelativeTime(ts: number): string {
  const diff = Date.now() - ts;
  if (diff < 60_000) return "刚刚";
  if (diff < 3_600_000) return `${Math.floor(diff / 60_000)} 分钟前`;
  if (diff < 86_400_000) return `${Math.floor(diff / 3_600_000)} 小时前`;
  const d = new Date(ts);
  const now = new Date();
  if (d.getFullYear() === now.getFullYear()) return `${d.getMonth() + 1}月${d.getDate()}日`;
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

function groupConversations(items: ConversationMeta[]): { label: string; items: ConversationMeta[] }[] {
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const startOfYesterday = startOfToday - 86_400_000;
  const startOfWeek = startOfToday - 6 * 86_400_000;

  const groups: { label: string; items: ConversationMeta[] }[] = [];
  const today = items.filter((i) => i.updatedAt >= startOfToday);
  const yesterday = items.filter((i) => i.updatedAt >= startOfYesterday && i.updatedAt < startOfToday);
  const week = items.filter((i) => i.updatedAt >= startOfWeek && i.updatedAt < startOfYesterday);
  const older = items.filter((i) => i.updatedAt < startOfWeek);
  if (today.length) groups.push({ label: "今天", items: today });
  if (yesterday.length) groups.push({ label: "昨天", items: yesterday });
  if (week.length) groups.push({ label: "近 7 天", items: week });
  if (older.length) groups.push({ label: "更早", items: older });
  return groups;
}

export default function HistoryPanel({
  conversations,
  currentId,
  runningId,
  onSelect,
  onDelete,
  onNew,
  onClose,
}: Props) {
  const [query, setQuery] = useState("");
  const [pendingDeleteId, setPendingDeleteId] = useState<string | null>(null);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return conversations;
    return conversations.filter((c) => c.title.toLowerCase().includes(q));
  }, [conversations, query]);

  const groups = useMemo(() => groupConversations(filtered), [filtered]);

  return (
    <div className="history-overlay" onClick={onClose}>
      <div className="history-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconChats size={16} />} title="历史对话" onClose={onClose} />

        <button className="history-new-btn" onClick={onNew}>
          <IconPlus size={14} /> 新对话
        </button>

        <div className="history-search">
          <IconSearch size={14} />
          <input
            value={query}
            placeholder="搜索对话…"
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>

        <div className="history-list">
          <MotionList>
          {groups.length === 0 && (
            <div key="empty">
              {query.trim()
                ? <EmptyState icon={<IconSearch size={20} />} title="没有匹配的对话" hint="换个关键词试试" />
                : <EmptyState icon={<IconChats size={20} />} title="还没有历史对话" hint="发出第一条指令后，对话会自动保存在这里。" />}
            </div>
          )}
          {groups.map((group) => (
            <div key={group.label} className="history-group">
              <div className="history-group-label">{group.label}</div>
              <MotionList>{group.items.map((c) => (
                <HistoryRow key={c.id} conversation={c} current={c.id === currentId} running={c.id === runningId}
                  confirming={pendingDeleteId === c.id} onSelect={onSelect} onDelete={onDelete} onConfirm={setPendingDeleteId} />
              ))}</MotionList>
            </div>
          ))}
          </MotionList>
        </div>
      </div>
    </div>
  );
}

const HistoryRow = memo(function HistoryRow({ conversation: c, current, running, confirming, onSelect, onDelete, onConfirm }: {
  conversation: ConversationMeta; current: boolean; running: boolean; confirming: boolean;
  onSelect: (id: string) => void; onDelete: (id: string) => void; onConfirm: (id: string | null) => void;
}) {
  return <div className={`history-item ${current ? "active" : ""} ${running ? "running" : ""}`}>
    <button className="history-item-main" onClick={() => onSelect(c.id)} title={c.title}>
      <span className="history-item-title">
        {running && <span className="history-running-dot" title="正在执行" />}
        <span className="history-item-title-text">{c.title}</span>
      </span>
      <span className="history-item-time">{running ? "执行中" : formatRelativeTime(c.updatedAt)}</span>
    </button>
    {confirming ? <div className="history-item-confirm">
      <button className="btn tiny danger" onClick={() => onDelete(c.id)}>删除</button>
      <button className="btn tiny" onClick={() => onConfirm(null)}>取消</button>
    </div> : <button className="history-item-del" title="删除对话" onClick={() => onConfirm(c.id)}><IconTrash size={14} /></button>}
  </div>;
});
