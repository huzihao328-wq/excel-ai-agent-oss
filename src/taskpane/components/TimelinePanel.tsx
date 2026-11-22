import { useEffect, useState } from "react";
import Collapse from "./Collapse";
import MotionList from "./MotionList";
import { EmptyState, PanelHeader } from "./Panel";
import { IconChevronRight, IconHistory, IconUndo } from "./Icons";
import { OpLogEntry, UndoEntry } from "../../core/types";
import { OP_KIND_LABELS, getOpLog, onOpLogChange } from "../../core/oplog";
import { getUndoStack, onUndoStackChange } from "../../core/tools/excelTools";

interface Props {
  onClose: () => void;
  onUndoEntry: (entry: UndoEntry) => void;
  undoBusy?: boolean;
}

function formatTime(ts: number): string {
  const d = new Date(ts);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

export default function TimelinePanel({ onClose, onUndoEntry, undoBusy }: Props) {
  const [entries, setEntries] = useState<OpLogEntry[]>(() => [...getOpLog()].reverse());
  const [undoIds, setUndoIds] = useState<Set<string>>(() => new Set(getUndoStack().map((e) => e.id)));
  const [expanded, setExpanded] = useState<string | null>(null);

  useEffect(() => {
    const offLog = onOpLogChange(() => setEntries([...getOpLog()].reverse()));
    const offUndo = onUndoStackChange(() => setUndoIds(new Set(getUndoStack().map((e) => e.id))));
    return () => {
      offLog();
      offUndo();
    };
  }, []);

  return (
    <div className="history-overlay" onClick={onClose}>
      <div className="history-panel timeline-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconHistory size={16} />} title="操作记录" onClose={onClose} />

        <p className="memory-hint">本次会话 AI 对工作簿做过的改动，按时间倒序。点击条目可查看修改前后的对比。</p>

        <div className="history-list">
          <MotionList>
          {entries.length === 0 && <div key="empty"><EmptyState icon={<IconHistory size={20} />} title="还没有改动记录" hint="AI 修改工作簿后，每一步都会记在这里，可以查看前后对比并逐条撤回。" /></div>}
          {entries.map((entry) => {
            const open = expanded === entry.id;
            const canUndo = !!entry.undoEntryId && undoIds.has(entry.undoEntryId);
            return (
              <div key={entry.id} className={`tl-item ${open ? "open" : ""}`}>
                <button type="button"
                  className="tl-head disclosure-button" aria-expanded={open} aria-controls={`log-${entry.id}`}
                  onClick={() => setExpanded(open ? null : entry.id)}
                  title={(entry.before || entry.after) ? "点击查看前后对比" : undefined}
                >
                  <span className="tl-time">{formatTime(entry.timestamp)}</span>
                  <span className={`tl-badge ${entry.kind}`}>{OP_KIND_LABELS[entry.kind]}</span>
                  <span className="tl-desc">{entry.description}</span>
                  {(entry.before || entry.after) && <span className={`tl-expand${open ? " open" : ""}`} aria-hidden="true"><IconChevronRight size={12} /></span>}
                </button>
                <Collapse open={open} id={`log-${entry.id}`}>
                {(entry.before || entry.after) && (
                  <div className="tl-diff">
                    {entry.before && (
                      <div className="tl-diff-col">
                        <div className="tl-diff-label">修改前</div>
                        <pre>{entry.before}</pre>
                      </div>
                    )}
                    {entry.after && (
                      <div className="tl-diff-col">
                        <div className="tl-diff-label">修改后</div>
                        <pre>{entry.after}</pre>
                      </div>
                    )}
                  </div>
                )}
                {canUndo && entry.undoEntryId && (
                  <div className="tl-actions">
                    <button
                      className="btn tiny"
                      disabled={undoBusy}
                      onClick={() => {
                        const ue = getUndoStack().find((e) => e.id === entry.undoEntryId);
                        if (ue) onUndoEntry(ue);
                      }}
                    >
                        <IconUndo size={12} /> 撤回到此节点
                    </button>
                  </div>
                )}
                </Collapse>
              </div>
            );
          })}
          </MotionList>
        </div>
      </div>
    </div>
  );
}
