import { useState } from "react";
import { UndoEntry } from "../../core/types";
import { TaskCheckpoint } from "../../core/tools/checkpoint";
import Collapse from "./Collapse";
import { IconChevronRight, IconUndo } from "./Icons";

function formatCheckpointTime(ts: number): string {
  const d = new Date(ts);
  return `${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")}`;
}

/**
 * 唯一的回退入口：平时只露出「撤销上一步」，展开后可逐步撤回、整段恢复到任务开始，
 * 或进入操作记录看修改前后对比。取代原来分开的撤销栏和检查点栏。
 */
export default function RollbackBar({
  undoEntries,
  checkpoint,
  cpBusy,
  disabled,
  undoBusy,
  onUndo,
  onCheckpointAction,
  onOpenTimeline,
}: {
  undoEntries: UndoEntry[];
  checkpoint: TaskCheckpoint | null;
  cpBusy: boolean;
  disabled: boolean;
  undoBusy: boolean;
  onUndo: (entry: UndoEntry) => void;
  onCheckpointAction: (action: "restore" | "discard") => void;
  onOpenTimeline: () => void;
}) {
  const [open, setOpen] = useState(false);
  if (!undoEntries.length && !checkpoint && !cpBusy) return null;
  const latest = undoEntries[undoEntries.length - 1];
  const checkpointMeta = checkpoint
    ? `${checkpoint.sheets.length} 个 Sheet${checkpoint.label ? ` · ${checkpoint.label}` : ""} · ${formatCheckpointTime(checkpoint.createdAt)}`
    : "";

  return (
    <div className="rollback-bar">
      <div className="rollback-main">
        {latest ? (
          <>
            <button
              className="btn tiny rollback-undo"
              disabled={disabled || undoBusy}
              title={`撤销：${latest.description}`}
              onClick={() => onUndo(latest)}
            >
              <IconUndo size={12} /> {undoBusy ? "处理中…" : "撤销上一步"}
            </button>
            <span className="rollback-desc" title={latest.description}>{latest.description}</span>
          </>
        ) : (
          <span className="rollback-desc" title={checkpoint?.label}>
            {cpBusy ? "正在处理检查点…" : `可恢复到任务开始 · ${checkpointMeta}`}
          </span>
        )}
        <button
          className="rollback-toggle"
          aria-expanded={open}
          aria-controls="rollback-details"
          onClick={() => setOpen((v) => !v)}
        >
          回退选项
          <span className={`undo-caret${open ? " open" : ""}`} aria-hidden="true"><IconChevronRight size={12} /></span>
        </button>
      </div>
      <Collapse open={open} id="rollback-details">
        <div className="rollback-panel">
          {undoEntries.length > 0 && (
            <section className="rollback-section" aria-label="逐步撤销">
              <div className="rollback-heading">逐步撤销（{undoEntries.length}）</div>
              <div className="undo-list">
                {[...undoEntries].reverse().map((entry, reverseIndex) => (
                  <div key={entry.id} className="undo-item">
                    <span className="undo-desc" title={entry.description}>{entry.description}</span>
                    <button className="btn tiny" disabled={disabled || undoBusy} onClick={() => onUndo(entry)}>
                      {reverseIndex === 0 ? "撤销" : `撤回到此处 (${reverseIndex + 1})`}
                    </button>
                  </div>
                ))}
              </div>
            </section>
          )}
          {(checkpoint || cpBusy) && (
            <section className="rollback-section" aria-label="整段恢复">
              <div className="rollback-heading">整段恢复（含删行、删表）</div>
              <div className="undo-item">
                <span className="undo-desc" title={checkpoint?.label}>
                  {cpBusy ? "正在处理检查点…" : `任务开始前 · ${checkpointMeta}`}
                </span>
                {checkpoint && (
                  <span className="checkpoint-actions">
                    <button className="btn tiny" disabled={disabled || cpBusy} onClick={() => onCheckpointAction("restore")}>
                      恢复到任务开始
                    </button>
                    <button className="btn tiny" disabled={disabled || cpBusy} onClick={() => onCheckpointAction("discard")}>
                      丢弃
                    </button>
                  </span>
                )}
              </div>
            </section>
          )}
          <button className="card-link rollback-timeline" onClick={onOpenTimeline}>
            查看完整操作记录（含修改前后对比）
          </button>
        </div>
      </Collapse>
    </div>
  );
}
