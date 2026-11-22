import { useEffect, useState } from "react";
import MotionList from "./MotionList";
import { PanelHeader } from "./Panel";
import { IconBookmark, IconClose } from "./Icons";
import {
  MemoryEntry,
  MemoryScope,
  clearMemories,
  loadAllMemories,
  onMemoryChange,
  removeMemory,
  upsertMemory,
} from "../../core/memory";

interface Props {
  onClose: () => void;
}

export default function MemoryPanel({ onClose }: Props) {
  const [items, setItems] = useState<MemoryEntry[]>(loadAllMemories);
  const [draft, setDraft] = useState("");
  const [scope, setScope] = useState<MemoryScope>("workbook");
  const [error, setError] = useState("");
  // 「清空」用面板内二次确认，不用 window.confirm：任务窗格里原生弹窗可能被宿主吞掉，点了没反应
  const [pendingClear, setPendingClear] = useState<MemoryScope | null>(null);

  useEffect(() => {
    setItems(loadAllMemories());
    return onMemoryChange(() => setItems(loadAllMemories()));
  }, []);

  const global = items.filter((e) => e.scope === "global");
  const workbook = items.filter((e) => e.scope === "workbook");

  const add = () => {
    setError("");
    try {
      upsertMemory({ content: draft, scope });
      setDraft("");
    } catch (e: any) {
      setError(e?.message || String(e));
    }
  };

  return (
    <div className="history-overlay" onClick={onClose}>
      <div className="history-panel memory-panel" onClick={(e) => e.stopPropagation()}>
        <PanelHeader icon={<IconBookmark size={16} />} title="长期记忆" onClose={onClose} />

        <p className="memory-hint">
          全局习惯存在本机，升级插件后会自动继承；当前工作簿的约定会跟着这个文件走。AI 做事时会自动带上，也可以自己写入。
        </p>

        <div className="memory-add">
          <textarea
            value={draft}
            rows={2}
            placeholder="例如：销售表用订单号和客户表关联"
            onChange={(e) => setDraft(e.target.value)}
          />
          <div className="memory-add-row">
            <select value={scope} onChange={(e) => setScope(e.target.value as MemoryScope)}>
              <option value="workbook">本工作簿</option>
              <option value="global">全局习惯</option>
            </select>
            <button className="btn primary" disabled={!draft.trim()} onClick={add}>
              记住
            </button>
          </div>
          {error && <div className="memory-error">{error}</div>}
        </div>

        <div className="history-list">
          <MemoryGroup
            title="当前工作簿"
            items={workbook}
            empty="还没有这个文件的约定"
            confirming={pendingClear === "workbook"}
            onClearRequest={() => setPendingClear("workbook")}
            onClearCancel={() => setPendingClear(null)}
            onClear={() => {
              clearMemories("workbook");
              setPendingClear(null);
            }}
            onDelete={(id) => removeMemory(id)}
          />
          <MemoryGroup
            title="全局习惯"
            items={global}
            empty="还没有全局习惯"
            confirming={pendingClear === "global"}
            onClearRequest={() => setPendingClear("global")}
            onClearCancel={() => setPendingClear(null)}
            onClear={() => {
              clearMemories("global");
              setPendingClear(null);
            }}
            onDelete={(id) => removeMemory(id)}
          />
        </div>
      </div>
    </div>
  );
}

function MemoryGroup({
  title,
  items,
  empty,
  confirming,
  onClearRequest,
  onClearCancel,
  onClear,
  onDelete,
}: {
  title: string;
  items: MemoryEntry[];
  empty: string;
  confirming: boolean;
  onClearRequest: () => void;
  onClearCancel: () => void;
  onClear: () => void;
  onDelete: (id: string) => void;
}) {
  return (
    <div className="history-group">
      <div className="history-group-label memory-group-head">
        <span>
          {title} · {items.length}
        </span>
        {items.length > 0 && (confirming ? (
          <div className="history-item-confirm">
            <button className="btn tiny danger" onClick={onClear}>确定清空</button>
            <button className="btn tiny" onClick={onClearCancel}>取消</button>
          </div>
        ) : (
          <button className="memory-clear" onClick={onClearRequest}>
            清空
          </button>
        ))}
      </div>
      <MotionList>
      {items.length === 0 && <div key="empty" className="history-empty">{empty}</div>}
      {items.map((e) => (
        <div key={e.id} className="memory-item">
          <div className="memory-item-main">
            <div className="memory-item-text">{e.content}</div>
            {e.key && <div className="memory-item-key">{e.key}</div>}
          </div>
          <button className="history-item-del" title="忘掉" onClick={() => onDelete(e.id)}>
            <IconClose size={13} />
          </button>
        </div>
      ))}</MotionList>
    </div>
  );
}
