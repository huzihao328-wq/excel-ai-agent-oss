import { ReactNode } from "react";
import { IconClose } from "./Icons";

/** 抽屉 / 面板统一的标题栏：图标 + 标题 + 关闭。关闭按钮保持 title，供键盘焦点和测试定位。 */
export function PanelHeader({ icon, title, onClose, closeTitle = "关闭", closeLabel, children }: {
  icon?: ReactNode;
  title: string;
  onClose: () => void;
  closeTitle?: string;
  /** 需要文字按钮时（如「收起」）传入；默认是 × 图标 */
  closeLabel?: string;
  children?: ReactNode;
}) {
  return (
    <div className="history-header">
      <h2>{icon && <span className="panel-title-icon" aria-hidden="true">{icon}</span>}{title}</h2>
      {children}
      <button className={closeLabel ? "btn tiny" : "icon-btn"} onClick={onClose} title={closeTitle} aria-label={closeLabel ? undefined : closeTitle}>
        {closeLabel ?? <IconClose size={15} />}
      </button>
    </div>
  );
}

/** 空状态：图标 + 一句话 + 下一步提示，代替孤零零的一行灰字 */
export function EmptyState({ icon, title, hint, action }: { icon: ReactNode; title: string; hint?: string; action?: ReactNode }) {
  return (
    <div className="panel-empty">
      <div className="panel-empty-icon" aria-hidden="true">{icon}</div>
      <div className="panel-empty-title">{title}</div>
      {hint && <p className="panel-empty-hint">{hint}</p>}
      {action}
    </div>
  );
}
