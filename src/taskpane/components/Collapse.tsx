import { ReactNode, useLayoutEffect, useRef } from "react";
import { MOTION } from "../motion";
import Presence from "./Presence";

/** 真实高度参与布局；CSS grid 过渡可随时反向，关闭内容同时失去交互能力。 */
export default function Collapse({ open, children, id, className = "", duration = MOTION.duration.expand, appear = true }: {
  open: boolean; children: ReactNode; id?: string; className?: string; duration?: number;
  /** 挂载时已展开是否从 0 高度长出（默认是） */
  appear?: boolean;
}) {
  const root = useRef<HTMLDivElement>(null);
  const opener = useRef<HTMLElement | null>(null);
  const wasOpen = useRef(false);
  if (open && !wasOpen.current && !root.current?.contains(document.activeElement)) opener.current = document.activeElement as HTMLElement;
  wasOpen.current = open;
  useLayoutEffect(() => {
    if (!open && root.current?.contains(document.activeElement)) {
      queueMicrotask(() => {
        if (!wasOpen.current && opener.current?.isConnected && !opener.current.closest('[inert]')) opener.current.focus({ preventScroll: true });
      });
    }
  }, [open]);
  return (
    <Presence open={open} duration={duration} appear={appear}>
      <div ref={root} id={id} className={`motion-collapse ${className}`}>
        <div className="motion-collapse-inner">{children}</div>
      </div>
    </Presence>
  );
}
