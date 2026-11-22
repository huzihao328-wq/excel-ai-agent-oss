import { CSSProperties, ReactNode, useLayoutEffect, useRef, useState } from "react";
import { MOTION, useReducedMotion } from "../motion";
import { useOverlayFocus } from "../hooks/useOverlayFocus";

interface Props {
  /** 是否显示。变为 false 后子树先带上 is-exiting 播完退场动画，再真正卸载 */
  open: boolean;
  /** 退场时长（ms），同时传给 CSS 过渡和卸载计时器 */
  duration?: number;
  /** 关闭后保留挂载（display:none），用于需要保留内部状态的面板 */
  keepMounted?: boolean;
  /** 打开不足这么久就被关闭时直接卸载、不播退场，避免刚渲染就被替换的内容闪一下 */
  minOpenMs?: number;
  focus?: "dialog" | "popover";
  /** 挂载时已打开是否播进场（默认播）。替换现有内容的场景关掉，避免高度从 0 长起造成跳动 */
  appear?: boolean;
  children: ReactNode;
}

type Phase = "entering" | "open" | "exiting" | "closed";

/**
 * 延迟卸载容器：用 display: contents 包一层，不影响子元素的定位与布局，
 * 通过 CSS transition 连接进退场状态，快速反向操作时从当前画面继续过渡。
 */
export default function Presence({ open, duration = MOTION.duration.exit, keepMounted = false, minOpenMs = 0, focus, appear = true, children }: Props) {
  const reducedMotion = useReducedMotion();
  const [phase, setPhase] = useState<Phase>(open ? (appear ? "entering" : "open") : "closed");
  const root = useRef<HTMLDivElement>(null);
  const openedAt = useRef(open ? Date.now() : 0);
  // 退场期间子树用最后一次打开时的内容，避免父组件已把数据置空（如 previewFile=null）导致子组件出错
  const lastChildren = useRef(children);
  if (open) lastChildren.current = children;

  useLayoutEffect(() => {
    if (open) {
      openedAt.current = Date.now();
      setPhase((prev) => {
        if (reducedMotion) return "open";
        return prev === "closed" ? "entering" : prev === "exiting" ? "open" : prev;
      });
      return;
    }
    setPhase((prev) => {
      if (prev === "closed") return prev;
      const shownFor = Date.now() - openedAt.current;
      if (duration <= 0 || shownFor < minOpenMs || reducedMotion) return "closed";
      return "exiting";
    });
  }, [open, duration, minOpenMs, reducedMotion]);

  useOverlayFocus(root, open, phase !== "closed", focus);

  useLayoutEffect(() => {
    if (phase === "entering") {
      // 先绘制起点，再切到目标；不强制同步布局，也不依赖较新的 @starting-style。
      let frame = requestAnimationFrame(() => {
        frame = requestAnimationFrame(() => setPhase("open"));
      });
      return () => cancelAnimationFrame(frame);
    }
    if (phase === "exiting") {
      const timer = window.setTimeout(() => setPhase("closed"), duration);
      return () => window.clearTimeout(timer);
    }
  }, [phase, duration]);

  useLayoutEffect(() => {
    // 退场的内容仍在 DOM 中，但不能继续接收 Tab 焦点或点击。
    if (root.current) root.current.inert = !open;
  }, [open, phase]);

  if (phase === "closed" && !keepMounted) return null;
  return (
    <div
      ref={root}
      className={`presence${phase === "entering" ? " is-entering" : phase === "exiting" ? " is-exiting" : ""}`}
      style={{
        display: phase === "closed" ? "none" : "contents",
        "--presence-exit-duration": `${duration}ms`,
      } as CSSProperties}
      aria-hidden={!open || undefined}
    >
      {keepMounted || open ? children : lastChildren.current}
    </div>
  );
}
