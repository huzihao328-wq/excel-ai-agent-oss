import { RefObject, useLayoutEffect, useRef } from "react";
import { animate, MOTION } from "../motion";

const MAX_HEIGHT = 160;

/** Resize for wrapped text as well as explicit newlines, including pane resizes. */
export function useAutoSizeTextarea(ref: RefObject<HTMLTextAreaElement>, value: string): void {
  const animation = useRef<Animation>();
  const measured = useRef(false);
  const resize = (element: HTMLTextAreaElement) => {
    const from = element.getBoundingClientRect().height;
    animation.current?.cancel();
    element.style.height = "auto";
    const height = element.scrollHeight;
    element.style.height = `${Math.min(MAX_HEIGHT, height)}px`;
    element.style.overflowY = height > MAX_HEIGHT ? "auto" : "hidden";
    const to = element.getBoundingClientRect().height;
    if (measured.current && Math.abs(from - to) > 1) {
      animation.current = animate(element, [{ height: `${from}px` }, { height: `${to}px` }], MOTION.duration.micro);
    }
    measured.current = true;
  };
  // 内容变化 → 只重算高度
  useLayoutEffect(() => {
    const element = ref.current;
    if (element) resize(element);
  }, [ref, value]);

  // 窗格宽度变化 → 换行位置变了也要重算；观察器只建一次，不随每次按键重建
  useLayoutEffect(() => {
    const element = ref.current;
    if (!element || typeof ResizeObserver === "undefined") return;
    let width = element.clientWidth;
    const observer = new ResizeObserver(() => {
      if (width === element.clientWidth) return;
      width = element.clientWidth;
      resize(element);
    });
    observer.observe(element);
    return () => { observer.disconnect(); animation.current?.cancel(); };
  }, [ref]);
}
