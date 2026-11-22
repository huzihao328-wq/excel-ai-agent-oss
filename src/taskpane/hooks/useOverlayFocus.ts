import { RefObject, useLayoutEffect, useRef } from "react";

const layers: HTMLElement[] = [];
const focusable = (root: HTMLElement) => Array.from(root.querySelectorAll<HTMLElement>(
  'button:not(:disabled), input:not(:disabled), textarea:not(:disabled), select:not(:disabled), a[href], [tabindex="0"]'
)).filter(element => !element.closest('[inert], [aria-hidden="true"]') && element.getClientRects().length > 0);

/** 模态层限制 Tab；普通菜单只管理打开/返回焦点，不改变原有点击路径。 */
export function useOverlayFocus(root: RefObject<HTMLDivElement>, open: boolean, mounted: boolean, mode: "dialog" | "popover" | undefined): void {
  const trigger = useRef<HTMLElement | null>(null);
  const wasOpen = useRef(false);
  // 在子组件 autoFocus / layout effect 运行前保存入口。
  if (open && !wasOpen.current && mode && !root.current?.contains(document.activeElement)) trigger.current = document.activeElement as HTMLElement;
  wasOpen.current = open;

  useLayoutEffect(() => {
    const element = root.current;
    if (!mode || !open || !mounted || !element) return;
    layers.push(element);
    const panel = element.querySelector<HTMLElement>('[role="dialog"], [role="alertdialog"], .history-panel, .preview-panel') ?? element;
    if (mode === "dialog") {
      panel.setAttribute("tabindex", "-1");
      if (!panel.hasAttribute("role")) panel.setAttribute("role", "dialog");
      panel.setAttribute("aria-modal", "true");
      if (!panel.hasAttribute("aria-label") && !panel.hasAttribute("aria-labelledby")) {
        panel.setAttribute("aria-label", panel.querySelector('h2, .preview-title')?.textContent ?? "对话框");
      }
    }
    if (!element.contains(document.activeElement)) (focusable(element)[0] ?? panel).focus({ preventScroll: true });
    const isTop = () => layers[layers.length - 1] === element;
    const onKey = (event: KeyboardEvent) => {
      if (!isTop() || event.defaultPrevented || event.isComposing || event.key !== "Tab" || mode !== "dialog") return;
      const targets = focusable(element);
      const index = targets.indexOf(document.activeElement as HTMLElement);
      if (index < 0 || (!event.shiftKey && index === targets.length - 1) || (event.shiftKey && index === 0)) {
        event.preventDefault();
        (targets[event.shiftKey ? targets.length - 1 : 0] ?? panel).focus({ preventScroll: true });
      }
    };
    const onFocus = (event: FocusEvent) => {
      if (mode === "dialog" && isTop() && !element.contains(event.target as Node)) (focusable(element)[0] ?? panel).focus({ preventScroll: true });
    };
    document.addEventListener("keydown", onKey);
    document.addEventListener("focusin", onFocus);
    return () => {
      const wasTop = isTop();
      const index = layers.indexOf(element);
      if (index >= 0) layers.splice(index, 1);
      document.removeEventListener("keydown", onKey);
      document.removeEventListener("focusin", onFocus);
      // 已经由业务流程转移到输入框/下一层时，不抢回焦点。
      // layout cleanup 属于 mutation 阶段，React 随后会恢复 selection；等 commit 完成再返回焦点。
      queueMicrotask(() => {
        if (layers.includes(element)) return;
        if (wasTop && (element.contains(document.activeElement) || document.activeElement === document.body)) {
          const target = trigger.current;
          if (target?.isConnected && !target.closest('[inert], [aria-hidden="true"]')) target.focus({ preventScroll: true });
          else document.querySelector<HTMLTextAreaElement>('.input-row textarea')?.focus({ preventScroll: true });
        }
      });
    };
  }, [root, open, mounted, mode]);
}
