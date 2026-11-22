import { useId, useLayoutEffect, useRef } from "react";

interface Props {
  options: { value: string; label: string }[];
  value: string;
  onChange: (value: string) => void;
  label: string;
  className: string;
  buttonClassName?: string;
  panelId?: string;
}

/** 选中背景是同一个元素；文字和业务状态立即更新，背景接续当前轨迹。 */
export default function MotionTabs({ options, value, onChange, label, className, buttonClassName = "", panelId }: Props) {
  const root = useRef<HTMLDivElement>(null);
  const id = useId();
  useLayoutEffect(() => {
    const element = root.current;
    if (!element) return;
    let frame = 0;
    const measure = () => {
      const selected = element.querySelector<HTMLElement>('[data-selected="true"]');
      if (!selected) return;
      element.style.setProperty("--selection-x", `${selected.offsetLeft}px`);
      element.style.setProperty("--selection-y", `${selected.offsetTop}px`);
      element.style.setProperty("--selection-width", `${selected.offsetWidth}px`);
      element.style.setProperty("--selection-height", `${selected.offsetHeight}px`);
    };
    measure();
    frame = requestAnimationFrame(() => { element.dataset.ready = "true"; });
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    for (const button of element.querySelectorAll('button')) observer.observe(button);
    return () => { observer.disconnect(); cancelAnimationFrame(frame); };
  }, [value, options.map(option => option.value + option.label).join("\0")]);
  return (
    <div ref={root} className={`motion-tabs ${className}`} role={panelId ? "tablist" : "group"} aria-label={label} onKeyDown={event => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
      event.preventDefault();
      const index = options.findIndex(option => option.value === value);
      const next = event.key === "Home" ? 0 : event.key === "End" ? options.length - 1
        : (index + (event.key === "ArrowRight" ? 1 : -1) + options.length) % options.length;
      onChange(options[next].value);
      const button = root.current?.querySelectorAll('button')[next];
      button?.focus({ preventScroll: true });
      button?.scrollIntoView({ block: "nearest", inline: "nearest" });
    }}>
      <span className="motion-tab-indicator" aria-hidden="true" />
      {options.map(option => <button
        key={option.value} type="button" id={`${id}-${option.value}`} role={panelId ? "tab" : undefined}
        aria-selected={panelId ? value === option.value : undefined}
        aria-pressed={panelId ? undefined : value === option.value}
        aria-controls={panelId} tabIndex={value === option.value ? 0 : -1}
        data-selected={value === option.value}
        className={`${buttonClassName}${value === option.value ? " active" : ""}`}
        onClick={() => onChange(option.value)}
      >{option.label}</button>)}
    </div>
  );
}
