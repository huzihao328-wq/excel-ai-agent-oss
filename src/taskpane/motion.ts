import { useSyncExternalStore } from "react";

/** JS / CSS 共用的动效规范。业务状态从不等待动画完成。 */
export const MOTION = {
  duration: { press: 100, micro: 140, exit: 180, standard: 240, expand: 280, panel: 320, welcome: 400, stagger: 40 },
  easing: {
    out: "cubic-bezier(0.22, 1, 0.36, 1)",
    inOut: "cubic-bezier(0.4, 0, 0.2, 1)",
    spring: "cubic-bezier(0.22, 1.12, 0.36, 1)",
    sheet: "cubic-bezier(0.32, 0.72, 0, 1)",
  },
  distance: { hover: 1, micro: 4, content: 8, dialog: 12, panel: 24 },
  scale: { press: 0.98, iconPress: 0.94, enter: 0.985, chip: 0.97, pop: 0.92 },
  maxAnimatedItems: 36,
} as const;

/**
 * 物理弹簧（与 SwiftUI 相同的 response / dampingFraction 参数）。
 * 采样成 CSS linear() 曲线，CSS transition 与 WAAPI 共用；旧 WebView 不支持 linear() 时退回 cubic-bezier。
 */
interface SpringSpec { response: number; damping: number; fallback: string; fallbackDuration: number; }
const SPRING_SPECS = {
  /** 面板、对话框、消息入场：临界阻尼，无回弹，末段极长极柔 */
  smooth: { response: 0.38, damping: 1, fallback: "cubic-bezier(0.32, 0.72, 0, 1)", fallbackDuration: 380 },
  /** 按压回弹、选中指示、列表重排：轻微过冲 */
  snappy: { response: 0.32, damping: 0.78, fallback: "cubic-bezier(0.22, 1.08, 0.36, 1)", fallbackDuration: 300 },
  /** 图标、徽标、发送按钮等小元素弹出：可见回弹 */
  bouncy: { response: 0.4, damping: 0.66, fallback: "cubic-bezier(0.34, 1.4, 0.5, 1)", fallbackDuration: 420 },
  /** 列表 FLIP：短而稳，频繁筛选时不拖尾 */
  list: { response: 0.28, damping: 0.88, fallback: "cubic-bezier(0.22, 1, 0.36, 1)", fallbackDuration: 280 },
} satisfies Record<string, SpringSpec>;
export type SpringName = keyof typeof SPRING_SPECS;

const supportsLinearEasing = typeof CSS !== "undefined" && typeof CSS.supports === "function"
  && CSS.supports("transition-timing-function", "linear(0, 0.5 50%, 1)");

/** 数值积分阻尼振子，直到位移和速度都收敛；时长取稳定时刻，采样点均匀分布。 */
function sampleSpring({ response, damping }: SpringSpec): { easing: string; duration: number } {
  const omega = (2 * Math.PI) / response;
  const stiffness = omega * omega;
  const friction = 2 * damping * omega;
  const dt = 1 / 600;
  let x = 0, v = 0, t = 0, settledAt = 0;
  const trace: number[] = [];
  while (t < 2) {
    // 半隐式欧拉，步长足够小，结果与解析解误差 < 1e-4
    v += (stiffness * (1 - x) - friction * v) * dt;
    x += v * dt;
    t += dt;
    trace.push(x);
    if (Math.abs(1 - x) > 0.004 || Math.abs(v) > 0.05) settledAt = t; // 0.4% 以内肉眼不可分辨，截掉拖尾
    else if (t - settledAt > 0.05) break;
  }
  const duration = Math.round(settledAt * 1000 / 10) * 10;
  const steps = Math.min(64, Math.max(24, Math.round(settledAt * 90)));
  const stops: string[] = [];
  for (let i = 0; i <= steps; i++) {
    const index = Math.min(trace.length - 1, Math.round((i / steps) * settledAt / dt) - 1);
    const value = i === 0 ? 0 : i === steps ? 1 : trace[Math.max(0, index)];
    stops.push(String(Math.round(value * 10000) / 10000));
  }
  return { easing: `linear(${stops.join(", ")})`, duration };
}

export const SPRING = Object.fromEntries(Object.entries(SPRING_SPECS).map(([name, spec]) => [
  name, supportsLinearEasing ? sampleSpring(spec) : { easing: spec.fallback, duration: spec.fallbackDuration },
])) as Record<SpringName, { easing: string; duration: number }>;

const reducedQuery = typeof window !== "undefined" ? window.matchMedia?.("(prefers-reduced-motion: reduce)") : undefined;
export const prefersReducedMotion = () => reducedQuery?.matches ?? false;
export const subscribeMotionPreference = (notify: () => void) => {
  reducedQuery?.addEventListener("change", notify);
  return () => reducedQuery?.removeEventListener("change", notify);
};
export const useReducedMotion = () => useSyncExternalStore(subscribeMotionPreference, prefersReducedMotion, () => false);

/** 在首帧前安装；CSS 和 WAAPI 不分别维护两套数值。 */
export function installMotionTokens(): void {
  const style = document.documentElement.style;
  for (const [name, value] of Object.entries(MOTION.duration)) style.setProperty(`--motion-${name}`, `${value}ms`);
  for (const [name, value] of Object.entries(MOTION.easing)) style.setProperty(`--motion-ease-${name}`, value);
  for (const [name, value] of Object.entries(MOTION.distance)) style.setProperty(`--motion-distance-${name}`, `${value}px`);
  for (const [name, value] of Object.entries(MOTION.scale)) style.setProperty(`--motion-scale-${name}`, String(value));
  // --spring-x 可直接放进 transition / animation 简写：“时长 曲线”
  for (const [name, { easing, duration }] of Object.entries(SPRING)) {
    style.setProperty(`--spring-${name}`, `${duration}ms ${easing}`);
    style.setProperty(`--spring-${name}-duration`, `${duration}ms`);
    style.setProperty(`--spring-${name}-ease`, easing);
  }
}

/** 以弹簧方式滚动到位置；目标可以是函数（流式输出时底部持续变化）。用户滚轮 / 触摸 / 按下会立即中断。 */
export function springScroll(element: HTMLElement, target: number | (() => number), onDone?: () => void): () => void {
  const resolve = () => Math.max(0, Math.min(typeof target === "function" ? target() : target, element.scrollHeight - element.clientHeight));
  if (prefersReducedMotion()) { element.scrollTop = resolve(); onDone?.(); return () => undefined; }
  const { response, damping } = SPRING_SPECS.smooth;
  const omega = (2 * Math.PI) / response;
  let position = element.scrollTop, velocity = 0, last = performance.now(), frame = 0, stopped = false;
  const stop = () => {
    if (stopped) return;
    stopped = true;
    cancelAnimationFrame(frame);
    for (const type of ["wheel", "touchstart", "pointerdown", "keydown"]) element.removeEventListener(type, stop);
    onDone?.();
  };
  const step = (now: number) => {
    const dt = Math.min(0.032, (now - last) / 1000);
    last = now;
    const goal = resolve();
    velocity += (omega * omega * (goal - position) - 2 * damping * omega * velocity) * dt;
    position += velocity * dt;
    if (Math.abs(goal - position) < 0.5 && Math.abs(velocity) < 8) { element.scrollTop = goal; stop(); return; }
    element.scrollTop = position;
    frame = requestAnimationFrame(step);
  };
  for (const type of ["wheel", "touchstart", "pointerdown", "keydown"]) element.addEventListener(type, stop, { passive: true, once: true });
  frame = requestAnimationFrame(step);
  return stop;
}

const activeAnimations = new Set<Animation>();
/** 有限动画完成后释放合成层；运行时改变系统偏好也会立即收敛。 */
export function animate(element: Element, frames: Keyframe[], duration: number = MOTION.duration.standard, easing: string = MOTION.easing.out): Animation | undefined {
  if (prefersReducedMotion() || !element.animate) return;
  const animation = element.animate(frames, { duration, easing });
  activeAnimations.add(animation);
  const release = () => activeAnimations.delete(animation);
  animation.addEventListener("finish", release, { once: true });
  animation.addEventListener("cancel", release, { once: true });
  return animation;
}
reducedQuery?.addEventListener("change", () => {
  if (prefersReducedMotion()) for (const animation of activeAnimations) animation.finish();
});
