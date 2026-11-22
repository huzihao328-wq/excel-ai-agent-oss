import { Children, Component, createRef, isValidElement, ReactElement, ReactNode } from "react";
import { animate, MOTION, prefersReducedMotion, SPRING, subscribeMotionPreference } from "../motion";

interface Props { children: ReactNode; className?: string; animateSize?: boolean; }
interface Entry { key: string; node: ReactElement; exitAt?: number; }
interface State { entries: Entry[]; }
interface Frame { rect: DOMRect; opacity: string; }
interface Snapshot { root: DOMRect; items: Map<string, Frame>; focusedKey?: string; }
const signature = (entries: Entry[]) => entries.map(entry => `${entry.key}:${entry.exitAt ?? ""}`).join("|");
const liveKeys = (entries: Entry[]) => entries.filter(entry => !entry.exitAt).map(entry => entry.key).join("|");

/**
 * 在 React 改 DOM 前读取可见位置，提交后批量读目标，再统一写入 FLIP。
 * 不监听每一帧、不在普通内容更新时测量；最多为 36 个可见条目创建动画。
 * 删除立即更新业务数据，短暂保留 inert 的视觉副本；同一 key 重现会复用节点。
 */
export default class MotionList extends Component<Props, State, Snapshot | null> {
  state: State = { entries: [] };
  private root = createRef<HTMLDivElement>();
  private nodes = new Map<string, HTMLDivElement>();
  private nodeRefs = new Map<string, (node: HTMLDivElement | null) => void>();
  private animations: Animation[] = [];
  private timer?: number;
  private unsubscribe?: () => void;

  componentDidMount(): void {
    this.unsubscribe = subscribeMotionPreference(() => {
      if (prefersReducedMotion()) this.setState(state => ({ entries: state.entries.filter(entry => !entry.exitAt) }));
    });
  }

  static getDerivedStateFromProps(props: Props, state: State): State {
    const children = Children.toArray(props.children).filter(isValidElement);
    const keys = new Set(children.map(child => String(child.key)));
    const entries: Entry[] = children.map(child => ({ key: String(child.key), node: child }));
    if (!prefersReducedMotion()) {
      state.entries.forEach((entry, index) => {
        if (!keys.has(entry.key)) entries.splice(Math.min(index, entries.length), 0, { ...entry, exitAt: entry.exitAt ?? Date.now() });
      });
    }
    return { entries };
  }

  getSnapshotBeforeUpdate(_props: Props, state: State): Snapshot | null {
    if (!this.root.current || prefersReducedMotion() || signature(state.entries) === signature(this.state.entries)) return null;
    // 只是清掉已播完的退场副本（absolute，不占位）：其他条目不会移动，不打断仍在进行的弹簧。
    const previousExits = new Set(state.entries.filter(entry => entry.exitAt).map(entry => entry.key));
    if (liveKeys(state.entries) === liveKeys(this.state.entries) && this.state.entries.every(entry => !entry.exitAt || previousExits.has(entry.key))) return null;
    const items = new Map<string, Frame>();
    // 所有读取发生在写入前；rect 包含尚未结束的动画，从实际视觉位置反向。
    const scrollport = this.root.current.closest('.history-list, .task-catalog, .composer')?.getBoundingClientRect();
    const top = Math.max(0, scrollport?.top ?? 0);
    const bottom = Math.min(innerHeight, scrollport?.bottom ?? innerHeight);
    for (const [key, node] of this.nodes) {
      const rect = node.getBoundingClientRect();
      if (rect.bottom >= top && rect.top <= bottom) items.set(key, { rect, opacity: getComputedStyle(node).opacity });
    }
    const focused = document.activeElement?.closest<HTMLElement>('[data-motion-key]');
    return { root: this.root.current.getBoundingClientRect(), items, focusedKey: focused && this.root.current.contains(focused) ? focused.dataset.motionKey : undefined };
  }

  componentDidUpdate(_props: Props, _state: State, snapshot: Snapshot | null): void {
    const root = this.root.current;
    if (snapshot && root) {
      this.animations.forEach(animation => animation.cancel());
      this.animations = [];
      for (const entry of this.state.entries) {
        if (entry.exitAt) continue;
        const node = this.nodes.get(entry.key);
        if (node?.style.left) { node.style.removeProperty("left"); node.style.removeProperty("top"); node.style.removeProperty("width"); node.style.removeProperty("height"); }
      }
      // Absolute 退场元素不参与目标布局，其他条目不必等待退出结束。
      const rootRect = root.getBoundingClientRect();
      const targets = new Map<string, DOMRect>();
      for (const [key, node] of this.nodes) targets.set(key, node.getBoundingClientRect());
      let count = 0;
      for (const entry of this.state.entries) {
        const node = this.nodes.get(entry.key);
        const to = targets.get(entry.key);
        if (!node || !to) continue;
        const from = snapshot.items.get(entry.key);
        const visibility = entry.exitAt && !from ? "hidden" : "";
        if (node.style.visibility !== visibility) node.style.visibility = visibility;
        if (entry.exitAt && from) {
          Object.assign(node.style, {
            left: `${from.rect.left - rootRect.left + root.scrollLeft}px`,
            top: `${from.rect.top - rootRect.top + root.scrollTop}px`,
            width: `${from.rect.width}px`, height: `${from.rect.height}px`,
          });
        }
        const visible = (rect: DOMRect) => rect.bottom >= 0 && rect.top <= innerHeight && rect.right >= 0 && rect.left <= innerWidth;
        if ((!visible(to) && (!from || !visible(from.rect))) || count >= MOTION.maxAnimatedItems) continue;
        let animation: Animation | undefined;
        if (entry.exitAt && from) {
          animation = animate(node, [{ opacity: from.opacity, transform: "none" }, { opacity: 0, transform: `translateY(-${MOTION.distance.micro}px)` }],
            Math.max(0, entry.exitAt + MOTION.duration.exit - Date.now()));
        } else if (from) {
          const x = (from.rect.left - snapshot.root.left) - (to.left - rootRect.left);
          const y = (from.rect.top - snapshot.root.top) - (to.top - rootRect.top);
          if (Math.abs(x) + Math.abs(y) > 0.5 || Number(from.opacity) < 1) {
            animation = animate(node, [{ transform: `translate(${x}px, ${y}px)`, opacity: from.opacity }, { transform: "none", opacity: 1 }], SPRING.list.duration, SPRING.list.easing);
          }
        } else {
          animation = animate(node, [{ opacity: 0, transform: `translateY(${MOTION.distance.content}px) scale(${MOTION.scale.enter})` }, { opacity: 1, transform: "none" }], SPRING.list.duration, SPRING.list.easing);
        }
        if (animation) { this.animations.push(animation); count++; }
      }
      const heightDelta = Math.abs(snapshot.root.height - rootRect.height);
      // 长列表筛选会改变数千像素高度；只移动可见条目，不逐帧重排整个滚动区域。
      if (this.props.animateSize !== false && heightDelta > 1 && heightDelta < innerHeight / 2) {
        const height = animate(root, [{ height: `${snapshot.root.height}px` }, { height: `${rootRect.height}px` }], SPRING.list.duration, SPRING.list.easing);
        if (height) this.animations.push(height);
      }
      if (snapshot.focusedKey && this.state.entries.some(entry => entry.key === snapshot.focusedKey && entry.exitAt)) {
        root.querySelector<HTMLElement>('.motion-list-item:not(.is-removing) button:not(:disabled)')?.focus({ preventScroll: true });
      }
    }
    this.scheduleRemoval();
  }

  private scheduleRemoval(): void {
    window.clearTimeout(this.timer);
    const exits = this.state.entries.flatMap(entry => entry.exitAt === undefined ? [] : [entry.exitAt]);
    if (!exits.length) return;
    this.timer = window.setTimeout(() => {
      const now = Date.now();
      this.setState(state => ({ entries: state.entries.filter(entry => entry.exitAt === undefined || now < entry.exitAt + MOTION.duration.exit) }));
    }, Math.max(0, Math.min(...exits) + MOTION.duration.exit - Date.now()));
  }

  componentWillUnmount(): void {
    this.unsubscribe?.();
    window.clearTimeout(this.timer);
    this.animations.forEach(animation => animation.cancel());
  }

  private refFor(key: string): (node: HTMLDivElement | null) => void {
    let ref = this.nodeRefs.get(key);
    if (!ref) {
      ref = node => { if (node) this.nodes.set(key, node); else { this.nodes.delete(key); this.nodeRefs.delete(key); } };
      this.nodeRefs.set(key, ref);
    }
    return ref;
  }

  render() {
    return <div ref={this.root} className={`motion-list ${this.props.className ?? ""}`}>
      {this.state.entries.map(entry => <div
        key={entry.key} data-motion-key={entry.key}
        ref={this.refFor(entry.key)} {...{ inert: entry.exitAt ? "" : undefined }}
        className={`motion-list-item${entry.exitAt ? " is-removing" : ""}`}
        aria-hidden={entry.exitAt ? true : undefined}
      >{entry.node}</div>)}
    </div>;
  }
}
