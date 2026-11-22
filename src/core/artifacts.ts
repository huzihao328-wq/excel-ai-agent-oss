/*
 * 结果引用：登记本会话 AI 产出的表/区域/图表/透视，随上下文注入
 * （[最近产出] 块），并在输入框上方显示为可点击的引用胶囊。
 * 内存会话态，保留最近若干条。
 */
import { Artifact } from "./types";

export type { Artifact } from "./types";

const MAX_ARTIFACTS = 12;

let artifacts: Artifact[] = [];
let listeners: (() => void)[] = [];

export function getArtifacts(): Artifact[] {
  return artifacts;
}

export function onArtifactsChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((f) => f !== fn);
  };
}

export function registerArtifact(a: Omit<Artifact, "id" | "createdAt">): void {
  const id = `art_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  // 同名同位置的产出视为更新（如重画同一张汇总表）
  artifacts = artifacts.filter((x) => !(x.name === a.name && x.sheet === a.sheet && x.kind === a.kind));
  artifacts.unshift({ ...a, id, createdAt: Date.now() });
  if (artifacts.length > MAX_ARTIFACTS) artifacts = artifacts.slice(0, MAX_ARTIFACTS);
  listeners.forEach((f) => f());
}

export function clearArtifacts(): void {
  if (artifacts.length === 0) return;
  artifacts = [];
  listeners.forEach((f) => f());
}

function ago(ts: number): string {
  const s = Math.max(1, Math.round((Date.now() - ts) / 1000));
  if (s < 60) return `${s} 秒前`;
  if (s < 3600) return `${Math.round(s / 60)} 分钟前`;
  return `${Math.round(s / 3600)} 小时前`;
}

const KIND_LABELS: Record<Artifact["kind"], string> = {
  sheet: "Sheet",
  range: "区域",
  chart: "图表",
  table: "Excel 表",
  pivot: "透视表",
};

/** 注入上下文的 [最近产出] 块 */
export function formatArtifactsBlock(): string {
  if (artifacts.length === 0) return "";
  const lines = artifacts.slice(0, 8).map((a, i) => {
    const pos = a.address ? `${a.sheet}!${a.address}` : a.sheet;
    return `${i + 1}. 「${a.name}」（${KIND_LABELS[a.kind]}）${pos} · ${ago(a.createdAt)} · ${a.brief}`;
  });
  return [
    "[最近产出]",
    "以下是你在本会话生成的最近产出。用户说「刚才的结果」「上次生成的表」「那个图表/透视」时指的就是它们：对照定位后直接在其 Sheet/地址上继续加工，不要重新生成一份。",
    ...lines,
  ].join("\n");
}
