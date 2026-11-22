/*
 * 操作时间线：会话内记录每次对工作簿的改动（含前后内容预览），
 * 供"操作记录"面板回放。内存态，刷新/新对话后清空。
 */
import { OpLogEntry, OpKind } from "./types";

const MAX_LOG = 100;

let log: OpLogEntry[] = [];
let listeners: (() => void)[] = [];

export function getOpLog(): OpLogEntry[] {
  return log;
}

export function onOpLogChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((f) => f !== fn);
  };
}

export function pushOpLog(entry: Omit<OpLogEntry, "id" | "timestamp">): void {
  log.push({
    ...entry,
    id: `op_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    timestamp: Date.now(),
  });
  if (log.length > MAX_LOG) log.shift();
  listeners.forEach((f) => f());
}

export function clearOpLog(): void {
  if (log.length === 0) return;
  log = [];
  listeners.forEach((f) => f());
}

/** 把二维网格渲染成文本预览（与确认卡片同风格，供 before/after 对比） */
export function gridPreview(grid: unknown[][] | undefined, maxRows = 8, maxCols = 8): string | undefined {
  if (!grid || grid.length === 0) return undefined;
  const rows = grid.slice(0, maxRows).map((r) =>
    (r || [])
      .slice(0, maxCols)
      .map((v) => (v === "" || v === null || v === undefined ? "·" : String(v).slice(0, 14)))
      .join(" | ")
  );
  const more = grid.length > maxRows ? `\n…共 ${grid.length} 行` : "";
  return rows.join("\n") + more;
}

export const OP_KIND_LABELS: Record<OpKind, string> = {
  write: "写入",
  structure: "结构",
  format: "格式",
  import: "导入",
  chart: "图表",
  table: "表格",
  undo: "撤销",
  other: "其他",
};
