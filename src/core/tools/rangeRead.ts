/** read_range 读多少、怎么抽样：纯函数，便于不启动 Excel 也能核对。 */

export const MAX_READ_CELLS = 5000;
export const SAMPLE_HEAD_ROWS = 30;
export const SAMPLE_TAIL_ROWS = 5;

export type RangeReadPlan =
  | { mode: "full"; startOffset: number; rowCount: number }
  | { mode: "page"; startOffset: number; rowCount: number; truncated: boolean }
  | {
      mode: "sample";
      headRows: number;
      tailRows: number;
      totalRows: number;
      totalCols: number;
    };

function toInt(value: unknown, fallback: number): number {
  const n = Math.trunc(Number(value));
  return Number.isFinite(n) ? n : fallback;
}

/**
 * 决定一次读取覆盖哪些行。
 * - 未指定分页且总单元格 ≤ 5000：整读
 * - 指定了 start_row / max_rows：按页读，超限时自动收窄
 * - 否则：表头方向前 30 行 + 末 5 行抽样
 */
export function planRangeRead(
  rowCount: number,
  columnCount: number,
  options?: { start_row?: unknown; max_rows?: unknown }
): RangeReadPlan {
  const rows = Math.max(0, Math.trunc(rowCount) || 0);
  const cols = Math.max(1, Math.trunc(columnCount) || 1);
  if (rows === 0) return { mode: "full", startOffset: 0, rowCount: 0 };

  const hasPaging = options?.start_row != null || options?.max_rows != null;
  if (hasPaging) {
    const startRow1 = Math.max(1, toInt(options?.start_row, 1));
    const startOffset = Math.min(rows - 1, startRow1 - 1);
    let take = rows - startOffset;
    if (options?.max_rows != null) take = Math.min(take, Math.max(1, toInt(options.max_rows, 1)));
    const cap = Math.max(1, Math.floor(MAX_READ_CELLS / cols));
    const truncated = take > cap || startOffset + Math.min(take, cap) < rows;
    return { mode: "page", startOffset, rowCount: Math.min(take, cap), truncated };
  }

  if (rows * cols <= MAX_READ_CELLS) {
    return { mode: "full", startOffset: 0, rowCount: rows };
  }

  const headRows = Math.min(SAMPLE_HEAD_ROWS, rows);
  const tailRows = Math.min(SAMPLE_TAIL_ROWS, Math.max(0, rows - headRows));
  return { mode: "sample", headRows, tailRows, totalRows: rows, totalCols: cols };
}

export const SAMPLE_HINT =
  "区域超过 5000 单元格，已返回前 30 行和末 5 行抽样。请用 column_stats 做统计、group_summary 分组汇总、copy_range 复制、fill_down 填充公式，或用 start_row/max_rows 分页读取；不要假设表只有抽样这么多行。";
