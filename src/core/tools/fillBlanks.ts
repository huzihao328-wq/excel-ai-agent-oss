/**
 * 取消合并后补齐、空白单元格按上方/左侧/固定值填充的纯计算。
 * 坐标都是相对传入网格的 0 基行列；由 excelTools.ts 负责读表、取消合并和写回。
 */
import { isEmptyCell } from "./safeValues";
import { isRowBlank } from "./cleanValues";

export type FillMode = "above" | "left" | "value";

export interface MergedArea {
  row: number;
  column: number;
  rows: number;
  columns: number;
}

export interface CellRef {
  row: number;
  column: number;
}

export interface FillCell extends CellRef {
  value: unknown;
  /** 值从哪个格来（用来沿用它的数字格式）；mode=value 时为 null */
  from: CellRef | null;
}

export interface FillPlan {
  cells: FillCell[];
  /** 合并区域补齐的格数 */
  mergedCells: number;
  /** 各列（网格内列号）按 mode 填充的空白格数，不含合并区域补齐 */
  perColumn: Record<number, number>;
  /** 上方/左侧没有可用值、只能留空的格数 */
  unfilled: number;
}

export interface FillOptions {
  /** 要填空白的列（网格内列号）；为空时只补齐合并区域 */
  columns: number[];
  mode: FillMode;
  /** mode=value 时填入的值 */
  value?: unknown;
  /** 从这一行开始处理（跳过表头），网格内行号 */
  startRow: number;
  merged?: MergedArea[];
  /** Excel 的 formulas 网格含公式及常量，可区分空结果公式和真正的空格。 */
  formulas?: unknown[][];
}

/** 合并区域只有左上角有值：竖向跨多行的合并（分类、部门这类）才补齐，单行横向合并多是标题，保持原样 */
export function isFillableMerge(area: MergedArea, startRow: number): boolean {
  return area.rows > 1 && area.row + area.rows > startRow;
}

export function planBlankFill(values: unknown[][], options: FillOptions): FillPlan {
  const work = values.map((row) => [...row]);
  const origin: (CellRef | null)[][] = values.map((row, r) => row.map((_, c) => ({ row: r, column: c })));
  const cells = new Map<string, FillCell>();
  const canFill = (r: number, c: number) => isEmptyCell(work[r][c]) && isEmptyCell(options.formulas?.[r]?.[c]);
  const set = (row: number, column: number, value: unknown, from: CellRef | null) => {
    work[row][column] = value;
    origin[row][column] = from;
    cells.set(`${row},${column}`, { row, column, value, from });
  };

  let mergedCells = 0;
  for (const area of options.merged || []) {
    if (!isFillableMerge(area, options.startRow)) continue;
    const top = values[area.row]?.[area.column];
    if (isEmptyCell(top)) continue;
    for (let r = Math.max(area.row, options.startRow); r < Math.min(area.row + area.rows, values.length); r++) {
      for (let c = area.column; c < area.column + area.columns && c < (values[r]?.length ?? 0); c++) {
        if ((r === area.row && c === area.column) || !canFill(r, c)) continue;
        set(r, c, top, { row: area.row, column: area.column });
        mergedCells++;
      }
    }
  }

  const perColumn: Record<number, number> = {};
  let unfilled = 0;
  const columns = [...new Set(options.columns)].sort((a, b) => a - b);
  // 整行都空的是空行（分隔行或已用区域尾巴），不是合并留下的空白
  const blankRow = (options.formulas ?? values).map(isRowBlank);
  const fill = (r: number, c: number, source: CellRef | null) => {
    const value = source ? work[source.row][source.column] : options.value;
    set(r, c, value, source ? origin[source.row][source.column] : null);
    perColumn[c] = (perColumn[c] || 0) + 1;
  };

  for (let r = options.startRow; r < work.length; r++) {
    if (blankRow[r]) continue;
    for (const c of columns) {
      if (c >= work[r].length || !canFill(r, c)) continue;
      if (options.mode === "value") {
        fill(r, c, null);
        continue;
      }
      const source = options.mode === "left" ? (c > 0 ? { row: r, column: c - 1 } : null) : prevDataRow(r, c);
      if (!source || isEmptyCell(work[source.row][source.column])) unfilled++;
      else fill(r, c, source);
    }
  }

  function prevDataRow(r: number, c: number): CellRef | null {
    for (let p = r - 1; p >= options.startRow; p--) {
      if (!blankRow[p]) return { row: p, column: c };
    }
    return null;
  }

  return { cells: [...cells.values()], mergedCells, perColumn, unfilled };
}

export interface FillRun {
  row: number;
  column: number;
  cells: FillCell[];
}

/** 同一列里行号连续的填充格合成一段，一段写一次，减少 Office.js 往返 */
export function fillRuns(cells: FillCell[]): FillRun[] {
  const sorted = [...cells].sort((a, b) => a.column - b.column || a.row - b.row);
  const runs: FillRun[] = [];
  for (const cell of sorted) {
    const last = runs[runs.length - 1];
    if (last && last.column === cell.column && last.row + last.cells.length === cell.row) last.cells.push(cell);
    else runs.push({ row: cell.row, column: cell.column, cells: [cell] });
  }
  return runs;
}

/** 填充格的外接矩形（网格内 0 基行列，含边界） */
export function fillBounds(cells: CellRef[]): { top: number; left: number; bottom: number; right: number } | null {
  if (!cells.length) return null;
  let top = Infinity;
  let left = Infinity;
  let bottom = -1;
  let right = -1;
  for (const c of cells) {
    top = Math.min(top, c.row);
    left = Math.min(left, c.column);
    bottom = Math.max(bottom, c.row);
    right = Math.max(right, c.column);
  }
  return { top, left, bottom, right };
}
