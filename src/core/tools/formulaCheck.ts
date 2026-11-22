/*
 * 写公式类工具的自检：写入后 Excel 已经算出结果，直接在工具结果里告诉模型有没有错误值、错在哪、可能的原因，
 * 省掉「写完再 read_range 读回抽查」这一轮，也避免模型只抽查前几行漏掉中间的 #N/A。
 */
import { colLetterToIndex, indexToColLetter } from "./formulaRefs";

const ERROR_RE = /^#(N\/A|REF!|NAME\?|VALUE!|DIV\/0!|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|BUSY!|FIELD!|CONNECT!|BLOCKED!|UNKNOWN!)/;
const SAMPLE_ADDRESSES = 5;
const SAMPLE_ROWS = 3;
const SAMPLE_COLS = 8;

const ERROR_HINTS: Record<string, string> = {
  "#N/A": "查找没匹配上：核对两边键的类型是否一致（文本数字 vs 数字）、有无首尾空格，或确实不存在",
  "#NAME?": "函数名拼错或当前 Excel 不支持（XLOOKUP/FILTER/UNIQUE 等），改用 INDEX+MATCH、SUMIFS 等",
  "#REF!": "引用失效：引用的行列/Sheet 被删除，或引用越界",
  "#VALUE!": "参与运算的单元格类型不对（例如文本参与加减），先清洗成数字或用 VALUE/IFERROR",
  "#DIV/0!": "除数为 0 或空，可用 IF(分母=0,\"\",…) 或 IFERROR 包一层",
  "#SPILL!": "动态数组的溢出区域被已有内容挡住，清空下方/右侧单元格或改写到空白处",
  "#NUM!": "数值无效或超出范围（如负数开方、日期超界）",
  "#NULL!": "区域交集为空，多半是区域之间少写了逗号",
  "#CALC!": "计算引擎无法得出结果（如 FILTER 无结果时未提供默认值）",
};

export interface FormulaCheck {
  /** 公式单元格数 */
  formula_cells: number;
  /** 错误值总数；0 表示全部算出结果 */
  error_cells: number;
  /** 每种错误的数量、示例地址和常见原因 */
  errors?: { type: string; count: number; cells: string[]; hint?: string }[];
  /** 结果为空文本的公式单元格数（查找没匹配上且设了 if_not_found="" 时也会是空） */
  blank_results: number;
  /** 前几行算出的值，便于确认结果说得通 */
  sample: unknown[][];
}

function parseStart(address: string): { col: number; row: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)/.exec(address.split("!").pop() || "");
  if (!m) return { col: 0, row: 1 };
  return { col: colLetterToIndex(m[1].toUpperCase()), row: Number(m[2]) };
}

export function checkFormulaResults(address: string, values: unknown[][], formulas: unknown[][]): FormulaCheck {
  const start = parseStart(address);
  const byType = new Map<string, { count: number; cells: string[] }>();
  let formulaCells = 0;
  let errorCells = 0;
  let blanks = 0;
  values.forEach((row, r) => {
    row.forEach((value, c) => {
      const formula = formulas[r]?.[c];
      if (typeof formula !== "string" || !formula.startsWith("=")) return;
      formulaCells++;
      if (value === "" || value == null) {
        blanks++;
        return;
      }
      const match = typeof value === "string" ? ERROR_RE.exec(value) : null;
      if (!match) return;
      errorCells++;
      const type = `#${match[1]}`;
      const entry = byType.get(type) ?? { count: 0, cells: [] };
      entry.count++;
      if (entry.cells.length < SAMPLE_ADDRESSES) entry.cells.push(`${indexToColLetter(start.col + c)}${start.row + r}`);
      byType.set(type, entry);
    });
  });
  const errors = Array.from(byType, ([type, e]) => ({ type, count: e.count, cells: e.cells, ...(ERROR_HINTS[type] ? { hint: ERROR_HINTS[type] } : {}) }))
    .sort((a, b) => b.count - a.count);
  return {
    formula_cells: formulaCells,
    error_cells: errorCells,
    ...(errors.length ? { errors } : {}),
    blank_results: blanks,
    sample: values.slice(0, SAMPLE_ROWS).map((row) => row.slice(0, SAMPLE_COLS)),
  };
}

/** 附在工具结果里的一句话结论，让模型不必自己解读计数 */
export function formulaCheckNote(check: FormulaCheck): string {
  if (!check.formula_cells) return "";
  if (!check.error_cells) return `已自动检查：${check.formula_cells} 个公式全部算出结果，无需再读回抽查。`;
  const top = check.errors?.map((e) => `${e.type} ${e.count} 个`).join("、");
  return `已自动检查：${check.formula_cells} 个公式中有 ${check.error_cells} 个错误值（${top}）。请按 errors 里的地址和原因修正，不要在回复里说已全部成功。`;
}
