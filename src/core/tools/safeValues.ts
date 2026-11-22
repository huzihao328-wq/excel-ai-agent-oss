/**
 * Prevent Excel from interpreting ordinary text as a formula when it is written
 * through Range.values. The leading apostrophe is Excel's quote prefix: it is
 * not displayed as part of the cell value, but forces the remaining content to
 * stay text. Intentional formulas must use the dedicated set_formulas tool.
 */
export function protectFormulaLikeText(value: unknown): unknown {
  return typeof value === "string" && /^[=+\-@]/.test(value) ? `'${value}` : value;
}

export function protectFormulaLikeTextGrid<T extends unknown[][]>(grid: T): unknown[][] {
  return grid.map((row) => row.map(protectFormulaLikeText));
}

/**
 * Range.values parses strings like typed input, so text read from a cell
 * ("00123", "2024-03", "1/2", "50%", "TRUE") would come back as a number,
 * date or boolean. Quote-prefix only strings Excel could coerce.
 */
export function keepTextAsText(value: unknown): unknown {
  if (typeof value !== "string") return value;
  if (/^[=+\-@']/.test(value)) return `'${value}`;
  const s = value.trim();
  if (!s) return value;
  if (/^(true|false)$/i.test(s)) return `'${value}`;
  if (/\d/.test(s) && /^[\d\s.,:/%$¥￥€£()eE+\-年月日]+$/.test(s)) return `'${value}`;
  return value;
}

/** Build a Range.formulas assignment that restores formulas and literal text. */
export function buildRestorableGrid(values: unknown[][], formulas: unknown[][]): unknown[][] {
  return formulas.map((row, r) =>
    row.map((formula, c) => {
      const value = values[r]?.[c];
      const isFormula = typeof formula === "string" && formula.startsWith("=") && formula !== value;
      if (isFormula) return formula;
      // 快照里的字符串本来就是文本。恢复时全部按文本写回，避免编号、日期、
      // 布尔文本被重新解析，也保留内容本身的首个单引号。
      return typeof value === "string" && value !== "" ? `'${value}` : value;
    })
  );
}

export function gridsEqual(left: unknown[][], right: unknown[][]): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

export function isEmptyCell(value: unknown): boolean {
  return value === "" || value === null || value === undefined;
}

function excelErrorCode(value: unknown): string | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  if (typeof record.error === "string" && record.error) return record.error;
  if (record.type === "Error" && typeof record.errorCode === "string") return record.errorCode;
  return null;
}

export function cellValuesEqual(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (isEmptyCell(left) && isEmptyCell(right)) return true;
  if (typeof left === "number" && typeof right === "number") {
    return left === right || (Number.isNaN(left) && Number.isNaN(right));
  }
  const leftError = excelErrorCode(left);
  const rightError = excelErrorCode(right);
  if (leftError && rightError) return leftError === rightError;
  return false;
}

function isFormulaText(formula: unknown): boolean {
  return typeof formula === "string" && formula.startsWith("=");
}

const GENERAL_NUMBER_FORMATS = new Set(["", "General", "G/通用格式", "G/通用", "@"]);

export function normalizeNumberFormat(value: unknown): string {
  const text = String(value ?? "").trim();
  return GENERAL_NUMBER_FORMATS.has(text) ? "General" : text;
}

export function numberFormatGridsEqual(left: unknown[][] | undefined, right: unknown[][] | undefined): boolean {
  if (!left && !right) return true;
  if (!left || !right) return false;
  if (left.length !== right.length) return false;
  for (let r = 0; r < left.length; r++) {
    const leftRow = left[r] || [];
    const rightRow = right[r] || [];
    if (leftRow.length !== rightRow.length) return false;
    for (let c = 0; c < leftRow.length; c++) {
      if (normalizeNumberFormat(leftRow[c]) !== normalizeNumberFormat(rightRow[c])) return false;
    }
  }
  return true;
}

/**
 * Compare cell state while allowing calculated formula results to recalculate,
 * empty cells to be null or "", and Excel error objects to match by code.
 * Formula cells are identified by a leading "=", including snapshots taken
 * before Excel finished calculating (when values still equal the formula text).
 */
export function cellStatesEqual(
  currentValues: unknown[][],
  currentFormulas: unknown[][],
  expectedValues: unknown[][],
  expectedFormulas: unknown[][]
): boolean {
  if (currentFormulas.length !== expectedFormulas.length) return false;
  for (let r = 0; r < expectedFormulas.length; r++) {
    if ((currentFormulas[r]?.length ?? 0) !== (expectedFormulas[r]?.length ?? 0)) return false;
    for (let c = 0; c < expectedFormulas[r].length; c++) {
      const expectedFormula = expectedFormulas[r][c];
      const currentFormula = currentFormulas[r]?.[c];
      if (!cellValuesEqual(currentFormula, expectedFormula)) return false;
      if (isFormulaText(expectedFormula) || isFormulaText(currentFormula)) continue;
      if (!cellValuesEqual(currentValues[r]?.[c], expectedValues[r]?.[c])) return false;
    }
  }
  return true;
}
