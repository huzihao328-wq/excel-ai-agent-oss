/* global Excel, Office */
import { AppSettings, AuditIssue, ChoiceOption, ChoiceResolution, ConfirmAnswer, FollowupAction, ToolResult, UndoEntry } from "../types";
import { auditData as auditGridData, AuditRowInput } from "../audit";
import { gridPreview, pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { toolImportExternalSheet, toolListExternalFiles, toolReadExternalSheet } from "./externalFiles";
import { toolExportFile } from "./exportWorkbook";
import { toolForget, toolListMemories, toolRemember } from "../memory";
import { toolListSkills, toolSaveSkill } from "../skills";
import {
  toolApplyFilter,
  toolClearFilter,
  toolCreatePivot,
  toolCreateTable,
  toolListTables,
  uniqueSheetName,
} from "./tableTools";
import {
  toolCompareSheets,
  toolGroupSummary,
  toolReconcileTables,
  toolSplitByColumn,
  toolStackSheets,
  toolUnpivotColumns,
} from "./reshapeTools";
import { toolRunScript } from "./scriptTool";
import { toolCreateSheetIndex, toolSetPageLayout } from "./layoutTools";
import { fillBounds, fillRuns, FillMode, isFillableMerge, MergedArea, planBlankFill } from "./fillBlanks";
import { DATE_TARGETS, EXTRACT_TARGET_LABELS, ExtractTarget, extractValue, normalizeExtractTargets } from "./textExtract";
import { checkpointRestoreHint, ensureCheckpoint, isInternalSheet } from "./checkpoint";
import { TOOL_DEFINITIONS } from "./definitions";
import { normalizeToolArguments, parseToolArgumentsJson } from "../toolArguments";
import {
  buildRestorableGrid,
  cellStatesEqual,
  gridsEqual,
  isEmptyCell,
  keepTextAsText,
  normalizeNumberFormat,
  numberFormatGridsEqual,
  protectFormulaLikeTextGrid,
} from "./safeValues";
import { planRangeRead, SAMPLE_HINT } from "./rangeRead";
import { checkFormulaResults, formulaCheckNote } from "./formulaCheck";
import { missingSheetHint, splitSheetQualifiedAddresses } from "./argHints";
import {
  buildLookupFormulaGrid,
  colLetterToIndex,
  indexToColLetter,
  padSplitParts,
  resolveColSpec,
  resolveDelimiter,
  splitCellText,
  textJoinFormula,
} from "./formulaRefs";
import {
  applyCleanOps,
  guessColumnNumberFormat,
  inferDefaultYear,
  isRowBlank,
  normalizeCleanOps,
  toIsoDate,
} from "./cleanValues";

/** 工具执行上下文：由 UI 注入确认回调与设置 */
export interface ToolExecContext {
  settings: AppSettings;
  signal?: AbortSignal;
  /** 请求用户确认破坏性操作，返回是否同意 */
  confirm: (title: string, detail: string) => Promise<boolean>;
  /**
   * 界面提供时优先使用：带上发起确认的工具名（用于「本次任务同类操作不再询问」），
   * 并能拿到用户拒绝时附的说明。工具实现本身只调用 confirm。
   */
  confirmDetailed?: (title: string, detail: string, meta: { tool: string }) => Promise<ConfirmAnswer>;
  /** 请用户从若干方向中选一个 */
  choose: (title: string, question: string, options: ChoiceOption[]) => Promise<ChoiceResolution>;
}

// ---------------- 撤销管理 ----------------

const MAX_UNDO = 20;
const MAX_SNAPSHOT_CELLS = 50000;
let undoStack: UndoEntry[] = [];
let undoListeners: (() => void)[] = [];
let undoInProgress = false;

export function getUndoStack(): UndoEntry[] {
  return undoStack;
}

export function onUndoStackChange(fn: () => void): () => void {
  undoListeners.push(fn);
  return () => {
    undoListeners = undoListeners.filter((f) => f !== fn);
  };
}

function notifyUndo() {
  undoListeners.forEach((f) => f());
}

function pushUndo(entry: Omit<UndoEntry, "id" | "timestamp">): string | null {
  const cells = entry.formulas.length * (entry.formulas[0]?.length ?? 0);
  if (cells > MAX_SNAPSHOT_CELLS) return null; // 太大不做快照
  const id = `undo_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  undoStack.push({ ...entry, id, timestamp: Date.now() });
  if (undoStack.length > MAX_UNDO) undoStack.shift();
  notifyUndo();
  return id;
}

export function clearUndoStack(): void {
  if (undoStack.length === 0) return;
  undoStack = [];
  notifyUndo();
}

function formatUndoExcelError(err: unknown, entry: UndoEntry): Error {
  const anyErr = err as { message?: string; code?: string; debugInfo?: { code?: string; message?: string } };
  const msg = anyErr?.message || anyErr?.debugInfo?.message || String(err);
  const code = String(anyErr?.code || anyErr?.debugInfo?.code || "");
  const text = `${code} ${msg}`;
  if (/ItemNotFound/i.test(text)) {
    return new Error(`找不到工作表「${entry.sheet}」，可能已被重命名或删除。可用「恢复到任务开始」整段回退。`);
  }
  if (/InvalidArgument|InvalidSelection|UnsupportedOperation/i.test(text)) {
    return new Error(`无法恢复 ${entry.sheet}!${entry.address}：该区域已不存在或结构已变。可用「恢复到任务开始」整段回退。`);
  }
  return err instanceof Error ? err : new Error(msg);
}

async function applyUndoUnlocked(id: string): Promise<string> {
  const idx = undoStack.findIndex((e) => e.id === id);
  if (idx < 0) throw new Error("该操作已不在撤销记录中");
  const entry = undoStack[idx];
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(entry.sheet);
      sheet.load("isNullObject");
      await context.sync();
      if (sheet.isNullObject) {
        throw new Error(`找不到工作表「${entry.sheet}」，可能已被重命名或删除。可用「恢复到任务开始」整段回退。`);
      }
      const range = sheet.getRange(entry.address);
      range.load(entry.expectedNumberFormats ? "values,formulas,numberFormat" : "values,formulas");
      await context.sync();
      if (!cellStatesEqual(range.values as any[][], range.formulas as any[][], entry.expectedValues, entry.expectedFormulas)) {
        throw new Error(`撤销已停止：${entry.sheet}!${entry.address} 在该操作后又被修改。请先检查当前内容，避免覆盖后续编辑。`);
      }
      if (entry.expectedNumberFormats && !numberFormatGridsEqual(range.numberFormat as any[][], entry.expectedNumberFormats)) {
        throw new Error(`撤销已停止：${entry.sheet}!${entry.address} 的格式在该操作后又被修改。请先检查当前格式。`);
      }
      range.formulas = entry.formulas as any;
      if (entry.numberFormats) range.numberFormat = entry.numberFormats as any;
      await context.sync();
    });
  } catch (e) {
    throw formatUndoExcelError(e, entry);
  }
  undoStack.splice(idx, 1);
  notifyUndo();
  pushOpLog({
    description: `撤销：恢复 ${entry.sheet}!${entry.address} 的原内容`,
    kind: "undo",
    sheet: entry.sheet,
    address: entry.address,
    after: gridPreview(entry.formulas),
  });
  return `已恢复 ${entry.sheet}!${entry.address} 的原内容`;
}

export async function applyUndo(id: string): Promise<string> {
  if (undoInProgress) throw new Error("已有撤销操作正在执行，请稍候");
  undoInProgress = true;
  try {
    return await applyUndoUnlocked(id);
  } finally {
    undoInProgress = false;
  }
}

/**
 * 撤回到指定节点：先撤销它之后的操作，再撤销该节点本身。
 * UI 允许选择任意历史节点，因此不能直接对旧快照做单步恢复，否则后续写入
 * 会触发冲突校验，表现成“撤销失败”。
 */
export async function applyUndoTo(id: string): Promise<string> {
  if (undoInProgress) throw new Error("已有撤销操作正在执行，请稍候");
  const idx = undoStack.findIndex((e) => e.id === id);
  if (idx < 0) throw new Error("该操作已不在撤销记录中");
  const ids = undoStack.slice(idx).map((e) => e.id).reverse();
  undoInProgress = true;
  let completed = 0;
  try {
    for (const undoId of ids) {
      await applyUndoUnlocked(undoId);
      completed++;
    }
  } catch (e: any) {
    const detail = e?.message || String(e);
    if (completed > 0) {
      throw new Error(`已成功撤销最近 ${completed} 项，但继续撤回时停止：${detail}`);
    }
    throw e;
  } finally {
    undoInProgress = false;
  }
  return ids.length === 1 ? "已撤销最近一项操作" : `已撤回到所选节点，共撤销 ${ids.length} 项操作`;
}

function countNonEmpty(formulas: any[][]): number {
  let c = 0;
  for (const row of formulas) for (const v of row) if (v !== "" && v !== null && v !== undefined) c++;
  return c;
}

function countNonEmptyOutsideSource(grid: any[][], srcRows: number, srcCols: number): number {
  let n = 0;
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (r < srcRows && c < srcCols) continue;
      const v = row[c];
      if (v !== "" && v !== null && v !== undefined) n++;
    }
  }
  return n;
}

function previewGrid(grid: any[][], maxRows = 4, maxCols = 6): string {
  const rows = grid.slice(0, maxRows).map((r) =>
    r
      .slice(0, maxCols)
      .map((v) => (v === "" || v === null || v === undefined ? "·" : String(v).slice(0, 12)))
      .join(" | ")
  );
  const more = grid.length > maxRows ? `\n…共 ${grid.length} 行` : "";
  return rows.join("\n") + more;
}

// ---------------- 各工具实现 ----------------

async function getWorkbookOverview(): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/position,items/visibility");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();

    const infos: {
      sheet: Excel.Worksheet;
      used: Excel.Range;
      name: string;
      visibility: string;
    }[] = [];
    for (const ws of sheets.items) {
      if (isInternalSheet(ws.name)) continue;
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("address,rowCount,columnCount,isNullObject");
      ws.tables.load("items/name");
      infos.push({ sheet: ws, used, name: ws.name, visibility: ws.visibility });
    }
    await context.sync();

    // 第二轮：读取表头和样本行
    const sampleReqs: { name: string; range: Excel.Range | null; info: (typeof infos)[0] }[] = [];
    for (const info of infos) {
      if (info.used.isNullObject) {
        sampleReqs.push({ name: info.name, range: null, info });
        continue;
      }
      const rows = Math.min(info.used.rowCount, 4); // 表头 + 最多 3 行样本
      const cols = Math.min(info.used.columnCount, 30);
      const sample = info.used.getAbsoluteResizedRange(rows, cols);
      sample.load("values");
      sampleReqs.push({ name: info.name, range: sample, info });
    }
    await context.sync();

    const tablePending: { sheet: string; name: string; range: Excel.Range; columns: Excel.TableColumnCollection }[] = [];
    for (const info of infos) {
      for (const t of info.sheet.tables.items) {
        const tr = t.getRange();
        tr.load("address");
        t.columns.load("items/name");
        tablePending.push({ sheet: info.name, name: t.name, range: tr, columns: t.columns });
      }
    }
    await context.sync();

    const result = sampleReqs.map(({ name, range, info }) => {
      if (!range) return { name, empty: true, visibility: info.visibility };
      const values = range.values as any[][];
      return {
        name,
        visibility: info.visibility,
        used_range: info.used.address.split("!").pop(),
        row_count: info.used.rowCount,
        column_count: info.used.columnCount,
        headers_guess: values[0],
        sample_rows: values.slice(1),
      };
    });
    const tables = tablePending.map((t) => ({
      sheet: t.sheet,
      name: t.name,
      range: t.range.address.split("!").pop(),
      columns: t.columns.items.map((c) => c.name),
    }));

    return {
      ok: true,
      data: { active_sheet: active.name, sheets: result, tables },
      summary: `工作簿共 ${result.length} 个 Sheet${tables.length ? `、${tables.length} 个 Excel 表` : ""}`,
    };
  });
}

function loadRangePayload(range: Excel.Range, includeFormulas: boolean): void {
  range.load(includeFormulas ? "address,rowCount,columnCount,values,formulas" : "address,rowCount,columnCount,values");
}

function rangePayload(range: Excel.Range, includeFormulas: boolean): Record<string, unknown> {
  const data: Record<string, unknown> = {
    address: range.address,
    row_count: range.rowCount,
    column_count: range.columnCount,
    values: range.values,
  };
  if (includeFormulas) data.formulas = range.formulas;
  return data;
}

async function readRange(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    // 模型偶尔把布尔发成字符串 "false"，!! 会把它当成 true
    const includeFormulas = args.include_formulas === true || String(args.include_formulas).toLowerCase() === "true";
    range.load("address,rowCount,columnCount,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) {
      return { ok: true, data: { empty: true }, summary: `${args.sheet} 是空表` };
    }

    const plan = planRangeRead(range.rowCount, range.columnCount, {
      start_row: args.start_row,
      max_rows: args.max_rows,
    });
    const shortAddr = `${args.sheet}!${range.address.split("!").pop()}`;

    if (plan.mode === "full") {
      loadRangePayload(range, includeFormulas);
      await context.sync();
      return {
        ok: true,
        data: rangePayload(range, includeFormulas),
        summary: `读取 ${shortAddr}（${range.rowCount}×${range.columnCount}）`,
      };
    }

    if (plan.mode === "page") {
      const page = range.getCell(plan.startOffset, 0).getAbsoluteResizedRange(plan.rowCount, range.columnCount);
      loadRangePayload(page, includeFormulas);
      await context.sync();
      const data: Record<string, unknown> = {
        ...rangePayload(page, includeFormulas),
        full_address: range.address,
        full_row_count: range.rowCount,
        full_column_count: range.columnCount,
        start_row: plan.startOffset + 1,
        truncated: plan.truncated,
      };
      if (plan.truncated) {
        data.hint = `只返回了第 ${plan.startOffset + 1} 行起的 ${plan.rowCount} 行。继续用 start_row=${plan.startOffset + plan.rowCount + 1} 分页，或改用 column_stats / copy_range。`;
      }
      return {
        ok: true,
        data,
        summary: plan.truncated
          ? `分页读取 ${shortAddr}（第 ${plan.startOffset + 1} 行起 ${plan.rowCount} 行，表共 ${range.rowCount} 行）`
          : `读取 ${shortAddr}（第 ${plan.startOffset + 1} 行起 ${plan.rowCount} 行）`,
      };
    }

    const head = range.getCell(0, 0).getAbsoluteResizedRange(plan.headRows, range.columnCount);
    loadRangePayload(head, includeFormulas);
    const tail =
      plan.tailRows > 0
        ? range.getCell(range.rowCount - plan.tailRows, 0).getAbsoluteResizedRange(plan.tailRows, range.columnCount)
        : null;
    if (tail) loadRangePayload(tail, includeFormulas);
    await context.sync();
    const data: Record<string, unknown> = {
      address: range.address,
      row_count: range.rowCount,
      column_count: range.columnCount,
      truncated: true,
      sampled: true,
      hint: SAMPLE_HINT,
      values: head.values,
      tail_values: tail ? tail.values : [],
      tail_start_row: tail ? range.rowCount - plan.tailRows + 1 : null,
    };
    if (includeFormulas) {
      data.formulas = head.formulas;
      data.tail_formulas = tail ? tail.formulas : [];
    }
    return {
      ok: true,
      data,
      summary: `抽样读取 ${shortAddr}（共 ${range.rowCount}×${range.columnCount}，返回前 ${plan.headRows} 行+末 ${plan.tailRows} 行）`,
    };
  });
}

async function writeCells(args: any, ctx: ToolExecContext, asFormulas: boolean): Promise<ToolResult> {
  const grid: any[][] = asFormulas ? args.formulas : args.values;
  if (!Array.isArray(grid) || grid.length === 0 || !Array.isArray(grid[0])) {
    return { ok: false, error: "values/formulas 必须是非空二维数组" };
  }
  // 容错：模型偶尔会把某一行写成标量而不是数组
  const safeGrid = grid.map((r) => (Array.isArray(r) ? r : [r]));
  const rows = safeGrid.length;
  const cols = safeGrid.reduce((max, row) => Math.max(max, row.length), 0);
  if (!cols || rows * cols > MAX_SNAPSHOT_CELLS) {
    return { ok: false, error: `写入区域必须非空且单次不超过 ${MAX_SNAPSHOT_CELLS} 个单元格，请分批写入。` };
  }
  if (safeGrid.some((row) => row.some((cell) => cell != null && !["string", "number", "boolean"].includes(typeof cell)))) {
    return { ok: false, error: "单元格只能是文本、数字、布尔值或 null，不能写入对象或数组。" };
  }
  // 补齐参差不齐的行
  const normalized = safeGrid.map((r) => {
    const row = [...r];
    while (row.length < cols) row.push("");
    return row;
  });
  const writeGrid = asFormulas ? normalized : protectFormulaLikeTextGrid(normalized);

  // 先做快照并检测覆盖
  let previewFormulas: any[][] = [];
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let targetAddr = "";
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const target = sheet.getRange(args.start_cell).getAbsoluteResizedRange(rows, cols);
    target.load("address,formulas");
    await context.sync();
    previewFormulas = target.formulas as any[][];
    targetAddr = target.address.split("!").pop() as string;
  });

  const overwritten = countNonEmpty(previewFormulas);
  if (overwritten > 0 && ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `覆盖 ${args.sheet}!${targetAddr} 中 ${overwritten} 个已有单元格`,
      `原内容（节选）:\n${previewGrid(previewFormulas)}\n\n将写入:\n${previewGrid(normalized)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了本次覆盖写入。请不要重试，向用户询问意图。" };
  }

  let targetChanged = false;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const target = sheet.getRange(args.start_cell).getAbsoluteResizedRange(rows, cols);
    target.load("address,values,formulas");
    await context.sync();
    if (!gridsEqual(target.formulas as any[][], previewFormulas)) {
      targetChanged = true;
      return;
    }
    oldFormulas = buildRestorableGrid(target.values as any[][], target.formulas as any[][]);
    targetAddr = target.address.split("!").pop() as string;
    checkToolAbort(ctx);
    if (asFormulas) target.formulas = writeGrid;
    else target.values = writeGrid;
    await context.sync();
    target.load("values,formulas");
    await context.sync();
    expectedValues = target.values as any[][];
    expectedFormulas = target.formulas as any[][];
  });
  if (targetChanged) {
    return { ok: false, error: `${args.sheet}!${targetAddr} 在预览/确认后发生了变化，本次写入已停止。请重新读取后再操作。` };
  }

  const undoId = pushUndo({
    description: `${asFormulas ? "写入公式" : "写入数据"} ${args.sheet}!${targetAddr}`,
    sheet: args.sheet,
    address: targetAddr,
    formulas: oldFormulas,
    expectedValues,
    expectedFormulas,
  });

  pushOpLog({
    description: `${asFormulas ? "写入公式" : "写入数据"} ${args.sheet}!${targetAddr}（${rows}×${cols}）`,
    kind: "write",
    sheet: args.sheet,
    address: targetAddr,
    before: gridPreview(oldFormulas),
    after: gridPreview(normalized),
    ...(undoId ? { undoEntryId: undoId } : {}),
  });

  // 写入空区域 = 产出新内容（新表/汇总区），登记为可引用产出
  if (!asFormulas && overwritten === 0 && rows >= 2 && cols >= 2) {
    registerArtifact({
      name: args.sheet,
      kind: "range",
      sheet: args.sheet,
      address: targetAddr,
      brief: `写入 ${rows}×${cols} 数据`,
    });
  }

  const check = asFormulas ? checkFormulaResults(targetAddr, expectedValues, expectedFormulas) : null;
  return {
    ok: true,
    data: {
      written_range: `${args.sheet}!${targetAddr}`,
      rows,
      cols,
      overwritten_cells: overwritten,
      ...(check?.formula_cells ? { formula_check: check, note: formulaCheckNote(check) } : {}),
    },
    summary: `已写入 ${args.sheet}!${targetAddr}（${rows}×${cols}）`,
    undoable: !!undoId,
  };
}

async function createSheet(args: any): Promise<ToolResult> {
  if (isInternalSheet(String(args.name || ""))) {
    return { ok: false, error: "不能使用保留的检查点备份表名。" };
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.add(args.name);
    if (args.activate !== false) sheet.activate();
    await context.sync();
    return { ok: true, data: { created: args.name }, summary: `已创建 Sheet「${args.name}」` };
  });
}

async function deleteSheet(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  if (isInternalSheet(String(args.name || ""))) {
    return { ok: false, error: "不能删除任务检查点备份表。" };
  }
  if (ctx.settings.confirmDestructive) {
    const missing = await missingSheetError(args.name);
    if (missing) return missing;
    const accepted = await ctx.confirm(
      `删除 Sheet「${args.name}」`,
      checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)
    );
    if (!accepted) return { ok: false, error: "用户拒绝删除该 Sheet。请不要重试。" };
  }
  return Excel.run(async (context) => {
    context.workbook.worksheets.getItem(args.name).delete();
    await context.sync();
    pushOpLog({ description: `删除 Sheet「${args.name}」`, kind: "structure", sheet: args.name });
    return { ok: true, data: { deleted: args.name }, summary: `已删除 Sheet「${args.name}」` };
  });
}

async function clearRange(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const what = args.what || "contents";
  // 快照（确认预览）
  let previewFormulas: any[][] = [];
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let oldFormats: string[][] | undefined;
  let expectedNumberFormats: string[][] | undefined;
  let addr = "";
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(args.sheet).getRange(args.range);
    range.load(what === "all" ? "address,formulas,numberFormat" : "address,formulas");
    await context.sync();
    previewFormulas = range.formulas as any[][];
    addr = range.address.split("!").pop() as string;
    if (what === "all") oldFormats = range.numberFormat as string[][];
  });
  const nonEmpty = countNonEmpty(previewFormulas);
  if (nonEmpty > 0 && ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `清空 ${args.sheet}!${addr}（${nonEmpty} 个非空单元格）`,
      `原内容（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝清空该区域。请不要重试。" };
  }
  let targetChanged = false;
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(args.sheet).getRange(args.range);
    if (what !== "formats") {
      range.load(what === "all" ? "address,values,formulas,numberFormat" : "address,values,formulas");
      await context.sync();
      if (!gridsEqual(range.formulas as any[][], previewFormulas)) {
        targetChanged = true;
        return;
      }
      oldFormulas = buildRestorableGrid(range.values as any[][], range.formulas as any[][]);
      addr = range.address.split("!").pop() as string;
      if (what === "all") oldFormats = range.numberFormat as string[][];
    }
    const applyTo =
      what === "formats" ? Excel.ClearApplyTo.formats : what === "all" ? Excel.ClearApplyTo.all : Excel.ClearApplyTo.contents;
    checkToolAbort(ctx);
    range.clear(applyTo);
    await context.sync();
    if (what !== "formats") {
      range.load(what === "all" ? "values,formulas,numberFormat" : "values,formulas");
      await context.sync();
      expectedValues = range.values as any[][];
      expectedFormulas = range.formulas as any[][];
      if (what === "all") expectedNumberFormats = range.numberFormat as string[][];
    }
  });
  if (targetChanged) {
    return { ok: false, error: `${args.sheet}!${addr} 在预览/确认后发生了变化，本次清空已停止。请重新读取后再操作。` };
  }
  if (what !== "formats") {
    const undoId = pushUndo({
      description: `清空 ${args.sheet}!${addr}`,
      sheet: args.sheet,
      address: addr,
      formulas: oldFormulas,
      expectedValues,
      expectedFormulas,
      ...(oldFormats ? { numberFormats: oldFormats } : {}),
      ...(expectedNumberFormats ? { expectedNumberFormats } : {}),
    });
    pushOpLog({
      description: `清空${what === "all" ? "内容与格式" : "内容"} ${args.sheet}!${addr}`,
      kind: "write",
      sheet: args.sheet,
      address: addr,
      before: gridPreview(oldFormulas),
      after: "（已清空）",
      ...(undoId ? { undoEntryId: undoId } : {}),
    });
    return { ok: true, data: { cleared: `${args.sheet}!${addr}` }, summary: `已清空 ${args.sheet}!${addr}`, undoable: !!undoId };
  }
  return { ok: true, data: { cleared: `${args.sheet}!${addr}` }, summary: `已清空 ${args.sheet}!${addr}`, undoable: false };
}

/** 把模型传来的参数安全地转成正整数。模型偶尔把数字发成字符串，直接做算术会因字符串拼接算出灾难性的错误区域（如 "5"+2-1=51）。 */
function toPositiveInt(value: unknown, name: string): number {
  const n = Number(value);
  if (!Number.isInteger(n) || n < 1) throw new Error(`参数 ${name} 必须是 ≥1 的整数，收到: ${JSON.stringify(value)}`);
  return n;
}

function parseColLetter(value: unknown, name: string): string {
  const s = String(value ?? "").trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(s)) {
    throw new Error(`参数 ${name} 必须是列字母（如 A、B、AA），收到: ${JSON.stringify(value)}`);
  }
  return s;
}

const MAX_STRUCTURAL_INSERT = 200;

/**
 * 请用户确认之前先看一眼目标表在不在：不能让用户为一步注定失败的操作做决定。
 * 只在确定表不存在时提前返回；查不了（Excel 暂时不可用）就不拦，由后面的实际操作报错。
 */
async function missingSheetError(name: unknown): Promise<ToolResult | null> {
  const sheetName = String(name ?? "").trim();
  if (!sheetName) return null;
  try {
    const existing = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.getItemOrNullObject(sheetName);
      sheet.load("isNullObject");
      sheets.load("items/name");
      await context.sync();
      return sheet.isNullObject === true ? sheets.items.map((s) => s.name).filter((n) => !isInternalSheet(n)) : null;
    });
    if (!existing) return null;
    // 表名参数叫 sheet 的工具由 executeTool 统一补候选；这里把候选直接带上，参数叫 name 的 delete_sheet 也能拿到
    const hint = missingSheetHint({ sheet: sheetName }, existing);
    return { ok: false, error: `本次未执行。${hint || `工作表「${sheetName}」不存在。`}` };
  } catch {
    return null;
  }
}

async function insertRows(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const beforeRow = toPositiveInt(args.before_row, "before_row");
  const count = toPositiveInt(args.count ?? 1, "count");
  if (count > MAX_STRUCTURAL_INSERT) {
    throw new Error(`一次最多插入 ${MAX_STRUCTURAL_INSERT} 行，收到 ${count}`);
  }
  if (count > 10 && ctx.settings.confirmDestructive) {
    const missing = await missingSheetError(args.sheet);
    if (missing) return missing;
    const accepted = await ctx.confirm(
      `在 ${args.sheet} 第 ${beforeRow} 行前插入 ${count} 行`,
      `插入行会把后续数据下移。${checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝插入行。请不要重试。" };
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const addr = `${beforeRow}:${beforeRow + count - 1}`;
    sheet.getRange(addr).insert(Excel.InsertShiftDirection.down);
    await context.sync();
    pushOpLog({
      description: `在 ${args.sheet} 第 ${beforeRow} 行前插入 ${count} 行`,
      kind: "structure",
      sheet: args.sheet,
      address: addr,
    });
    return { ok: true, data: { inserted_rows: count, at: beforeRow }, summary: `在第 ${beforeRow} 行前插入 ${count} 行` };
  });
}

async function deleteRows(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const startRow = toPositiveInt(args.start_row, "start_row");
  const count = toPositiveInt(args.count ?? 1, "count");
  if (ctx.settings.confirmDestructive) {
    const missing = await missingSheetError(args.sheet);
    if (missing) return missing;
    const accepted = await ctx.confirm(
      `删除 ${args.sheet} 第 ${startRow} 起 ${count} 行`,
      `删除行会移动后续数据。${checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝删除行。请不要重试。" };
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    sheet.getRange(`${startRow}:${startRow + count - 1}`).delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    pushOpLog({
      description: `删除 ${args.sheet} 第 ${startRow} 起共 ${count} 行`,
      kind: "structure",
      sheet: args.sheet,
      address: `${startRow}:${startRow + count - 1}`,
    });
    return { ok: true, data: { deleted_rows: count }, summary: `已删除 ${count} 行` };
  });
}

async function insertColumns(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const count = toPositiveInt(args.count ?? 1, "count");
  const col = parseColLetter(args.before_column, "before_column");
  if (count > MAX_STRUCTURAL_INSERT) {
    throw new Error(`一次最多插入 ${MAX_STRUCTURAL_INSERT} 列，收到 ${count}`);
  }
  if (count > 10 && ctx.settings.confirmDestructive) {
    const missing = await missingSheetError(args.sheet);
    if (missing) return missing;
    const accepted = await ctx.confirm(
      `在 ${args.sheet} 的 ${col} 列前插入 ${count} 列`,
      `插入列会把后续数据右移。${checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝插入列。请不要重试。" };
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const start = colLetterToIndex(col);
    const addr = `${col}:${indexToColLetter(start + count - 1)}`;
    sheet.getRange(addr).insert(Excel.InsertShiftDirection.right);
    await context.sync();
    pushOpLog({
      description: `在 ${args.sheet} ${col} 列前插入 ${count} 列`,
      kind: "structure",
      sheet: args.sheet,
      address: addr,
    });
    return { ok: true, data: { inserted_columns: count }, summary: `已在 ${col} 列前插入 ${count} 列` };
  });
}

async function deleteColumns(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const count = toPositiveInt(args.count ?? 1, "count");
  const col = parseColLetter(args.start_column, "start_column");
  if (ctx.settings.confirmDestructive) {
    const missing = await missingSheetError(args.sheet);
    if (missing) return missing;
    const accepted = await ctx.confirm(
      `删除 ${args.sheet} 的 ${col} 起 ${count} 列`,
      checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)
    );
    if (!accepted) return { ok: false, error: "用户拒绝删除列。请不要重试。" };
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const start = colLetterToIndex(col);
    const addr = `${col}:${indexToColLetter(start + count - 1)}`;
    sheet.getRange(addr).delete(Excel.DeleteShiftDirection.left);
    await context.sync();
    pushOpLog({
      description: `删除 ${args.sheet} 的 ${col} 起共 ${count} 列`,
      kind: "structure",
      sheet: args.sheet,
      address: addr,
    });
    return { ok: true, data: { deleted_columns: count }, summary: `已删除 ${count} 列` };
  });
}

async function formatRange(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(args.sheet).getRange(args.range);
    if (args.merge === true) {
      // formulas 也包含常量，因此返回空字符串的公式仍算已占用。
      range.load("formulas");
      await context.sync();
      if (range.formulas.some((row, r) => row.some((cell, c) => (r !== 0 || c !== 0) && !isEmptyCell(cell)))) {
        throw new Error("左上角以外有非空单元格或公式，合并会丢失内容。请先移走这些内容再合并。");
      }
    }
    const f = range.format;
    if (args.bold !== undefined) f.font.bold = args.bold;
    if (args.italic !== undefined) f.font.italic = args.italic;
    if (args.font_size) f.font.size = args.font_size;
    if (args.font_color) f.font.color = args.font_color;
    if (args.fill_color) f.fill.color = args.fill_color;
    if (args.number_format) {
      range.load("rowCount,columnCount");
      await context.sync();
      const nf = Array.from({ length: range.rowCount }, () => Array(range.columnCount).fill(args.number_format));
      range.numberFormat = nf;
    }
    if (args.horizontal_alignment) f.horizontalAlignment = args.horizontal_alignment;
    if (args.border) {
      const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const;
      for (const edge of edges) {
        const b = f.borders.getItem(edge);
        b.style = Excel.BorderLineStyle.continuous;
        b.color = "#B0B7BD";
      }
    }
    if (args.vertical_alignment) f.verticalAlignment = args.vertical_alignment;
    if (args.wrap_text !== undefined) f.wrapText = args.wrap_text;
    if (args.merge === true) range.merge(false);
    else if (args.merge === false) range.unmerge();
    if (args.column_width != null) {
      const chars = Number(args.column_width);
      if (!Number.isFinite(chars) || chars < 0 || chars > 255) throw new Error("column_width 是 Excel 列宽字符数，取 0～255");
      // Office.js 用磅；Excel 默认字体下 1 个字符宽约 5.7 磅（8.43 字符 = 48 磅）
      f.columnWidth = Math.round(chars * 5.7 * 10) / 10;
    }
    if (args.row_height != null) {
      const pt = Number(args.row_height);
      if (!Number.isFinite(pt) || pt < 0 || pt > 409) throw new Error("row_height 以磅为单位，取 0～409");
      f.rowHeight = pt;
    }
    if (args.autofit_columns) f.autofitColumns();
    if (args.autofit_rows) f.autofitRows();
    await context.sync();
    pushOpLog({ description: `设置格式 ${args.sheet}!${args.range}`, kind: "format", sheet: args.sheet, address: args.range });
    return { ok: true, data: { formatted: `${args.sheet}!${args.range}` }, summary: `已设置 ${args.sheet}!${args.range} 格式` };
  });
}

const VALIDATION_OPERATORS: Record<string, string> = {
  between: "Between", not_between: "NotBetween", equal: "EqualTo", not_equal: "NotEqualTo",
  greater_than: "GreaterThan", less_than: "LessThan", greater_or_equal: "GreaterThanOrEqualTo", less_or_equal: "LessThanOrEqualTo",
};
const TWO_BOUND_OPERATORS = new Set(["between", "not_between"]);
/** Excel 对直接写在规则里的下拉选项有 255 字符上限 */
const MAX_LIST_SOURCE_CHARS = 255;

function validationBound(value: unknown, type: string, label: string): string | number {
  if (value === undefined || value === null || value === "") throw new Error(`缺少 ${label}`);
  if (typeof value === "number") return value;
  const text = String(value).trim();
  if (text.startsWith("=")) return text;
  if (type === "date") {
    if (!/^\d{4}-\d{1,2}-\d{1,2}$/.test(text)) throw new Error(`${label} 日期请写成 yyyy-mm-dd，或以 = 开头的公式/单元格引用`);
    return text;
  }
  const n = Number(text);
  if (!Number.isFinite(n)) throw new Error(`${label} 必须是数字，或以 = 开头的公式/单元格引用`);
  return n;
}

async function setDataValidation(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "");
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能修改检查点备份表。" };
  const type = String(args.type || "");
  let rule: Excel.DataValidationRule | null = null;
  let describe = "";
  if (type === "list") {
    const source = typeof args.source === "string" ? args.source.trim() : "";
    const items: string[] = Array.isArray(args.items) ? args.items.map((v: unknown) => String(v ?? "").trim()).filter(Boolean) : [];
    if (!source && !items.length) return { ok: false, error: "type=list 需要 items（选项数组）或 source（如 =参数!$A$2:$A$20）" };
    let listSource: string;
    if (source) {
      listSource = source.startsWith("=") ? source : `=${source}`;
      describe = `下拉列表（来源 ${listSource}）`;
    } else {
      const bad = items.find((item) => item.includes(","));
      if (bad) return { ok: false, error: `选项「${bad}」含英文逗号，不能直接写在规则里。请先把选项写到一列，再用 source 引用那一列。` };
      listSource = Array.from(new Set(items)).join(",");
      if (listSource.length > MAX_LIST_SOURCE_CHARS) {
        return { ok: false, error: `选项合计超过 Excel 的 ${MAX_LIST_SOURCE_CHARS} 字符上限。请先把选项写到一列（可新建「参数」表），再用 source 引用。` };
      }
      describe = `下拉列表（${items.length} 项）`;
    }
    rule = { list: { inCellDropDown: args.in_cell_dropdown !== false, source: listSource } };
  } else if (type === "whole_number" || type === "decimal" || type === "date" || type === "text_length") {
    const opKey = String(args.operator || "between");
    const operator = VALIDATION_OPERATORS[opKey];
    if (!operator) return { ok: false, error: `operator 只能是 ${Object.keys(VALIDATION_OPERATORS).join("、")}` };
    let basic: Excel.BasicDataValidation;
    try {
      const formula1 = validationBound(args.min ?? args.value, type, TWO_BOUND_OPERATORS.has(opKey) ? "min" : "value");
      basic = { operator: operator as Excel.DataValidationOperator, formula1 };
      if (TWO_BOUND_OPERATORS.has(opKey)) basic.formula2 = validationBound(args.max, type, "max");
    } catch (e: any) {
      return { ok: false, error: e.message };
    }
    const key = type === "whole_number" ? "wholeNumber" : type === "text_length" ? "textLength" : type;
    rule = { [key]: basic } as Excel.DataValidationRule;
    describe = `${{ whole_number: "整数", decimal: "数字", date: "日期", text_length: "文本长度" }[type]}限制（${opKey}）`;
  } else if (type === "custom") {
    const formula = typeof args.formula === "string" ? args.formula.trim() : "";
    if (!formula.startsWith("=")) return { ok: false, error: "type=custom 需要以 = 开头的 formula，如 =COUNTIF($A:$A,A2)=1" };
    rule = { custom: { formula } };
    describe = "自定义公式校验";
  } else if (type !== "clear") {
    return { ok: false, error: "type 只能是 list、whole_number、decimal、date、text_length、custom、clear" };
  }

  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(args.range);
    const dv = range.dataValidation;
    dv.clear();
    if (rule) {
      dv.rule = rule;
      dv.ignoreBlanks = args.allow_blank !== false;
      if (args.error_message) {
        dv.errorAlert = {
          showAlert: true,
          style: args.reject_invalid === false ? Excel.DataValidationAlertStyle.warning : Excel.DataValidationAlertStyle.stop,
          title: "输入无效",
          message: String(args.error_message).slice(0, 225),
        };
      } else if (args.reject_invalid === false) {
        dv.errorAlert = { showAlert: true, style: Excel.DataValidationAlertStyle.warning, title: "输入无效", message: "输入的值不符合规则。" };
      }
      if (args.prompt_message) {
        dv.prompt = { showPrompt: true, title: "提示", message: String(args.prompt_message).slice(0, 255) };
      }
    }
    range.load("address");
    await context.sync();
    const address = range.address.split("!").pop() as string;
    const text = rule ? `为 ${sheetName}!${address} 设置${describe}` : `清除 ${sheetName}!${address} 的数据验证`;
    pushOpLog({ description: text, kind: "format", sheet: sheetName, address });
    return {
      ok: true,
      data: { range: `${sheetName}!${address}`, type, rule: describe || "cleared", note: rule ? "只约束今后的输入，已有的不合规值不会被改动或标出。" : undefined },
      summary: `已${text}`,
    };
  });
}

async function setVisibility(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "");
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能修改检查点备份表。" };
  const hidden = args.hidden !== false;
  const target = String(args.target || "");
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    if (target === "sheet") {
      sheets.load("items/name,items/visibility");
      await context.sync();
      const ws = sheets.items.find((s) => s.name === sheetName);
      if (!ws) throw new Error(`找不到工作表「${sheetName}」`);
      if (hidden) {
        const visible = sheets.items.filter((s) => s.visibility === Excel.SheetVisibility.visible && !isInternalSheet(s.name));
        if (visible.length <= 1 && visible[0]?.name === sheetName) return { ok: false, error: "这是唯一可见的工作表，Excel 不允许把它隐藏。" };
        if (ws.visibility === Excel.SheetVisibility.visible) {
          const active = sheets.getActiveWorksheet();
          active.load("name");
          await context.sync();
          if (active.name === sheetName) {
            const other = visible.find((s) => s.name !== sheetName);
            other?.activate();
          }
        }
      }
      ws.visibility = hidden ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;
      await context.sync();
      const text = `${hidden ? "隐藏" : "取消隐藏"}工作表「${sheetName}」`;
      pushOpLog({ description: text, kind: "structure", sheet: sheetName });
      return { ok: true, data: { sheet: sheetName, hidden }, summary: `已${text}` };
    }
    if (target !== "rows" && target !== "columns") return { ok: false, error: "target 只能是 rows、columns 或 sheet" };
    const spec = String(args.range || "").trim();
    const ok = target === "rows" ? /^\d+(:\d+)?$/.test(spec) : /^[A-Za-z]{1,3}(:[A-Za-z]{1,3})?$/.test(spec);
    if (!ok) return { ok: false, error: target === "rows" ? "rows 的 range 写成行号，如 3 或 3:10" : "columns 的 range 写成列字母，如 C 或 C:F" };
    const full = spec.includes(":") ? spec : `${spec}:${spec}`;
    const range = sheets.getItem(sheetName).getRange(full.toUpperCase());
    if (target === "rows") range.rowHidden = hidden;
    else range.columnHidden = hidden;
    await context.sync();
    const text = `${hidden ? "隐藏" : "取消隐藏"}「${sheetName}」的${target === "rows" ? "行" : "列"} ${full.toUpperCase()}`;
    pushOpLog({ description: text, kind: "format", sheet: sheetName, address: full.toUpperCase() });
    return { ok: true, data: { sheet: sheetName, target, range: full.toUpperCase(), hidden }, summary: `已${text}` };
  });
}

async function sortRange(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "sort_range 需要 sheet" };
  const hasHeaders = args.has_headers !== false;
  const keySpecs: { column: unknown; ascending?: boolean }[] =
    Array.isArray(args.keys) && args.keys.length
      ? args.keys.map((k: any) =>
          k && typeof k === "object" ? { column: k.column ?? k.key ?? k.key_column, ascending: k.ascending } : { column: k }
        )
      : [{ column: args.key_column ?? args.column, ascending: args.ascending }];
  if (keySpecs.some((k) => k.column == null || k.column === "")) {
    return { ok: false, error: "sort_range 需要 key_column（表头名、列字母或区域内 0 基列号），或 keys 数组" };
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,rowCount,columnCount,columnIndex,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) return { ok: false, error: `${sheetName} 是空表，无法排序` };
    const headerRange = range.getCell(0, 0).getAbsoluteResizedRange(1, range.columnCount);
    headerRange.load("values");
    await context.sync();
    const headers = ((headerRange.values as any[][])[0] || []).map((v) => String(v ?? "").trim());
    const fields = keySpecs.map((spec, i) => ({
      key: resolveColSpec(headers, spec.column, range.columnIndex, i === 0 ? "key_column" : `keys[${i}]`),
      ascending: spec.ascending !== false,
    }));
    range.sort.apply(fields, false, hasHeaders);
    await context.sync();
    const addr = range.address.split("!").pop() as string;
    const labels = fields.map((f) => headers[f.key] || `第${f.key + 1}列`);
    pushOpLog({
      description: `排序 ${sheetName}!${addr}（${labels.join(" → ")}）`,
      kind: "format",
      sheet: sheetName,
      address: addr,
    });
    return {
      ok: true,
      data: { sorted: `${sheetName}!${addr}`, keys: labels, ascending: fields.map((f) => f.ascending) },
      summary: `已按「${labels.join(" → ")}」排序 ${sheetName}!${addr}`,
    };
  });
}

const ERROR_RE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NULL!|NUM!|GETTING_DATA|SPILL!)/i;

export function analyzeGrid(values: any[][], formulas: any[][]): {
  hasFormula: boolean;
  hasError: boolean;
  formulaPreview?: string;
  errors: string[];
} {
  let hasFormula = false;
  const errors: string[] = [];
  let formulaPreview: string | undefined;
  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < (formulas[r]?.length ?? 0); c++) {
      const f = formulas[r][c];
      if (typeof f === "string" && f.startsWith("=")) {
        hasFormula = true;
        if (!formulaPreview) formulaPreview = f.length > 80 ? `${f.slice(0, 80)}…` : f;
      }
      const v = values[r]?.[c];
      if (typeof v === "string" && ERROR_RE.test(v) && !errors.includes(v)) errors.push(v);
    }
  }
  return { hasFormula, hasError: errors.length > 0, formulaPreview, errors };
}

async function inspectRange(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = args.sheet
      ? context.workbook.worksheets.getItem(args.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const range = args.range ? sheet.getRange(args.range) : context.workbook.getSelectedRange();
    range.load("address,rowCount,columnCount,values,formulas,numberFormat,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) {
      return { ok: true, data: { empty: true }, summary: "选区为空" };
    }
    const cells = range.rowCount * range.columnCount;
    if (cells > 400) {
      return {
        ok: false,
        error: `区域 ${range.address} 有 ${cells} 个单元格，inspect_range 上限 400。请缩小范围。`,
      };
    }
    const values = range.values as any[][];
    const formulas = range.formulas as any[][];
    const analysis = analyzeGrid(values, formulas);
    const cellsOut: { address: string; value: unknown; formula?: string; error?: string }[] = [];
    const start = range.address.split("!").pop() || "";
    // 只展开前 40 个格子的明细，避免刷屏
    let n = 0;
    for (let r = 0; r < formulas.length && n < 40; r++) {
      for (let c = 0; c < (formulas[r]?.length ?? 0) && n < 40; c++) {
        const f = formulas[r][c];
        const v = values[r]?.[c];
        const isFormula = typeof f === "string" && f.startsWith("=");
        const err = typeof v === "string" && ERROR_RE.test(v) ? v : undefined;
        if (isFormula || err || args.include_all_cells) {
          cellsOut.push({
            address: `${start}[${r},${c}]`,
            value: v,
            ...(isFormula ? { formula: f } : {}),
            ...(err ? { error: err } : {}),
          });
          n++;
        }
      }
    }
    return {
      ok: true,
      data: {
        sheet: sheet.name,
        address: range.address,
        row_count: range.rowCount,
        column_count: range.columnCount,
        has_formula: analysis.hasFormula,
        errors: analysis.errors,
        values,
        formulas,
        notable_cells: cellsOut,
      },
      summary: analysis.hasError
        ? `${sheet.name}!${range.address.split("!").pop()} 含错误 ${analysis.errors.join("、")}`
        : analysis.hasFormula
          ? `${sheet.name}!${range.address.split("!").pop()} 含公式`
          : `已查看 ${sheet.name}!${range.address.split("!").pop()}`,
    };
  });
}

function normalizeFollowups(raw: unknown): FollowupAction[] {
  if (!Array.isArray(raw)) return [];
  const out: FollowupAction[] = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") continue;
    const rec = item as Record<string, unknown>;
    const label = String(rec.label ?? rec.title ?? "").trim();
    const prompt = String(rec.prompt ?? rec.text ?? "").trim();
    if (!label || !prompt) continue;
    out.push({ label: label.slice(0, 16), prompt });
    if (out.length >= 3) break;
  }
  return out;
}

function suggestFollowups(args: any): ToolResult {
  const actions = normalizeFollowups(args.actions);
  if (actions.length < 2) {
    return { ok: false, error: "suggest_followups 需要 2～3 个 actions（label + prompt）" };
  }
  return {
    ok: true,
    data: { actions },
    followups: actions,
    summary: `建议下一步：${actions.map((a) => a.label).join("、")}`,
  };
}

async function findCells(args: any): Promise<ToolResult> {
  const text = String(args.text ?? "").trim();
  if (!text) return { ok: false, error: "find_cells 需要非空的 text 参数" };
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const targets = args.sheet
      ? sheets.items.filter((s) => s.name === args.sheet && !isInternalSheet(s.name))
      : sheets.items.filter((s) => !isInternalSheet(s.name));
    const searches = targets.map((ws) => {
      const found = ws.findAllOrNullObject(args.text, {
        completeMatch: !!args.complete_match,
        matchCase: !!args.match_case,
      });
      found.load("address");
      return { name: ws.name, found };
    });
    await context.sync();
    const matches: { sheet: string; addresses: string[] }[] = [];
    let total = 0;
    for (const s of searches) {
      if (!s.found.isNullObject && s.found.address) {
        const addrs = s.found.address.split(",").map((a) => a.split("!").pop() as string);
        total += addrs.length;
        matches.push({ sheet: s.name, addresses: addrs.slice(0, 100) });
      }
    }
    return {
      ok: true,
      data: { total_matches: total, matches },
      summary: total ? `找到 ${total} 处「${args.text}」` : `未找到「${args.text}」`,
    };
  });
}

async function createChart(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const dataRange = sheet.getRange(args.data_range);
    const chart = sheet.charts.add(args.chart_type as Excel.ChartType, dataRange, Excel.ChartSeriesBy.auto);
    if (args.title) chart.title.text = args.title;
    if (args.position_start_cell) chart.setPosition(args.position_start_cell);
    await context.sync();
    const chartName = args.title || `图表@${args.sheet}`;
    pushOpLog({
      description: `在 ${args.sheet} 创建图表「${chartName}」`,
      kind: "chart",
      sheet: args.sheet,
      ...(args.position_start_cell ? { address: args.position_start_cell } : {}),
    });
    registerArtifact({
      name: chartName,
      kind: "chart",
      sheet: args.sheet,
      ...(args.position_start_cell ? { address: args.position_start_cell } : {}),
      brief: `图表（${args.data_range}，${args.chart_type}）`,
    });
    return { ok: true, data: { chart_created: true }, summary: `已在 ${args.sheet} 创建${args.title ? `「${args.title}」` : ""}图表` };
  });
}

// ---------------- 大区域扫描（体检 / 公式诊断共用） ----------------

const MAX_SCAN_CELLS = 50000;

interface GridWindow {
  /** 窗口首行在目标区域内的 0 基偏移 */
  startRow: number;
  values: any[][];
  formulas?: any[][];
}

/**
 * 读取目标区域（rowStart/colStart 为 0 基绝对坐标）的 values（可选 formulas）。
 * 总量不超过 MAX_SCAN_CELLS 时整读；否则读头部 60% + 中段均匀 5 块 + 尾部。
 */
async function readWindows(
  context: Excel.RequestContext,
  ws: Excel.Worksheet,
  rowStart: number,
  colStart: number,
  totalRows: number,
  cols: number,
  includeFormulas: boolean
): Promise<GridWindow[]> {
  const props = includeFormulas ? "values,formulas" : "values";
  const makeRange = (r0: number, n: number) => {
    const rg = ws.getRangeByIndexes(rowStart + r0, colStart, n, cols);
    rg.load(props);
    return rg;
  };
  if (totalRows * cols <= MAX_SCAN_CELLS) {
    const rg = makeRange(0, totalRows);
    await context.sync();
    return [
      { startRow: 0, values: rg.values as any[][], ...(includeFormulas ? { formulas: rg.formulas as any[][] } : {}) },
    ];
  }
  const budgetRows = Math.max(1000, Math.floor(MAX_SCAN_CELLS / cols));
  const headRows = Math.min(totalRows, Math.floor(budgetRows * 0.6));
  const tailRows = Math.min(Math.max(totalRows - headRows, 0), 200);
  const blockRows = Math.floor(Math.max(0, budgetRows - headRows - tailRows) / 5);
  const ranges: { start: number; range: Excel.Range }[] = [{ start: 0, range: makeRange(0, headRows) }];
  if (tailRows > 0) ranges.push({ start: totalRows - tailRows, range: makeRange(totalRows - tailRows, tailRows) });
  if (blockRows >= 50) {
    const gapStart = headRows;
    const gapEnd = totalRows - tailRows;
    const gap = gapEnd - gapStart;
    const maxBlocks = Math.min(5, Math.max(0, Math.floor(gap / blockRows)));
    if (maxBlocks === 1) {
      const start = gapStart + Math.floor((gap - blockRows) / 2);
      ranges.push({ start, range: makeRange(start, blockRows) });
    } else if (maxBlocks > 1) {
      const span = gap - blockRows;
      for (let b = 0; b < maxBlocks; b++) {
        const start = gapStart + Math.floor((span * b) / (maxBlocks - 1));
        ranges.push({ start, range: makeRange(start, blockRows) });
      }
    }
  }
  await context.sync();
  return ranges.map(({ start, range }) => ({
    startRow: start,
    values: range.values as any[][],
    ...(includeFormulas ? { formulas: range.formulas as any[][] } : {}),
  }));
}

const AUDIT_TYPE_LABELS: Record<AuditIssue["type"], string> = {
  missing: "缺失值",
  duplicate: "重复行",
  type: "类型不一致",
  outlier: "离群值",
  whitespace: "首尾空格",
};

async function toolAuditData(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const ws = args.sheet
      ? context.workbook.worksheets.getItem(args.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = args.range ? ws.getRange(args.range) : ws.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,isNullObject,rowIndex,columnIndex");
    await context.sync();
    if ((used as any).isNullObject) {
      return { ok: false, error: `${ws.name} 是空表，没有可体检的数据` };
    }
    const totalRows = used.rowCount;
    const cols = used.columnCount;
    const rowStart = used.rowIndex;
    const colStart = used.columnIndex;
    const firstAbs = rowStart + 1;
    const lastAbs = rowStart + totalRows;
    const requestedHeader = Number(args.header_row);
    const headerRow =
      Number.isFinite(requestedHeader) && requestedHeader >= 1
        ? Math.min(Math.max(Math.trunc(requestedHeader), firstAbs), lastAbs)
        : firstAbs;

    const headerRange = ws.getRangeByIndexes(headerRow - 1, colStart, 1, cols);
    headerRange.load("values");
    await context.sync();
    const headers: any[] = [...((headerRange.values as any[][])[0] || [])];

    const windows = await readWindows(context, ws, rowStart, colStart, totalRows, cols, false);
    const dataRows: AuditRowInput[] = [];
    const seenRows = new Set<number>();
    let sampledRows = 0;
    for (const w of windows) {
      w.values.forEach((cells, i) => {
        const absRow = rowStart + w.startRow + i + 1; // 1 基绝对行号
        if (absRow === headerRow || seenRows.has(absRow)) return;
        seenRows.add(absRow);
        sampledRows++;
        if (absRow > headerRow) dataRows.push({ row: absRow, cells });
      });
    }
    const totalDataRows = Math.max(0, lastAbs - headerRow);
    const audit = auditGridData(headers, dataRows, { sampledRows, totalDataRows, colStart });
    const addr = used.address.split("!").pop() as string;
    const high = audit.issues.filter((i) => i.severity === "high").length;
    const medium = audit.issues.filter((i) => i.severity === "medium").length;
    const low = audit.issues.length - high - medium;

    // 写报告 Sheet
    const reportName = await uniqueSheetName(context, "数据体检报告");
    const reportWs = context.workbook.worksheets.add(reportName);
    const now = new Date();
    const pad2 = (n: number) => String(n).padStart(2, "0");
    const time = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())} ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
    const sevLabel = (s: string) => (s === "high" ? "高" : s === "medium" ? "中" : "低");
    const padRow = (r: any[]) => {
      const row = [...r];
      while (row.length < 7) row.push("");
      return row;
    };
    const reportRows = [
      padRow(["数据体检报告"]),
      padRow(["检查范围", `${ws.name}!${addr}`]),
      padRow([
        "数据规模",
        `${totalDataRows} 行 × ${cols} 列（不含表头${audit.stats.sampled ? `，抽样 ${dataRows.length} 行分析` : ""}）`,
      ]),
      padRow(["发现问题", `${audit.issues.length} 项（高 ${high} / 中 ${medium} / 低 ${low}）`]),
      padRow(["生成时间", time]),
      padRow([]),
      padRow(["编号", "类型", "严重度", "位置", "数量", "示例", "修复建议"]),
      ...audit.issues.map((i) =>
        padRow([i.id, AUDIT_TYPE_LABELS[i.type], sevLabel(i.severity), i.column ? `「${i.column}」列` : "全表", i.count, i.sample, i.suggestion])
      ),
    ];
    reportWs.getRangeByIndexes(0, 0, reportRows.length, 7).values = reportRows;
    const title = reportWs.getRange("A1");
    title.format.font.bold = true;
    title.format.font.size = 14;
    const infoCol = reportWs.getRange("A2:A5");
    infoCol.format.font.bold = true;
    const tableHeader = reportWs.getRange("A7:G7");
    tableHeader.format.font.bold = true;
    tableHeader.format.font.color = "#FFFFFF";
    tableHeader.format.fill.color = "#1A8C4E";
    if (audit.issues.length > 0) {
      const tableBody = reportWs.getRange(`A8:G${7 + audit.issues.length}`);
      for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const) {
        const b = tableBody.format.borders.getItem(edge);
        b.style = Excel.BorderLineStyle.continuous;
        b.color = "#B0B7BD";
      }
    }
    reportWs.getRange("A1:G7").format.autofitColumns();
    reportWs.activate();
    await context.sync();

    pushOpLog({
      description: `生成数据体检报告（${ws.name}!${addr}，${audit.issues.length} 个问题）`,
      kind: "other",
      sheet: reportName,
      address: `A1:G${7 + audit.issues.length}`,
    });
    registerArtifact({
      name: reportName,
      kind: "sheet",
      sheet: reportName,
      address: `A1:G${7 + audit.issues.length}`,
      brief: `数据体检报告：${ws.name} ${audit.issues.length} 个问题（高${high}/中${medium}/低${low}）`,
    });

    return {
      ok: true,
      data: {
        stats: audit.stats,
        issues: audit.issues,
        report_sheet: reportName,
        hint: "报告已写入新 Sheet，用户界面会出现可勾选的问题卡片。现在只在回复里总结问题（按严重度），等用户勾选并点「修复选中项」后再动手修原表，不要现在就改。",
      },
      summary: `体检完成：${audit.issues.length} 个问题（高 ${high} / 中 ${medium} / 低 ${low}），报告已写入「${reportName}」`,
      audit: { sheet: ws.name, range: addr, reportSheet: reportName, issues: audit.issues },
      ...(audit.issues.length
        ? {
            followups: [
              ...(high
                ? [
                    {
                      label: "修复高严重度",
                      prompt: `请修复数据体检（${ws.name}!${addr}）中严重度为 high 的 ${high} 个问题。明细见上方体检结果，按每条的 suggestion/detail 定位处理，完成后逐项汇报。`,
                    },
                  ]
                : []),
              {
                label: "清洗到新 Sheet",
                prompt: `根据刚才的体检结果，把 ${ws.name} 的数据清洗后写入新 Sheet：去除重复行、纠正类型、去掉首尾空格，缺失值和离群值按建议处理，并保留原表不动。`,
              },
            ],
          }
        : {}),
    };
  });
}

async function toolDiagnoseErrors(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const ws = args.sheet
      ? context.workbook.worksheets.getItem(args.sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const used = args.range ? ws.getRange(args.range) : ws.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,isNullObject,rowIndex,columnIndex");
    await context.sync();
    if ((used as any).isNullObject) {
      return { ok: false, error: `${ws.name} 是空表` };
    }
    const totalRows = used.rowCount;
    const cols = used.columnCount;
    const rowStart = used.rowIndex;
    const colStart = used.columnIndex;
    const windows = await readWindows(context, ws, rowStart, colStart, totalRows, cols, true);

    interface Cluster {
      errorType: string;
      count: number;
      addresses: string[];
      formulas: string[];
      pattern: string;
    }
    const clusters = new Map<string, Cluster>();
    let totalErrors = 0;
    for (const w of windows) {
      w.values.forEach((row, i) => {
        const fRow = w.formulas?.[i] ?? [];
        for (let c = 0; c < row.length; c++) {
          const v = row[c];
          if (typeof v !== "string" || !ERROR_RE.test(v)) continue;
          totalErrors++;
          const formula = typeof fRow[c] === "string" && (fRow[c] as string).startsWith("=") ? (fRow[c] as string) : "";
          // 公式模式：把单元格引用/数字/字符串字面量换成占位符，剩下的骨架相同的算一簇
          const pattern = formula
            .replace(/\$?[A-Z]{1,3}\$?\d+/gi, "R")
            .replace(/\d+(\.\d+)?/g, "N")
            .replace(/"[^"]*"/g, "S");
          const key = `${v}|${pattern}`;
          const cl = clusters.get(key) || { errorType: v, count: 0, addresses: [], formulas: [], pattern };
          cl.count++;
          if (cl.addresses.length < 6) {
            cl.addresses.push(`${indexToColLetter(colStart + c)}${rowStart + w.startRow + i + 1}`);
          }
          if (formula && cl.formulas.length < 2) cl.formulas.push(formula);
          clusters.set(key, cl);
        }
      });
    }
    const addr = used.address.split("!").pop() as string;
    const sampled = totalRows * cols > MAX_SCAN_CELLS;
    const scannedCells = windows.reduce((n, w) => n + w.values.length * cols, 0);
    if (totalErrors === 0) {
      return {
        ok: true,
        data: { sheet: ws.name, scanned_range: addr, sampled, scanned_cells: scannedCells, total_errors: 0 },
        summary: sampled
          ? `${ws.name}!${addr} 抽样扫描未发现公式错误（未扫完全表）`
          : `${ws.name}!${addr} 未发现公式错误`,
      };
    }
    const clusterList = [...clusters.values()]
      .sort((a, b) => b.count - a.count)
      .slice(0, 12)
      .map((cl) => ({
        error: cl.errorType,
        count: cl.count,
        example_addresses: cl.addresses,
        example_formulas: cl.formulas,
        formula_pattern: cl.pattern,
      }));
    return {
      ok: true,
      data: {
        sheet: ws.name,
        scanned_range: addr,
        sampled,
        scanned_cells: scannedCells,
        total_errors: totalErrors,
        clusters: clusterList,
        hint: sampled
          ? `大表已抽样扫描，以上 ${totalErrors} 处错误仅覆盖已读窗口（${scannedCells} 个单元格），不是全表总数。先解释每一簇的根因并给出修复方案；用户确认后再用 set_formulas 修复。`
          : "先在回复里解释每一簇错误的根因（引用代表公式和地址），并给出修复方案；用户确认要修（或指令本就要求修好）后再用 set_formulas 修复。",
      },
      summary: sampled
        ? `${ws.name}!${addr} 抽样扫描发现 ${totalErrors} 处错误（未扫完全表），聚成 ${clusters.size} 类`
        : `${ws.name}!${addr} 共 ${totalErrors} 处错误，聚成 ${clusters.size} 类`,
    };
  });
}

function copyTypeFrom(value: unknown): Excel.RangeCopyType {
  const key = String(value || "all").toLowerCase();
  if (key === "values" || key === "value") return Excel.RangeCopyType.values;
  if (key === "formulas" || key === "formula") return Excel.RangeCopyType.formulas;
  if (key === "formats" || key === "format") return Excel.RangeCopyType.formats;
  return Excel.RangeCopyType.all;
}

function sanitizeSheetName(name: string): string {
  const cleaned = String(name || "")
    .replace(/[\\/?*[\]:]/g, "_")
    .replace(/'/g, "")
    .trim();
  if (!cleaned) throw new Error("Sheet 名称不能为空");
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned;
}

async function copyRange(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sourceSheet = String(args.source_sheet || args.sheet || "").trim();
  const sourceRange = String(args.source_range || args.range || "").trim();
  const destSheetName = String(args.dest_sheet || args.target_sheet || sourceSheet).trim();
  const destStart = String(args.dest_start_cell || args.start_cell || "").trim();
  const transpose = !!args.transpose;
  if (!sourceSheet || !sourceRange) return { ok: false, error: "copy_range 需要 source_sheet 和 source_range" };
  if (!destStart) return { ok: false, error: "copy_range 需要 dest_start_cell，例如 'A1' 或 'E1'" };
  if (isInternalSheet(sourceSheet) || isInternalSheet(destSheetName)) {
    return { ok: false, error: "不能对检查点备份表做复制。" };
  }

  let previewFormulas: any[][] = [];
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let destAddr = "";
  let srcAddr = "";
  let rows = 0;
  let cols = 0;
  let destRows = 0;
  let destCols = 0;

  await Excel.run(async (context) => {
    const src = context.workbook.worksheets.getItem(sourceSheet).getRange(sourceRange);
    src.load("address,rowCount,columnCount,rowIndex,columnIndex");
    await context.sync();
    srcAddr = `${sourceSheet}!${src.address.split("!").pop()}`;
    rows = src.rowCount;
    cols = src.columnCount;
    destRows = transpose ? cols : rows;
    destCols = transpose ? rows : cols;
    const destSheet = context.workbook.worksheets.getItem(destSheetName);
    const dest = destSheet.getRange(destStart).getAbsoluteResizedRange(destRows, destCols);
    dest.load("address,formulas,rowIndex,columnIndex");
    await context.sync();
    destAddr = dest.address.split("!").pop() as string;
    previewFormulas = dest.formulas as any[][];
    if (sourceSheet === destSheetName) {
      const overlap =
        src.rowIndex < dest.rowIndex + destRows &&
        dest.rowIndex < src.rowIndex + rows &&
        src.columnIndex < dest.columnIndex + destCols &&
        dest.columnIndex < src.columnIndex + cols;
      if (overlap) throw new Error("源区域与目标区域重叠，请先复制到空白位置或新 Sheet。");
    }
  });

  const overwritten = countNonEmpty(previewFormulas);
  if (overwritten > 0 && ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `覆盖 ${destSheetName}!${destAddr} 中 ${overwritten} 个已有单元格`,
      `将把 ${srcAddr}（${rows}×${cols}${transpose ? "，转置" : ""}）复制到此处。\n原内容（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了本次覆盖复制。请不要重试，向用户询问意图。" };
  }

  let targetChanged = false;
  await Excel.run(async (context) => {
    const src = context.workbook.worksheets.getItem(sourceSheet).getRange(sourceRange);
    const dest = context.workbook.worksheets.getItem(destSheetName).getRange(destStart).getAbsoluteResizedRange(destRows, destCols);
    dest.load("address,values,formulas");
    await context.sync();
    if (!gridsEqual(dest.formulas as any[][], previewFormulas)) {
      targetChanged = true;
      return;
    }
    oldFormulas = buildRestorableGrid(dest.values as any[][], dest.formulas as any[][]);
    destAddr = dest.address.split("!").pop() as string;
    dest.copyFrom(src, copyTypeFrom(args.copy_type), false, transpose);
    await context.sync();
    dest.load("values,formulas");
    await context.sync();
    expectedValues = dest.values as any[][];
    expectedFormulas = dest.formulas as any[][];
  });
  if (targetChanged) {
    return { ok: false, error: `${destSheetName}!${destAddr} 在预览/确认后发生了变化，本次复制已停止。` };
  }

  const undoId = pushUndo({
    description: `复制到 ${destSheetName}!${destAddr}`,
    sheet: destSheetName,
    address: destAddr,
    formulas: oldFormulas,
    expectedValues,
    expectedFormulas,
  });
  pushOpLog({
    description: `复制 ${srcAddr} → ${destSheetName}!${destAddr}${transpose ? "（转置）" : ""}`,
    kind: "write",
    sheet: destSheetName,
    address: destAddr,
    before: gridPreview(oldFormulas),
    after: gridPreview(expectedFormulas),
    ...(undoId ? { undoEntryId: undoId } : {}),
  });
  if (overwritten === 0 && destRows >= 2 && destCols >= 2) {
    registerArtifact({
      name: destSheetName,
      kind: "range",
      sheet: destSheetName,
      address: destAddr,
      brief: `复制自 ${srcAddr}`,
    });
  }
  return {
    ok: true,
    data: {
      source: srcAddr,
      destination: `${destSheetName}!${destAddr}`,
      rows: destRows,
      cols: destCols,
      transposed: transpose,
      overwritten_cells: overwritten,
    },
    summary: `已将 ${srcAddr} 复制到 ${destSheetName}!${destAddr}${transpose ? "（转置）" : ""}`,
    undoable: !!undoId,
  };
}

async function fillDown(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  const sourceRange = String(args.source_range || "").trim();
  if (!sheetName || !sourceRange) return { ok: false, error: "fill_down 需要 sheet 和 source_range（如 'E2' 或 'E2:G2'）" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能对检查点备份表填充。" };

  let destAddr = String(args.destination_range || "").trim();
  const endRow = args.end_row != null ? toPositiveInt(args.end_row, "end_row") : null;
  if (!destAddr && !endRow) {
    return { ok: false, error: "fill_down 需要 destination_range（如 'E2:E100'）或 end_row（填充到该行）" };
  }

  let previewFormulas: any[][] = [];
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let filledAddr = "";
  let srcRows = 0;
  let srcCols = 0;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const src = sheet.getRange(sourceRange);
    src.load("address,rowCount,columnCount,rowIndex,columnIndex");
    await context.sync();
    srcRows = src.rowCount;
    srcCols = src.columnCount;
    if (endRow != null && endRow <= src.rowIndex + src.rowCount) {
      throw new Error(`end_row 必须大于源区域末行（当前末行 ${src.rowIndex + src.rowCount}）`);
    }
    const dest = destAddr
      ? sheet.getRange(destAddr)
      : sheet.getRangeByIndexes(src.rowIndex, src.columnIndex, endRow! - src.rowIndex, src.columnCount);
    dest.load("address,rowCount,columnCount,rowIndex,columnIndex,formulas");
    await context.sync();
    if (dest.rowIndex !== src.rowIndex || dest.columnIndex !== src.columnIndex) {
      throw new Error("目标区域必须从源区域的左上角开始，并完全包含源区域。");
    }
    if (dest.rowCount < src.rowCount || dest.columnCount < src.columnCount) {
      throw new Error("目标区域必须包含源区域。");
    }
    if (dest.rowCount === src.rowCount && dest.columnCount === src.columnCount) {
      throw new Error("目标区域与源区域相同，没有可填充的单元格。请指定更大的 destination_range 或更大的 end_row。");
    }
    filledAddr = dest.address.split("!").pop() as string;
    previewFormulas = dest.formulas as any[][];
  });

  const overwritten = countNonEmptyOutsideSource(previewFormulas, srcRows, srcCols);
  if (overwritten > 0 && ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `向下填充将覆盖 ${sheetName}!${filledAddr} 中源区域以外的 ${overwritten} 个已有单元格`,
      `原内容（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了本次填充。请不要重试。" };
  }

  let targetChanged = false;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const src = sheet.getRange(sourceRange);
    src.load("rowCount,columnCount,rowIndex,columnIndex");
    await context.sync();
    const dest = destAddr
      ? sheet.getRange(destAddr)
      : sheet.getRangeByIndexes(src.rowIndex, src.columnIndex, endRow! - src.rowIndex, src.columnCount);
    dest.load("address,values,formulas");
    await context.sync();
    if (!gridsEqual(dest.formulas as any[][], previewFormulas)) {
      targetChanged = true;
      return;
    }
    oldFormulas = buildRestorableGrid(dest.values as any[][], dest.formulas as any[][]);
    filledAddr = dest.address.split("!").pop() as string;
    src.autoFill(dest, Excel.AutoFillType.fillDefault);
    await context.sync();
    dest.load("values,formulas");
    await context.sync();
    expectedValues = dest.values as any[][];
    expectedFormulas = dest.formulas as any[][];
  });
  if (targetChanged) {
    return { ok: false, error: `${sheetName}!${filledAddr} 在预览/确认后发生了变化，本次填充已停止。` };
  }

  const undoId = pushUndo({
    description: `填充 ${sheetName}!${filledAddr}`,
    sheet: sheetName,
    address: filledAddr,
    formulas: oldFormulas,
    expectedValues,
    expectedFormulas,
  });
  const fillCheck = checkFormulaResults(filledAddr, expectedValues, expectedFormulas);
  pushOpLog({
    description: `向下填充 ${sheetName}!${filledAddr}`,
    kind: "write",
    sheet: sheetName,
    address: filledAddr,
    before: gridPreview(oldFormulas),
    after: gridPreview(expectedFormulas),
    ...(undoId ? { undoEntryId: undoId } : {}),
  });
  return {
    ok: true,
    data: {
      filled: `${sheetName}!${filledAddr}`,
      source_rows: srcRows,
      source_cols: srcCols,
      ...(fillCheck.formula_cells ? { formula_check: fillCheck, note: formulaCheckNote(fillCheck) } : {}),
    },
    summary: `已将 ${sheetName}!${sourceRange} 填充到 ${filledAddr}`,
    undoable: !!undoId,
  };
}

function cellMatchesFind(value: unknown, find: string, complete: boolean, matchCase: boolean): boolean {
  if (value == null || value === "") return false;
  const s = String(value);
  const hay = matchCase ? s : s.toLowerCase();
  const needle = matchCase ? find : find.toLowerCase();
  return complete ? hay === needle : hay.includes(needle);
}

async function findReplace(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  const findText = String(args.find ?? args.text ?? "");
  const replaceWith = String(args.replace ?? args.replacement ?? "");
  if (!sheetName) return { ok: false, error: "find_replace 需要 sheet" };
  if (!findText) return { ok: false, error: "find_replace 需要非空的 find" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能在检查点备份表上替换。" };

  const complete = !!args.complete_match;
  const matchCase = !!args.match_case;
  let matchCount = 0;
  let previewAddrs: string[] = [];
  let addr = "";
  let cells = 0;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    addr = range.address.split("!").pop() as string;
    cells = range.rowCount * range.columnCount;
    if (cells > MAX_SNAPSHOT_CELLS) {
      matchCount = -1;
      return;
    }
    range.load("values,formulas");
    await context.sync();
    const values = range.values as any[][];
    const formulas = range.formulas as any[][];
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (!cellMatchesFind(formulas[r][c], findText, complete, matchCase) && !cellMatchesFind(values[r]?.[c], findText, complete, matchCase)) {
          continue;
        }
        matchCount++;
        if (previewAddrs.length < 20) {
          previewAddrs.push(`${indexToColLetter(range.columnIndex + c)}${range.rowIndex + r + 1}`);
        }
      }
    }
  });

  if (matchCount === 0) {
    return { ok: true, data: { replaced: 0 }, summary: `未找到「${findText}」` };
  }

  if (ctx.settings.confirmDestructive) {
    const countLabel = matchCount > 0 ? `${matchCount} 处` : "全部匹配项";
    const accepted = await ctx.confirm(
      `在 ${sheetName}!${addr} 把「${findText}」替换为「${replaceWith}」（${countLabel}）`,
      previewAddrs.length ? `示例位置: ${previewAddrs.join("、")}${matchCount > previewAddrs.length ? "…" : ""}` : "区域较大，未列出每一处匹配。"
    );
    if (!accepted) return { ok: false, error: "用户拒绝了本次查找替换。请不要重试。" };
  }

  let replaced = matchCount;
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let canUndo = cells <= MAX_SNAPSHOT_CELLS;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,values,formulas,rowCount,columnCount,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    addr = range.address.split("!").pop() as string;
    if (canUndo) oldFormulas = buildRestorableGrid(range.values as any[][], range.formulas as any[][]);
    const result = range.replaceAll(findText, replaceWith, {
      completeMatch: complete,
      matchCase,
    });
    if (result && typeof (result as any).load === "function") {
      (result as any).load("replacedCount");
      await context.sync();
      if (typeof (result as any).replacedCount === "number") {
        replaced = (result as any).replacedCount;
      }
    } else {
      await context.sync();
    }
    if (canUndo) {
      range.load("values,formulas");
      await context.sync();
      expectedValues = range.values as any[][];
      expectedFormulas = range.formulas as any[][];
    }
  });

  const undoId = canUndo
    ? pushUndo({
        description: `替换 ${sheetName}!${addr} 中的「${findText}」`,
        sheet: sheetName,
        address: addr,
        formulas: oldFormulas,
        expectedValues,
        expectedFormulas,
      })
    : null;
  pushOpLog({
    description: `替换 ${sheetName}!${addr}：「${findText}」→「${replaceWith}」（${replaced} 处）`,
    kind: "write",
    sheet: sheetName,
    address: addr,
    before: canUndo ? gridPreview(oldFormulas) : `「${findText}」×${matchCount}`,
    after: `「${replaceWith}」×${replaced}`,
    ...(undoId ? { undoEntryId: undoId } : {}),
  });
  return {
    ok: true,
    data: { replaced, find: findText, replace: replaceWith, range: `${sheetName}!${addr}` },
    summary: `已替换 ${replaced} 处「${findText}」→「${replaceWith}」`,
    undoable: !!undoId,
  };
}

async function removeDuplicates(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "remove_duplicates 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能对检查点备份表去重。" };
  const hasHeaders = args.has_headers !== false;

  let addr = "";
  let colCount = 0;
  let rowCount = 0;
  let colStart = 0;
  let previewFormulas: any[][] = [];
  let headerNames: string[] = [];

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,rowCount,columnCount,columnIndex,isNullObject,formulas,values");
    await context.sync();
    if ((range as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    addr = range.address.split("!").pop() as string;
    colCount = range.columnCount;
    rowCount = range.rowCount;
    colStart = range.columnIndex;
    previewFormulas = range.formulas as any[][];
    headerNames = ((range.values as any[][])[0] || []).map((v) => String(v ?? "").trim());
  });

  const columns: number[] = Array.isArray(args.columns)
    ? args.columns.map((c: unknown, i: number) => {
        if (typeof c === "number") return c;
        const s = String(c).trim();
        if (/^\d+$/.test(s)) return parseInt(s, 10);
        const byHeader = headerNames.findIndex((h) => h === s);
        if (byHeader >= 0) return byHeader;
        if (/^[A-Za-z]{1,3}$/.test(s)) return colLetterToIndex(s.toUpperCase()) - colStart;
        throw new Error(`无法识别列 columns[${i}]=${JSON.stringify(c)}，请用列字母、0 基列号或表头名`);
      })
    : Array.from({ length: colCount }, (_, i) => i);
  if (!columns.length || columns.some((column) => !Number.isInteger(column) || column < 0 || column >= colCount)) {
    return { ok: false, error: `去重列必须在当前区域内（0 至 ${colCount - 1}），且至少选择一列。` };
  }

  if (ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `对 ${sheetName}!${addr} 去重（保留首次出现的行）`,
      `区域 ${rowCount} 行 × ${colCount} 列。${checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)}\n原内容（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了去重。请不要重试。" };
  }

  let removed = 0;
  let uniqueRemaining = 0;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(addr);
    range.load("formulas");
    await context.sync();
    if (!gridsEqual(range.formulas as any[][], previewFormulas)) {
      throw new Error(`${sheetName}!${addr} 在预览/确认后发生了变化，本次去重已停止。请重新读取后再操作。`);
    }
    checkToolAbort(ctx);
    const result = range.removeDuplicates(columns, hasHeaders);
    result.load("removed,uniqueRemaining");
    await context.sync();
    removed = result.removed;
    uniqueRemaining = result.uniqueRemaining;
  });

  pushOpLog({
    description: `去重 ${sheetName}!${addr}（删除 ${removed} 行，保留 ${uniqueRemaining}）`,
    kind: "structure",
    sheet: sheetName,
    address: addr,
    before: gridPreview(previewFormulas),
    after: `保留 ${uniqueRemaining} 行`,
  });
  return {
    ok: true,
    data: { range: `${sheetName}!${addr}`, removed, unique_remaining: uniqueRemaining, columns },
    summary: `已去重 ${sheetName}!${addr}：删除 ${removed} 行，剩余 ${uniqueRemaining} 行`,
    undoable: false,
  };
}

async function renameSheet(args: any): Promise<ToolResult> {
  const name = String(args.name || args.sheet || "").trim();
  const newName = sanitizeSheetName(String(args.new_name || args.to || ""));
  if (!name) return { ok: false, error: "rename_sheet 需要当前表名 name" };
  if (isInternalSheet(name) || isInternalSheet(newName)) return { ok: false, error: "不能重命名检查点备份表。" };
  if (name === newName) return { ok: true, data: { name }, summary: `Sheet 已是「${name}」` };
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(name);
    sheet.name = newName;
    await context.sync();
    pushOpLog({ description: `重命名 Sheet「${name}」→「${newName}」`, kind: "structure", sheet: newName });
    return { ok: true, data: { from: name, to: newName }, summary: `已将 Sheet「${name}」重命名为「${newName}」` };
  });
}

async function copySheet(args: any): Promise<ToolResult> {
  const name = String(args.name || args.sheet || "").trim();
  if (!name) return { ok: false, error: "copy_sheet 需要源表名 name" };
  if (isInternalSheet(name)) return { ok: false, error: "不能复制检查点备份表。" };
  const requested = args.new_name ? sanitizeSheetName(String(args.new_name)) : "";
  if (requested && isInternalSheet(requested)) return { ok: false, error: "不能使用保留的检查点备份表名。" };
  return Excel.run(async (context) => {
    const src = context.workbook.worksheets.getItem(name);
    const copied = src.copy(Excel.WorksheetPositionType.after, src);
    copied.load("name");
    await context.sync();
    let finalName = copied.name;
    if (requested && copied.name !== requested) {
      finalName = await uniqueSheetName(context, requested);
      copied.name = finalName;
    }
    copied.activate();
    await context.sync();
    pushOpLog({ description: `复制 Sheet「${name}」为「${finalName}」`, kind: "structure", sheet: finalName });
    registerArtifact({
      name: finalName,
      kind: "sheet",
      sheet: finalName,
      brief: `复制自「${name}」`,
    });
    return { ok: true, data: { from: name, to: finalName }, summary: `已复制 Sheet「${name}」为「${finalName}」` };
  });
}

function fnValue(result: Excel.FunctionResult<any>): number | null {
  if (result.error) return null;
  const v = result.value;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

async function columnStats(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "column_stats 需要 sheet" };
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,isNullObject,rowIndex,columnIndex");
    await context.sync();
    if ((used as any).isNullObject) return { ok: false, error: `${sheetName} 是空表` };

    const totalRows = used.rowCount;
    const totalCols = used.columnCount;
    const headerCount = args.has_headers === false ? 0 : 1;
    const dataRows = Math.max(0, totalRows - headerCount);
    if (dataRows === 0) {
      return { ok: false, error: `${sheetName}!${used.address.split("!").pop()} 没有数据行` };
    }

    const headerRange = used.getCell(0, 0).getAbsoluteResizedRange(1, totalCols);
    headerRange.load("values");
    await context.sync();
    const headers = (headerRange.values as any[][])[0] || [];

    const wanted: number[] = [];
    if (Array.isArray(args.columns) && args.columns.length) {
      for (let i = 0; i < args.columns.length; i++) {
        const raw = args.columns[i];
        if (typeof raw === "number") {
          wanted.push(Math.trunc(raw));
          continue;
        }
        const s = String(raw).trim();
        const byHeader = headers.findIndex((h) => String(h).trim() === s);
        if (byHeader >= 0) wanted.push(byHeader);
        else if (/^[A-Za-z]{1,3}$/.test(s)) wanted.push(colLetterToIndex(s.toUpperCase()) - used.columnIndex);
        else if (/^\d+$/.test(s)) wanted.push(parseInt(s, 10));
      }
    } else {
      for (let c = 0; c < Math.min(totalCols, 24); c++) wanted.push(c);
    }
    const colIndexes = [...new Set(wanted.filter((c) => c >= 0 && c < totalCols))];
    if (!colIndexes.length) return { ok: false, error: "没有匹配到要统计的列" };

    const fn = context.workbook.functions;
    const pending = colIndexes.map((c) => {
      const data = used.getCell(headerCount, c).getAbsoluteResizedRange(dataRows, 1);
      const countA = fn.countA(data);
      const blanks = fn.countBlank(data);
      const nums = fn.count(data);
      const sum = fn.sum(data);
      const min = fn.min(data);
      const max = fn.max(data);
      const avg = fn.average(data);
      countA.load("value,error");
      blanks.load("value,error");
      nums.load("value,error");
      sum.load("value,error");
      min.load("value,error");
      max.load("value,error");
      avg.load("value,error");
      const uniqueSample = dataRows <= 8000 ? data : null;
      if (uniqueSample) uniqueSample.load("values");
      return { c, countA, blanks, nums, sum, min, max, avg, uniqueSample };
    });
    await context.sync();

    const columns = pending.map((p) => {
      const header = String(headers[p.c] ?? "").trim() || `第${p.c + 1}列`;
      const numericCount = fnValue(p.nums) ?? 0;
      let unique: number | null = null;
      if (p.uniqueSample) {
        const seen = new Set<string>();
        for (const row of p.uniqueSample.values as any[][]) {
          const v = row[0];
          if (v === "" || v == null) continue;
          seen.add(String(v));
        }
        unique = seen.size;
      }
      const rec: Record<string, unknown> = {
        header,
        column_index: p.c,
        non_empty: fnValue(p.countA),
        blanks: fnValue(p.blanks),
        numeric_count: numericCount,
      };
      if (numericCount > 0) {
        rec.sum = fnValue(p.sum);
        rec.min = fnValue(p.min);
        rec.max = fnValue(p.max);
        rec.average = fnValue(p.avg);
      }
      if (unique != null) rec.unique_count = unique;
      return rec;
    });

    const addr = used.address.split("!").pop();
    return {
      ok: true,
      data: {
        sheet: sheetName,
        range: addr,
        row_count: totalRows,
        data_rows: dataRows,
        column_count: totalCols,
        unique_counted: dataRows <= 8000,
        columns,
      },
      summary: `已统计 ${sheetName}!${addr} 的 ${columns.length} 列（${dataRows} 行数据）`,
    };
  });
}

async function applyConditionalFormat(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  const rangeAddr = String(args.range || "").trim();
  const rule = String(args.rule || "duplicate").toLowerCase();
  if (!sheetName || !rangeAddr) return { ok: false, error: "apply_conditional_format 需要 sheet 和 range" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能对检查点备份表设置条件格式。" };
  const fill = String(args.fill_color || (rule === "duplicate" ? "#FFEB9C" : "#FFC7CE"));
  const font = args.font_color ? String(args.font_color) : "";

  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddr);
    range.load("address");
    await context.sync();
    const addr = range.address.split("!").pop() as string;

    if (rule === "duplicate" || rule === "unique") {
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
      cf.preset.rule = {
        criterion:
          rule === "unique"
            ? Excel.ConditionalFormatPresetCriterion.uniqueValues
            : Excel.ConditionalFormatPresetCriterion.duplicateValues,
      };
      cf.preset.format.fill.color = fill;
      if (font) cf.preset.format.font.color = font;
    } else if (rule === "color_scale") {
      range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    } else if (rule === "data_bar") {
      range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    } else {
      const opMap: Record<string, Excel.ConditionalCellValueOperator> = {
        greater_than: Excel.ConditionalCellValueOperator.greaterThan,
        less_than: Excel.ConditionalCellValueOperator.lessThan,
        equal: Excel.ConditionalCellValueOperator.equalTo,
        between: Excel.ConditionalCellValueOperator.between,
      };
      const op = opMap[rule];
      if (!op) throw new Error(`不支持的 rule: ${rule}。可用 duplicate/unique/greater_than/less_than/equal/between/color_scale/data_bar`);
      if (args.value == null) throw new Error(`${rule} 需要 value`);
      const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      cf.cellValue.rule = {
        formula1: String(args.value),
        ...(rule === "between" ? { formula2: String(args.value2 ?? "") } : {}),
        operator: op,
      };
      cf.cellValue.format.fill.color = fill;
      if (font) cf.cellValue.format.font.color = font;
    }
    await context.sync();
    pushOpLog({
      description: `条件格式 ${rule} → ${sheetName}!${addr}`,
      kind: "format",
      sheet: sheetName,
      address: addr,
    });
    return {
      ok: true,
      data: { range: `${sheetName}!${addr}`, rule },
      summary: `已为 ${sheetName}!${addr} 应用条件格式（${rule}）`,
    };
  });
}

function normalizeChoiceOptions(raw: unknown): ChoiceOption[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const options: ChoiceOption[] = [];
  raw.forEach((item, i) => {
    if (!item || typeof item !== "object") return;
    const rec = item as Record<string, unknown>;
    const label = String(rec.label ?? rec.title ?? rec.text ?? "").trim();
    if (!label) return;
    let id = String(rec.id ?? "").trim() || `opt_${i + 1}`;
    if (seen.has(id)) id = `${id}_${i + 1}`;
    seen.add(id);
    const description = rec.description != null ? String(rec.description).trim() : "";
    options.push({
      id,
      label,
      ...(description ? { description } : {}),
      ...(rec.recommended ? { recommended: true } : {}),
    });
  });
  return options.slice(0, 4);
}

async function askUserChoice(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const options = normalizeChoiceOptions(args.options);
  if (options.length < 2) {
    return { ok: false, error: "ask_user_choice 至少需要 2 个选项（options[].label）" };
  }
  const title = String(args.title || "请选择一个方向").trim();
  const question = String(args.question || "").trim();
  const result = await ctx.choose(title, question, options);
  if (result.kind === "cancelled") {
    return { ok: false, error: "用户取消了选择。请停止当前任务，等待用户下一条指令。" };
  }
  if (result.kind === "custom") {
    return {
      ok: true,
      data: {
        custom: result.text,
        hint: `用户没有采用给出的选项，而是要求：「${result.text}」。请按这个要求执行，不要再给选项让用户选。`,
      },
      summary: `用户提出了自己的做法：${result.text}`,
    };
  }
  if (result.kind === "skipped") {
    const rec = options.find((o) => o.recommended) || options[0];
    return {
      ok: true,
      data: {
        skipped: true,
        hint: `用户让你自行判断。请按推荐方向「${rec.label}」执行，不要再问一遍。`,
        recommended_id: rec.id,
        recommended_label: rec.label,
      },
      summary: "用户让你自行判断",
    };
  }
  return {
    ok: true,
    data: { selected_id: result.id, selected_label: result.label },
    summary: `用户选择了「${result.label}」`,
  };
}

const MAX_LOOKUP_DATA_ROWS = 30000;
const MAX_SPLIT_ROWS = 20000;

type SheetLayout = {
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  headers: string[];
};

async function loadSheetLayout(sheetName: string, rangeArg?: string): Promise<SheetLayout> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = rangeArg ? sheet.getRange(rangeArg) : sheet.getUsedRangeOrNullObject(true);
    used.load("rowCount,columnCount,rowIndex,columnIndex,isNullObject");
    await context.sync();
    if ((used as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    const headerRange = used.getCell(0, 0).getAbsoluteResizedRange(1, used.columnCount);
    headerRange.load("values");
    await context.sync();
    return {
      rowIndex: used.rowIndex,
      columnIndex: used.columnIndex,
      rowCount: used.rowCount,
      columnCount: used.columnCount,
      headers: ((headerRange.values as any[][])[0] || []).map((v) => String(v ?? "").trim()),
    };
  });
}

function parseDestStartColumn(spec: unknown, dest: SheetLayout): number {
  if (spec == null || spec === "") return dest.columnIndex + dest.columnCount;
  const s = String(spec).trim();
  const byHeader = dest.headers.findIndex((h) => h === s);
  if (byHeader >= 0) return dest.columnIndex + byHeader;
  if (/^[A-Za-z]{1,3}$/.test(s)) return colLetterToIndex(s.toUpperCase());
  if (typeof spec === "number" || /^\d+$/.test(s)) {
    const n = typeof spec === "number" ? Math.trunc(spec) : parseInt(s, 10);
    if (!Number.isFinite(n) || n < 0 || n >= dest.headers.length) {
      throw new Error(`dest_start_column 列号 ${n} 超出范围 0..${dest.headers.length - 1}（相对当前表）`);
    }
    return dest.columnIndex + n;
  }
  throw new Error(`无法识别 dest_start_column「${s}」，请用列字母、表头名或相对当前表的 0 基列号`);
}

function returnColumnList(args: any): unknown[] {
  if (Array.isArray(args.return_columns) && args.return_columns.length) return args.return_columns;
  if (args.return_column != null && args.return_column !== "") return [args.return_column];
  return [];
}

async function addLookupColumn(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const destSheet = String(args.dest_sheet || args.sheet || "").trim();
  const lookupSheet = String(args.lookup_sheet || "").trim();
  if (!destSheet) return { ok: false, error: "add_lookup_column 需要 dest_sheet" };
  if (!lookupSheet) return { ok: false, error: "add_lookup_column 需要 lookup_sheet" };
  if (isInternalSheet(destSheet) || isInternalSheet(lookupSheet)) {
    return { ok: false, error: "不能对检查点备份表做查找匹配。" };
  }
  const returns = returnColumnList(args);
  if (!returns.length) return { ok: false, error: "add_lookup_column 需要 return_columns 或 return_column" };
  if (args.dest_key == null || args.dest_key === "") return { ok: false, error: "add_lookup_column 需要 dest_key（目标表匹配列）" };
  if (args.lookup_key == null || args.lookup_key === "") return { ok: false, error: "add_lookup_column 需要 lookup_key（查找表匹配列）" };

  const destHasHeaders = args.dest_has_headers === false ? false : true;
  const lookupHasHeaders = args.lookup_has_headers === false ? false : true;
  const engine = String(args.engine || "xlookup").toLowerCase() === "vlookup" ? "vlookup" : "xlookup";

  const dest = await loadSheetLayout(destSheet, args.dest_range);
  const lookup = await loadSheetLayout(lookupSheet, args.lookup_range);
  const destKeyRel = resolveColSpec(dest.headers, args.dest_key, dest.columnIndex, "dest_key");
  const lookupKeyRel = resolveColSpec(lookup.headers, args.lookup_key, lookup.columnIndex, "lookup_key");
  const returnRels = returns.map((spec, i) =>
    resolveColSpec(lookup.headers, spec, lookup.columnIndex, `return_columns[${i}]`)
  );

  const destHeaderCount = destHasHeaders ? 1 : 0;
  const lookupHeaderCount = lookupHasHeaders ? 1 : 0;
  const destDataRows = dest.rowCount - destHeaderCount;
  const lookupDataRows = lookup.rowCount - lookupHeaderCount;
  if (destDataRows < 1) return { ok: false, error: `${destSheet} 没有数据行，无法填写查找公式` };
  if (lookupDataRows < 1) return { ok: false, error: `${lookupSheet} 没有数据行，无法作为查找表` };
  if (destDataRows > MAX_LOOKUP_DATA_ROWS) {
    return { ok: false, error: `目标表有 ${destDataRows} 行，超过 ${MAX_LOOKUP_DATA_ROWS}，请缩小 dest_range` };
  }

  const destFirstDataRow = dest.rowIndex + destHeaderCount + 1;
  const destLastDataRow = dest.rowIndex + dest.rowCount;
  const lookupFirstDataRow = lookup.rowIndex + lookupHeaderCount + 1;
  const lookupLastDataRow = lookup.rowIndex + lookup.rowCount;
  const destStartAbs = parseDestStartColumn(args.dest_start_column, dest);
  const destHeaders = Array.isArray(args.dest_headers)
    ? args.dest_headers.map((h: unknown) => String(h ?? "").trim())
    : returnRels.map((rel) => lookup.headers[rel] || `列${indexToColLetter(lookup.columnIndex + rel)}`);
  while (destHeaders.length < returnRels.length) destHeaders.push(`列${destHeaders.length + 1}`);

  const formulas = buildLookupFormulaGrid({
    destKeyLetter: indexToColLetter(dest.columnIndex + destKeyRel),
    destFirstDataRow,
    destLastDataRow,
    lookupSheet,
    lookupKeyLetter: indexToColLetter(lookup.columnIndex + lookupKeyRel),
    lookupFirstDataRow,
    lookupLastDataRow,
    returns: returnRels.map((rel, i) => ({
      letter: indexToColLetter(lookup.columnIndex + rel),
      header: destHeaders[i] || lookup.headers[rel] || `列${i + 1}`,
    })),
    ifNotFound: args.if_not_found ?? "",
    engine,
    includeHeaders: destHasHeaders,
  });

  const startCell = `${indexToColLetter(destStartAbs)}${dest.rowIndex + 1}`;
  const written = await writeCells({ sheet: destSheet, start_cell: startCell, formulas }, ctx, true);
  if (!written.ok) return written;
  const check = (written.data as { formula_check?: { blank_results: number } }).formula_check;
  const notFound = args.if_not_found ?? "";
  return {
    ok: true,
    data: {
      ...(written.data as object),
      ...(check?.blank_results && notFound === ""
        ? { unmatched_hint: `有 ${check.blank_results} 个结果为空：多半是没匹配上（if_not_found 为空时显示为空白），也可能是查找表对应值本身为空。报告给用户，必要时核对键的类型和空格。` }
        : {}),
      dest_sheet: destSheet,
      lookup_sheet: lookupSheet,
      dest_key: dest.headers[destKeyRel] || args.dest_key,
      lookup_key: lookup.headers[lookupKeyRel] || args.lookup_key,
      return_columns: formulas[0],
      engine,
      data_rows: destDataRows,
    },
    summary: `已在 ${destSheet} 写入 ${returnRels.length} 列查找公式（按 ${lookupSheet} 匹配，${destDataRows} 行）`,
    undoable: written.undoable,
  };
}

async function splitColumn(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "split_column 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能对检查点备份表拆分列。" };
  if (args.column == null || args.column === "") return { ok: false, error: "split_column 需要 column（表头或列字母）" };
  let delimiter: string;
  try {
    delimiter = resolveDelimiter(args.delimiter);
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
  const hasHeaders = args.has_headers === false ? false : true;
  const trimParts = args.trim !== false;

  const layout = await loadSheetLayout(sheetName, args.range);
  const srcRel = resolveColSpec(layout.headers, args.column, layout.columnIndex, "column");
  const headerCount = hasHeaders ? 1 : 0;
  const dataRows = layout.rowCount - headerCount;
  if (dataRows < 1) return { ok: false, error: `${sheetName} 没有数据行` };
  if (dataRows > MAX_SPLIT_ROWS) {
    return { ok: false, error: `有 ${dataRows} 行，超过 ${MAX_SPLIT_ROWS}，请缩小 range` };
  }

  let srcValues: any[][] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
    const col = used.getCell(0, srcRel).getAbsoluteResizedRange(layout.rowCount, 1);
    col.load("values");
    await context.sync();
    srcValues = col.values as any[][];
  });

  const dataVals = srcValues.slice(headerCount).map((row) => row[0]);
  let width = args.max_parts != null ? Math.max(1, Math.trunc(Number(args.max_parts))) : 0;
  if (!width) {
    let maxSeen = 1;
    let foundSplit = false;
    for (const v of dataVals) {
      if (v == null || String(v).trim() === "") continue;
      const parts = splitCellText(v, delimiter, undefined, trimParts);
      if (parts.length > 1) foundSplit = true;
      maxSeen = Math.max(maxSeen, parts.length);
    }
    if (!foundSplit) {
      return {
        ok: false,
        error: `列「${layout.headers[srcRel] || args.column}」中没有找到分隔符，请确认 delimiter（当前按 ${JSON.stringify(delimiter)} 拆）`,
      };
    }
    width = Math.min(maxSeen, 8);
  }

  const srcHeader = layout.headers[srcRel] || String(args.column);
  const destHeaders = Array.isArray(args.dest_headers)
    ? args.dest_headers.map((h: unknown) => String(h ?? "").trim())
    : Array.from({ length: width }, (_, i) => `${srcHeader}_${i + 1}`);
  while (destHeaders.length < width) destHeaders.push(`${srcHeader}_${destHeaders.length + 1}`);

  const values: any[][] = [];
  if (hasHeaders) values.push(destHeaders.slice(0, width));
  for (const v of dataVals) {
    values.push(padSplitParts(splitCellText(v, delimiter, width, trimParts), width));
  }

  const destStartAbs = parseDestStartColumn(args.dest_start_column, layout);
  const startCell = `${indexToColLetter(destStartAbs)}${layout.rowIndex + 1}`;
  const written = await writeCells({ sheet: sheetName, start_cell: startCell, values }, ctx, false);
  if (!written.ok) return written;
  return {
    ok: true,
    data: {
      ...(written.data as object),
      column: srcHeader,
      delimiter,
      parts: width,
      dest_headers: destHeaders.slice(0, width),
    },
    summary: `已将 ${sheetName}「${srcHeader}」按分隔符拆成 ${width} 列`,
    undoable: written.undoable,
  };
}

async function combineColumns(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "combine_columns 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能对检查点备份表合并列。" };
  const cols = Array.isArray(args.columns) ? args.columns : [];
  if (cols.length < 2) return { ok: false, error: "combine_columns 需要至少两列 columns" };
  const hasHeaders = args.has_headers === false ? false : true;
  const separator = args.separator == null ? "" : String(args.separator);
  const ignoreEmpty = args.ignore_empty !== false;

  const layout = await loadSheetLayout(sheetName, args.range);
  const rels = cols.map((spec: unknown, i: number) =>
    resolveColSpec(layout.headers, spec, layout.columnIndex, `columns[${i}]`)
  );
  const headerCount = hasHeaders ? 1 : 0;
  const dataRows = layout.rowCount - headerCount;
  if (dataRows < 1) return { ok: false, error: `${sheetName} 没有数据行` };
  if (dataRows > MAX_LOOKUP_DATA_ROWS) {
    return { ok: false, error: `有 ${dataRows} 行，超过 ${MAX_LOOKUP_DATA_ROWS}，请缩小 range` };
  }

  const destHeader =
    String(args.dest_header || "").trim() ||
    rels.map((rel: number) => layout.headers[rel] || indexToColLetter(layout.columnIndex + rel)).join(separator || "_");
  const destStartAbs = parseDestStartColumn(args.dest_start_column ?? args.dest_column, layout);
  const firstDataRow = layout.rowIndex + headerCount + 1;
  const lastDataRow = layout.rowIndex + layout.rowCount;
  const letters: string[] = rels.map((rel: number) => indexToColLetter(layout.columnIndex + rel));

  const formulas: string[][] = [];
  if (hasHeaders) formulas.push([destHeader]);
  for (let r = firstDataRow; r <= lastDataRow; r++) {
    formulas.push([textJoinFormula(separator, letters.map((letter) => `${letter}${r}`), ignoreEmpty)]);
  }

  const startCell = `${indexToColLetter(destStartAbs)}${layout.rowIndex + 1}`;
  const written = await writeCells({ sheet: sheetName, start_cell: startCell, formulas }, ctx, true);
  if (!written.ok) return written;
  return {
    ok: true,
    data: { ...(written.data as object), dest_header: destHeader, columns: letters, separator },
    summary: `已在 ${sheetName} 用公式合并 ${letters.join("+")} 为「${destHeader}」`,
    undoable: written.undoable,
  };
}

async function cleanColumn(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "clean_column 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能清洗检查点备份表。" };
  if (args.column == null || args.column === "") return { ok: false, error: "clean_column 需要 column（表头或列字母）" };
  let ops: ReturnType<typeof normalizeCleanOps>;
  try {
    ops = normalizeCleanOps(args.operations ?? args.operation);
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
  const hasHeaders = args.has_headers === false ? false : true;
  const layout = await loadSheetLayout(sheetName, args.range);
  const srcRel = resolveColSpec(layout.headers, args.column, layout.columnIndex, "column");
  const headerCount = hasHeaders ? 1 : 0;
  const dataRows = layout.rowCount - headerCount;
  if (dataRows < 1) return { ok: false, error: `${sheetName} 没有数据行` };
  if (dataRows > MAX_SPLIT_ROWS) return { ok: false, error: `有 ${dataRows} 行，超过 ${MAX_SPLIT_ROWS}，请缩小 range` };

  let srcValues: any[][] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
    const col = used.getCell(0, srcRel).getAbsoluteResizedRange(layout.rowCount, 1);
    col.load("values");
    await context.sync();
    srcValues = col.values as any[][];
  });

  const dataVals = srcValues.slice(headerCount).map((row) => row[0]);
  const defaultYear =
    args.default_year != null
      ? Math.trunc(Number(args.default_year))
      : inferDefaultYear(dataVals, new Date().getFullYear());
  const cleaned = dataVals.map((v) => applyCleanOps(v, ops, { defaultYear }));
  const changed = cleaned.filter((c) => c.changed).length;
  const unparsed = cleaned.filter((c) => c.unparsed).length;
  if (changed === 0 && unparsed === 0) {
    return {
      ok: true,
      data: { column: layout.headers[srcRel] || args.column, operations: ops, changed: 0 },
      summary: `列「${layout.headers[srcRel] || args.column}」无需清洗`,
    };
  }

  const srcHeader = layout.headers[srcRel] || String(args.column);
  // 金额转大写会把数字变成文字，默认写到新列，原金额还能继续参与计算
  const toUpperAmount = ops.includes("rmb_upper");
  const inplace = !args.dest_start_column && (toUpperAmount ? args.inplace === true : args.inplace !== false);
  const destHeader = String(args.dest_header || "").trim() || (inplace ? srcHeader : `${srcHeader}_${toUpperAmount ? "大写" : "清洗"}`);
  const values: any[][] = [];
  if (hasHeaders) values.push([destHeader]);
  for (const c of cleaned) values.push([c.value ?? ""]);

  const destStartAbs = inplace ? layout.columnIndex + srcRel : parseDestStartColumn(args.dest_start_column, layout);
  const startCell = `${indexToColLetter(destStartAbs)}${layout.rowIndex + 1}`;
  const written = await writeCells({ sheet: sheetName, start_cell: startCell, values }, ctx, false);
  if (!written.ok) return written;
  return {
    ok: true,
    data: {
      ...(written.data as object),
      column: srcHeader,
      operations: ops,
      changed,
      unparsed,
      default_year: ops.includes("to_date") ? defaultYear : undefined,
      inplace,
    },
    summary: `已清洗「${srcHeader}」：${ops.join("+")}，改写 ${changed} 格${unparsed ? `，${unparsed} 格无法解析已保留原值` : ""}`,
    undoable: written.undoable,
  };
}

async function convertToValues(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "convert_to_values 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能转换检查点备份表。" };
  const rangeArg = String(args.range || "").trim();

  let addr = "";
  let formulaCount = 0;
  let oldFormulas: any[][] = [];
  let expectedValues: any[][] = [];
  let expectedFormulas: any[][] = [];
  let previewFormulas: any[][] = [];

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = rangeArg ? sheet.getRange(rangeArg) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,values,formulas,rowCount,columnCount,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    addr = range.address.split("!").pop() as string;
    const formulas = range.formulas as any[][];
    previewFormulas = formulas;
    for (const row of formulas) {
      for (const f of row) {
        if (typeof f === "string" && f.startsWith("=")) formulaCount++;
      }
    }
  });

  if (formulaCount === 0) {
    return { ok: true, data: { range: `${sheetName}!${addr}`, formula_cells: 0 }, summary: `${sheetName}!${addr} 没有公式，无需转成值` };
  }
  if (ctx.settings.confirmDestructive) {
    const accepted = await ctx.confirm(
      `把 ${sheetName}!${addr} 中 ${formulaCount} 个公式转成当前显示值`,
      `转成值后公式会消失，只保留计算结果。可用撤销或任务检查点恢复。\n原公式（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了转成值。请不要重试。" };
  }

  let targetChanged = false;
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = rangeArg ? sheet.getRange(rangeArg) : sheet.getUsedRange();
    range.load("address,values,formulas");
    await context.sync();
    if (!gridsEqual(range.formulas as any[][], previewFormulas)) {
      targetChanged = true;
      return;
    }
    oldFormulas = buildRestorableGrid(range.values as any[][], range.formulas as any[][]);
    addr = range.address.split("!").pop() as string;
    const currentVals = (range.values as any[][]).map((row) => [...row]);
    range.values = currentVals as any;
    await context.sync();
    range.load("values,formulas");
    await context.sync();
    expectedValues = range.values as any[][];
    expectedFormulas = range.formulas as any[][];
  });
  if (targetChanged) {
    return { ok: false, error: `${sheetName}!${addr} 在确认后发生了变化，本次转成值已停止。` };
  }

  const undoId = pushUndo({
    description: `公式转成值 ${sheetName}!${addr}`,
    sheet: sheetName,
    address: addr,
    formulas: oldFormulas,
    expectedValues,
    expectedFormulas,
  });
  pushOpLog({
    description: `公式转成值 ${sheetName}!${addr}（${formulaCount} 个公式）`,
    kind: "write",
    sheet: sheetName,
    address: addr,
    before: gridPreview(oldFormulas),
    after: gridPreview(expectedValues),
    ...(undoId ? { undoEntryId: undoId } : {}),
  });
  return {
    ok: true,
    data: { range: `${sheetName}!${addr}`, formula_cells: formulaCount },
    summary: `已将 ${sheetName}!${addr} 的 ${formulaCount} 个公式转成值`,
    undoable: !!undoId,
  };
}

async function polishRange(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "polish_range 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能整理检查点备份表。" };
  const rangeArg = String(args.range || "").trim();
  const header = args.header !== false;
  const freeze = args.freeze_header !== false;
  const borders = args.border !== false;
  const autofit = args.autofit !== false;
  const numberFormats = args.number_formats !== false;

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = rangeArg ? sheet.getRange(rangeArg) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,rowCount,columnCount,rowIndex,columnIndex,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) return { ok: false, error: `${sheetName} 是空表` };
    const addr = range.address.split("!").pop() as string;
    const headerRange = range.getCell(0, 0).getAbsoluteResizedRange(1, range.columnCount);
    headerRange.load("values");
    await context.sync();
    const headers = ((headerRange.values as any[][])[0] || []).map((v) => String(v ?? "").trim());

    if (header) {
      const hf = headerRange.format;
      hf.font.bold = true;
      hf.font.color = String(args.header_font_color || "#FFFFFF");
      hf.fill.color = String(args.header_fill || "#107C41");
      hf.horizontalAlignment = Excel.HorizontalAlignment.center;
    }
    if (borders) {
      const f = range.format;
      const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"] as const;
      for (const edge of edges) {
        const b = f.borders.getItem(edge);
        b.style = Excel.BorderLineStyle.continuous;
        b.color = "#B0B7BD";
      }
    }
    if (numberFormats && range.rowCount > (header ? 1 : 0)) {
      for (let c = 0; c < range.columnCount; c++) {
        const fmt = guessColumnNumberFormat(headers[c] || "");
        if (!fmt) continue;
        const data = range.getCell(header ? 1 : 0, c).getAbsoluteResizedRange(range.rowCount - (header ? 1 : 0), 1);
        const nf = Array.from({ length: range.rowCount - (header ? 1 : 0) }, () => [fmt]);
        data.numberFormat = nf as any;
      }
    }
    if (autofit) range.format.autofitColumns();
    if (freeze && header) {
      sheet.freezePanes.freezeRows(range.rowIndex + 1);
    }
    await context.sync();
    pushOpLog({ description: `整理格式 ${sheetName}!${addr}`, kind: "format", sheet: sheetName, address: addr });
    return {
      ok: true,
      data: { range: `${sheetName}!${addr}`, header, freeze_header: freeze && header, border: borders, autofit, number_formats: numberFormats },
      summary: `已整理 ${sheetName}!${addr}：${[header && "表头", borders && "边框", autofit && "列宽", numberFormats && "数字格式", freeze && header && "冻结表头"].filter(Boolean).join("、") || "无格式选项"}`,
    };
  });
}

async function deleteBlankRows(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "delete_blank_rows 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能清理检查点备份表的空行。" };
  const hasHeaders = args.has_headers === false ? false : true;

  let addr = "";
  let blankExcelRows: number[] = [];
  let previewFormulas: any[][] = [];
  let previewUsedAddress = "";

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address,rowCount,columnCount,rowIndex,values,formulas,isNullObject");
    const focus = args.range ? sheet.getRange(args.range) : used;
    focus.load("address,rowCount,rowIndex,isNullObject");
    await context.sync();
    if ((used as any).isNullObject || (focus as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    addr = focus.address.split("!").pop() as string;
    previewUsedAddress = used.address;
    previewFormulas = used.formulas as any[][];
    const values = used.values as any[][];
    const formulas = used.formulas as any[][];
    const usedRow0 = used.rowIndex;
    const focusStart = focus.rowIndex;
    const focusEnd = focus.rowIndex + focus.rowCount - 1;
    for (let sheetRow0 = focusStart; sheetRow0 <= focusEnd; sheetRow0++) {
      const excelRow = sheetRow0 + 1;
      if (hasHeaders && sheetRow0 === focusStart) continue;
      const i = sheetRow0 - usedRow0;
      if (i < 0 || i >= values.length) continue;
      if (isRowBlank(values[i] || []) && isRowBlank(formulas[i] || [])) blankExcelRows.push(excelRow);
    }
  });

  if (!blankExcelRows.length) {
    return { ok: true, data: { deleted: 0, range: `${sheetName}!${addr}` }, summary: `${sheetName} 没有整行空白` };
  }
  if (ctx.settings.confirmDestructive) {
    const shown = blankExcelRows.slice(0, 12).join("、");
    const extra = blankExcelRows.length > 12 ? ` 等 ${blankExcelRows.length} 行` : "";
    const accepted = await ctx.confirm(
      `删除 ${sheetName} 中 ${blankExcelRows.length} 行空白`,
      `将删除第 ${shown}${extra} 行。${checkpointRestoreHint(ctx.settings.autoCheckpoint !== false)}\n区域原内容（节选）:\n${previewGrid(previewFormulas)}`
    );
    if (!accepted) return { ok: false, error: "用户拒绝删除空行。请不要重试。" };
  }

  const blocks: { start: number; end: number }[] = [];
  for (const r of blankExcelRows) {
    const last = blocks[blocks.length - 1];
    if (last && last.end === r - 1) last.end = r;
    else blocks.push({ start: r, end: r });
  }
  blocks.sort((a, b) => b.start - a.start);

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load("address,formulas,isNullObject");
    await context.sync();
    if (used.isNullObject || used.address !== previewUsedAddress || !gridsEqual(used.formulas as any[][], previewFormulas)) {
      throw new Error(`${sheetName} 在预览/确认后发生了变化，本次删除空行已停止。请重新读取后再操作。`);
    }
    checkToolAbort(ctx);
    for (const b of blocks) {
      sheet.getRange(`${b.start}:${b.end}`).delete(Excel.DeleteShiftDirection.up);
    }
    await context.sync();
  });

  pushOpLog({
    description: `删除 ${sheetName} 空白行 ${blankExcelRows.length} 行`,
    kind: "structure",
    sheet: sheetName,
    address: addr,
    before: gridPreview(previewFormulas),
    after: `删除 ${blankExcelRows.length} 行`,
  });
  return {
    ok: true,
    data: { deleted: blankExcelRows.length, rows: blankExcelRows, range: `${sheetName}!${addr}` },
    summary: `已删除 ${sheetName} 中 ${blankExcelRows.length} 行空白`,
    undoable: false,
  };
}

async function freezePanes(args: any): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "freeze_panes 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能冻结检查点备份表。" };
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    if (args.unfreeze) {
      sheet.freezePanes.unfreeze();
      await context.sync();
      pushOpLog({ description: `取消冻结「${sheetName}」窗格`, kind: "format", sheet: sheetName });
      return { ok: true, data: { sheet: sheetName, unfrozen: true }, summary: `已取消「${sheetName}」的冻结窗格` };
    }
    let rows = args.rows ?? (args.columns != null ? 0 : 1);
    let cols = args.columns ?? 0;
    if (args.cell) {
      // cell 沿用 Excel UI 的分界单元格语义；freezeAt 接收的是实际冻结区域。
      const cell = sheet.getRange(String(args.cell));
      cell.load("rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      if (cell.rowCount !== 1 || cell.columnCount !== 1) throw new Error("cell 必须是单个分界单元格，例如 B2");
      rows = cell.rowIndex;
      cols = cell.columnIndex;
    }
    if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 0 || cols < 0 || rows >= 1048576 || cols >= 16384) {
      throw new Error("冻结行列数必须是 Excel 范围内的非负整数");
    }
    sheet.freezePanes.unfreeze();
    if (rows && cols) sheet.freezePanes.freezeAt(sheet.getRangeByIndexes(0, 0, rows, cols));
    else if (rows) sheet.freezePanes.freezeRows(rows);
    else if (cols) sheet.freezePanes.freezeColumns(cols);
    await context.sync();
    const detail = args.cell ? `从 ${args.cell} 起` : `行 ${rows}、列 ${cols}`;
    pushOpLog({ description: `冻结「${sheetName}」窗格（${detail}）`, kind: "format", sheet: sheetName });
    return {
      ok: true,
      data: { sheet: sheetName, cell: args.cell || null, rows, columns: cols },
      summary: `已冻结「${sheetName}」窗格（${detail}）`,
    };
  });
}

// ---------------- 填充空白 / 取消合并 ----------------

const MAX_FILL_SCAN_CELLS = 300_000;
const FILL_SYNC_RUNS = 400;

/** 列出合并区域要 ExcelApi 1.13；更老的 Excel 只能按列取消合并后向下补 */
function canListMergedAreas(): boolean {
  try {
    return typeof Office !== "undefined" && !!Office.context?.requirements?.isSetSupported("ExcelApi", "1.13");
  } catch {
    return false;
  }
}

async function fillBlankCells(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "fill_blank_cells 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能修改检查点备份表。" };
  const mode: FillMode = args.mode === "left" || args.mode === "value" ? args.mode : "above";
  if (mode === "value" && (args.value == null || args.value === "" || typeof args.value === "object")) {
    return { ok: false, error: "mode=value 需要 value（要填入的值，如 0 或 \"无\"）" };
  }
  const hasHeaders = args.has_headers !== false;
  const unmerge = args.unmerge !== false;
  const rangeArg = String(args.range || "").trim();
  const startRow = hasHeaders ? 1 : 0;
  const listMerges = unmerge && canListMergedAreas();

  let area = { address: "", rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 0 };
  let values: unknown[][] = [];
  let formulas: unknown[][] = [];
  let formats: unknown[][] = [];
  let merged: MergedArea[] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = rangeArg ? sheet.getRange(rangeArg) : sheet.getUsedRangeOrNullObject();
    range.load("address,rowIndex,columnIndex,rowCount,columnCount,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) throw new Error(`${sheetName} 是空表`);
    if (range.rowCount * range.columnCount > MAX_FILL_SCAN_CELLS) {
      throw new Error(`区域有 ${range.rowCount} 行 × ${range.columnCount} 列，超过单次 ${MAX_FILL_SCAN_CELLS} 个单元格，请用 range 缩小到要处理的列`);
    }
    area = {
      address: range.address.split("!").pop() as string,
      rowIndex: range.rowIndex,
      columnIndex: range.columnIndex,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
    range.load("values,formulas,numberFormat");
    const mergedAreas = listMerges ? range.getMergedAreasOrNullObject() : null;
    mergedAreas?.load("isNullObject");
    await context.sync();
    values = range.values as unknown[][];
    formulas = range.formulas as unknown[][];
    formats = range.numberFormat as unknown[][];
    if (mergedAreas && !mergedAreas.isNullObject) {
      mergedAreas.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
      await context.sync();
      merged = mergedAreas.areas.items
        .map((a) => {
          const r0 = Math.max(a.rowIndex, area.rowIndex);
          const c0 = Math.max(a.columnIndex, area.columnIndex);
          const r1 = Math.min(a.rowIndex + a.rowCount, area.rowIndex + area.rowCount);
          const c1 = Math.min(a.columnIndex + a.columnCount, area.columnIndex + area.columnCount);
          return { row: r0 - area.rowIndex, column: c0 - area.columnIndex, rows: r1 - r0, columns: c1 - c0 };
        })
        .filter((a) => a.rows > 0 && a.columns > 0);
    }
  });

  const letter = (c: number) => indexToColLetter(area.columnIndex + c);
  const headers = (values[0] || []).map((v, c) => (hasHeaders ? String(v ?? "").trim() : "") || letter(c));
  const specs: unknown[] = Array.isArray(args.columns) ? args.columns.filter((c: unknown) => c != null && c !== "") : [];
  const columns = specs.map((spec, i) => resolveColSpec(headers, spec, area.columnIndex, `columns[${i}]`));
  let areas = unmerge ? merged.filter((m) => isFillableMerge(m, startRow)) : [];
  if (columns.length) areas = areas.filter((m) => columns.some((c) => c >= m.column && c < m.column + m.columns));
  if (!columns.length && !areas.length) {
    if (!unmerge) return { ok: false, error: "unmerge=false 时需要用 columns 指定要填空白的列" };
    if (!listMerges) {
      return { ok: false, error: "这个 Excel 版本拿不到合并单元格的位置，请用 columns 指定要填充的列（如 [\"部门\"]）：会先取消这些列的合并，再按上方的值填满" };
    }
    return {
      ok: true,
      data: { range: `${sheetName}!${area.address}`, merged_areas: 0, filled: 0 },
      summary: `${sheetName}!${area.address} 里没有跨行的合并单元格；要填普通空白格请用 columns 指定列`,
    };
  }

  const plan = planBlankFill(values, { columns, mode, value: args.value, startRow, merged: areas, formulas });
  const bounds = fillBounds(plan.cells);
  if (!bounds) {
    return {
      ok: true,
      data: { range: `${sheetName}!${area.address}`, filled: 0, unfilled: plan.unfilled },
      summary: plan.unfilled ? `没有可填的空白：${plan.unfilled} 格上方/左侧也没有值` : "所选列没有需要填充的空白格",
    };
  }
  const blockRows = bounds.bottom - bounds.top + 1;
  const blockCols = bounds.right - bounds.left + 1;
  const blockAddr = `${letter(bounds.left)}${area.rowIndex + bounds.top + 1}:${letter(bounds.right)}${area.rowIndex + bounds.bottom + 1}`;
  if (blockRows * blockCols > MAX_SNAPSHOT_CELLS) {
    return {
      ok: false,
      error: `要填充的格分布在 ${blockAddr}（${blockRows * blockCols} 格），超过单次 ${MAX_SNAPSHOT_CELLS} 格的撤销快照上限。请用 range 分段处理或减少 columns`,
    };
  }
  // 旧版 Excel 拿不到合并位置：取消这些列的合并后，原来合并区域里的空白已经在 plan 里按上方补齐
  const unmergeColumns = unmerge && !listMerges ? columns : [];
  const perColumn: Record<string, number> = {};
  for (const [c, n] of Object.entries(plan.perColumn)) perColumn[headers[Number(c)]] = n;
  const modeText = mode === "value" ? `填入「${args.value}」` : mode === "left" ? "按左侧的值" : "按上方的值";

  if (ctx.settings.confirmDestructive) {
    const lines: string[] = [];
    if (areas.length) lines.push(`取消 ${areas.length} 个合并区域，每格补上原来的值（${plan.mergedCells} 格）`);
    if (unmergeColumns.length) lines.push(`先取消「${unmergeColumns.map((c) => headers[c]).join("、")}」列里的合并单元格`);
    for (const [name, n] of Object.entries(perColumn)) lines.push(`「${name}」${modeText}填充 ${n} 格`);
    if (plan.unfilled) lines.push(`${plan.unfilled} 格上方/左侧没有值，保持空白`);
    const sample = plan.cells
      .slice(0, 6)
      .map((c) => `${letter(c.column)}${area.rowIndex + c.row + 1} ← ${String(c.value)}`)
      .join("\n");
    const accepted = await ctx.confirm(
      `填充 ${sheetName}!${blockAddr} 中 ${plan.cells.length} 个空白单元格`,
      `${lines.join("\n")}\n\n示例:\n${sample}\n\n单元格内容可撤销；取消合并需要用任务检查点回退。`
    );
    if (!accepted) return { ok: false, error: "用户拒绝了填充空白。请不要重试，向用户询问意图。" };
  }

  const snapshot = formulas.slice(bounds.top, bounds.bottom + 1).map((row) => row.slice(bounds.left, bounds.right + 1));
  let targetChanged = false;
  let oldGrid: unknown[][] = [];
  let expectedValues: unknown[][] = [];
  let expectedFormulas: unknown[][] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const block = sheet.getRangeByIndexes(area.rowIndex + bounds.top, area.columnIndex + bounds.left, blockRows, blockCols);
    block.load("values,formulas");
    await context.sync();
    if (!gridsEqual(block.formulas as unknown[][], snapshot)) {
      targetChanged = true;
      return;
    }
    oldGrid = buildRestorableGrid(block.values as unknown[][], block.formulas as unknown[][]);
    checkToolAbort(ctx);
    for (const m of areas) {
      sheet.getRangeByIndexes(area.rowIndex + m.row, area.columnIndex + m.column, m.rows, m.columns).unmerge();
    }
    for (const c of unmergeColumns) {
      sheet.getRangeByIndexes(area.rowIndex + startRow, area.columnIndex + c, area.rowCount - startRow, 1).unmerge();
    }
    let pending = 0;
    for (const run of fillRuns(plan.cells)) {
      const target = sheet.getRangeByIndexes(area.rowIndex + run.row, area.columnIndex + run.column, run.cells.length, 1);
      target.formulas = run.cells.map((c) => [typeof c.value === "string" ? keepTextAsText(c.value) : c.value]) as any[][];
      // 填进来的值沿用来源格的数字格式（日期、金额），否则日期会显示成序列号
      const current = run.cells.map((c) => formats[c.row]?.[c.column]);
      const wanted = run.cells.map((c, i) => (c.from ? formats[c.from.row]?.[c.from.column] : undefined) ?? current[i] ?? "General");
      if (wanted.some((f, i) => normalizeNumberFormat(f) !== normalizeNumberFormat(current[i]))) {
        target.numberFormat = wanted.map((f) => [f]) as any[][];
      }
      if (++pending % FILL_SYNC_RUNS === 0) await context.sync();
    }
    await context.sync();
    block.load("values,formulas");
    await context.sync();
    expectedValues = block.values as unknown[][];
    expectedFormulas = block.formulas as unknown[][];
  });
  if (targetChanged) {
    return { ok: false, error: `${sheetName}!${blockAddr} 在确认后发生了变化，本次填充已停止。请重新查看后再操作。` };
  }

  const undoId = pushUndo({
    description: `填充空白 ${sheetName}!${blockAddr}`,
    sheet: sheetName,
    address: blockAddr,
    formulas: oldGrid as any[][],
    expectedValues: expectedValues as any[][],
    expectedFormulas: expectedFormulas as any[][],
  });
  pushOpLog({
    description: `填充空白 ${sheetName}!${blockAddr}（${plan.cells.length} 格${areas.length ? `，取消 ${areas.length} 个合并` : ""}）`,
    kind: "write",
    sheet: sheetName,
    address: blockAddr,
    before: gridPreview(oldGrid),
    after: gridPreview(expectedValues),
    ...(undoId ? { undoEntryId: undoId } : {}),
  });
  return {
    ok: true,
    data: {
      range: `${sheetName}!${blockAddr}`,
      filled: plan.cells.length,
      merged_areas: areas.length,
      merged_cells: plan.mergedCells,
      per_column: perColumn,
      unfilled: plan.unfilled,
      mode,
      ...(unmergeColumns.length ? { note: "这个 Excel 版本拿不到合并区域位置：已取消所选列的合并，并按上方的值补齐。" } : {}),
    },
    summary: `已填充 ${sheetName} 的 ${plan.cells.length} 个空白格${areas.length ? `（取消 ${areas.length} 个合并区域）` : ""}`,
    undoable: !!undoId,
  };
}

// ---------------- 从文本提取信息 ----------------

async function extractText(args: any, ctx: ToolExecContext): Promise<ToolResult> {
  const sheetName = String(args.sheet || "").trim();
  if (!sheetName) return { ok: false, error: "extract_text 需要 sheet" };
  if (isInternalSheet(sheetName)) return { ok: false, error: "不能处理检查点备份表。" };
  if (args.column == null || args.column === "") return { ok: false, error: "extract_text 需要 column（源列：表头或列字母）" };
  let targets: ExtractTarget[];
  try {
    targets = normalizeExtractTargets(args.targets);
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
  const hasHeaders = args.has_headers === false ? false : true;
  const layout = await loadSheetLayout(sheetName, args.range);
  const srcRel = resolveColSpec(layout.headers, args.column, layout.columnIndex, "column");
  const headerCount = hasHeaders ? 1 : 0;
  const dataRows = layout.rowCount - headerCount;
  if (dataRows < 1) return { ok: false, error: `${sheetName} 没有数据行` };
  if (dataRows > MAX_SPLIT_ROWS) return { ok: false, error: `有 ${dataRows} 行，超过 ${MAX_SPLIT_ROWS}，请缩小 range` };
  if (layout.rowCount * targets.length > MAX_SNAPSHOT_CELLS) {
    return { ok: false, error: `结果有 ${layout.rowCount * targets.length} 格，超过单次 ${MAX_SNAPSHOT_CELLS} 格，请减少 targets 或用 range 分段` };
  }

  let srcValues: any[][] = [];
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = args.range ? sheet.getRange(args.range) : sheet.getUsedRange();
    const col = used.getCell(0, srcRel).getAbsoluteResizedRange(layout.rowCount, 1);
    col.load("values");
    await context.sync();
    srcValues = col.values as any[][];
  });

  const today = new Date();
  const options = { today, all: args.all_matches === true, separator: args.separator == null ? "、" : String(args.separator) };
  const srcHeader = layout.headers[srcRel] || String(args.column);
  const stats = targets.map(() => ({ found: 0, missing: 0 }));
  let precisionLost = 0;
  let invalidIds = 0;
  const rows = srcValues.slice(headerCount).map((row) => {
    let lost = false;
    let invalid = false;
    const out = targets.map((target, i) => {
      const result = extractValue(row[0], target, options);
      if (result.missing) stats[i].missing++;
      else if (result.value !== "") stats[i].found++;
      if (result.precisionLost) lost = true;
      if (result.invalidId) invalid = true;
      // 手机号、身份证号这类数字串保持文本；日期留给 Excel 识别成真正的日期
      return typeof result.value === "string" && !DATE_TARGETS.has(target) ? keepTextAsText(result.value) : result.value;
    });
    if (lost) precisionLost++;
    if (invalid) invalidIds++;
    return out;
  });
  const labels = targets.map((t) => EXTRACT_TARGET_LABELS[t]);
  if (stats.every((s) => s.found === 0)) {
    return {
      ok: false,
      error: `在「${srcHeader}」的 ${dataRows} 行里没有找到${labels.join("、")}。请确认列选对了；格式特殊时可改用 run_script 写规则提取`,
    };
  }

  const existing = new Set(layout.headers);
  const destHeaders = targets.map((_, i) => {
    const given = Array.isArray(args.dest_headers) ? String(args.dest_headers[i] ?? "").trim() : "";
    return given || (existing.has(labels[i]) ? `${srcHeader}_${labels[i]}` : labels[i]);
  });
  const destStartAbs = parseDestStartColumn(args.dest_start_column, layout);
  const startCell = `${indexToColLetter(destStartAbs)}${layout.rowIndex + 1}`;
  const written = await writeCells({ sheet: sheetName, start_cell: startCell, values: hasHeaders ? [destHeaders, ...rows] : rows }, ctx, false);
  if (!written.ok) return written;
  const dateColumns = targets.map((t, i) => (DATE_TARGETS.has(t) ? i : -1)).filter((i) => i >= 0);
  if (dateColumns.length) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      for (const i of dateColumns) {
        sheet.getRangeByIndexes(layout.rowIndex + headerCount, destStartAbs + i, dataRows, 1).numberFormat = Array.from({ length: dataRows }, () => ["yyyy-mm-dd"]);
      }
      await context.sync();
    });
  }

  const warnings: string[] = [];
  targets.forEach((_, i) => {
    if (stats[i].missing) warnings.push(`「${destHeaders[i]}」有 ${stats[i].missing} 行没找到，已留空`);
  });
  if (precisionLost) {
    warnings.push(`有 ${precisionLost} 个身份证号被存成了数字，末几位已丢失：出生日期和年龄照算，性别和校验无法判断。请把源列设为文本后重新录入或导入`);
  }
  if (invalidIds) {
    warnings.push(`有 ${invalidIds} 个身份证号校验未通过（出生日期无效或校验位错误）${targets.includes("id_check") ? "，原因见「身份证校验」列" : "，可加 id_check 查看原因"}`);
  }
  const todayIso = toIsoDate(today.getFullYear(), today.getMonth() + 1, today.getDate());
  return {
    ok: true,
    data: {
      ...(written.data as object),
      column: srcHeader,
      dest_headers: destHeaders,
      found: Object.fromEntries(destHeaders.map((h, i) => [h, stats[i].found])),
      warnings,
      ...(targets.includes("age") ? { note: `年龄按今天 ${todayIso} 计算的周岁，是静态值，日后不会自动更新。` } : {}),
    },
    summary: `已从「${srcHeader}」提取 ${destHeaders.join("、")}（${Math.max(...stats.map((s) => s.found))} 行有结果）`,
    undoable: written.undoable,
  };
}

// ---------------- 分发器 ----------------

/**
 * 工具注册表：每个工具的执行函数和行为标记都在这一处声明。
 * - mutating：会改工作簿，执行前先建任务检查点；只有部分参数组合才写表时写成按参数判断的函数
 * - meta：交互/记忆类工具，不计入「实质性操作」（决定是否值得复盘学习）
 * - memory：长期记忆工具，关闭记忆时整组禁用
 * - readOnly：只读取、不改工作簿也没有其他副作用；同批前面的工具失败时它照常执行，重复调用会被提示
 * 参数 schema 在 definitions.ts；两边的名称一致性由 tests/agent-runtime.cjs 校验。
 */
export interface ToolSpec {
  run: (args: any, ctx: ToolExecContext) => ToolResult | Promise<ToolResult>;
  mutating?: boolean | ((args: any) => boolean);
  meta?: boolean;
  memory?: boolean;
  readOnly?: boolean;
}

export const TOOL_REGISTRY: Record<string, ToolSpec> = {
  get_workbook_overview: { run: () => getWorkbookOverview(), readOnly: true },
  read_range: { run: (a) => readRange(a), readOnly: true },
  write_range: { run: (a, c) => writeCells(a, c, false), mutating: true },
  set_formulas: { run: (a, c) => writeCells(a, c, true), mutating: true },
  create_sheet: { run: (a) => createSheet(a), mutating: true },
  delete_sheet: { run: (a, c) => deleteSheet(a, c), mutating: true },
  clear_range: { run: (a, c) => clearRange(a, c), mutating: true },
  insert_rows: { run: (a, c) => insertRows(a, c), mutating: true },
  delete_rows: { run: (a, c) => deleteRows(a, c), mutating: true },
  insert_columns: { run: (a, c) => insertColumns(a, c), mutating: true },
  delete_columns: { run: (a, c) => deleteColumns(a, c), mutating: true },
  format_range: { run: (a) => formatRange(a), mutating: true },
  sort_range: { run: (a) => sortRange(a), mutating: true },
  find_cells: { run: (a) => findCells(a), readOnly: true },
  create_chart: { run: (a) => createChart(a), mutating: true },
  list_external_files: { run: () => toolListExternalFiles(), meta: true, readOnly: true },
  read_external_sheet: { run: (a) => toolReadExternalSheet(a), readOnly: true },
  import_external_sheet: { run: (a) => toolImportExternalSheet(a), mutating: true },
  export_file: { run: (a) => toolExportFile(a) },
  ask_user_choice: { run: (a, c) => askUserChoice(a, c), meta: true },
  inspect_range: { run: (a) => inspectRange(a), readOnly: true },
  list_tables: { run: () => toolListTables(), readOnly: true },
  create_table: { run: (a) => toolCreateTable(a), mutating: true },
  apply_autofilter: { run: (a) => toolApplyFilter(a), mutating: true },
  clear_autofilter: { run: (a) => toolClearFilter(a), mutating: true },
  create_pivot_table: { run: (a) => toolCreatePivot(a), mutating: true },
  suggest_followups: { run: (a) => suggestFollowups(a), meta: true },
  remember: { run: (a) => toolRemember(a), meta: true, memory: true },
  forget: { run: (a) => toolForget(a), meta: true, memory: true },
  list_memories: { run: () => toolListMemories(), meta: true, memory: true, readOnly: true },
  save_skill: { run: (a) => toolSaveSkill(a), meta: true },
  list_skills: { run: () => toolListSkills(), meta: true, readOnly: true },
  audit_data: { run: (a) => toolAuditData(a), mutating: true },
  diagnose_errors: { run: (a) => toolDiagnoseErrors(a), readOnly: true },
  copy_range: { run: (a, c) => copyRange(a, c), mutating: true },
  fill_down: { run: (a, c) => fillDown(a, c), mutating: true },
  find_replace: { run: (a, c) => findReplace(a, c), mutating: true },
  remove_duplicates: { run: (a, c) => removeDuplicates(a, c), mutating: true },
  rename_sheet: { run: (a) => renameSheet(a), mutating: true },
  copy_sheet: { run: (a) => copySheet(a), mutating: true },
  column_stats: { run: (a) => columnStats(a), readOnly: true },
  apply_conditional_format: { run: (a) => applyConditionalFormat(a), mutating: true },
  add_lookup_column: { run: (a, c) => addLookupColumn(a, c), mutating: true },
  split_column: { run: (a, c) => splitColumn(a, c), mutating: true },
  combine_columns: { run: (a, c) => combineColumns(a, c), mutating: true },
  freeze_panes: { run: (a) => freezePanes(a), mutating: true },
  clean_column: { run: (a, c) => cleanColumn(a, c), mutating: true },
  convert_to_values: { run: (a, c) => convertToValues(a, c), mutating: true },
  polish_range: { run: (a) => polishRange(a), mutating: true },
  delete_blank_rows: { run: (a, c) => deleteBlankRows(a, c), mutating: true },
  group_summary: { run: (a, c) => toolGroupSummary(a, c), mutating: true },
  reconcile_tables: { run: (a, c) => toolReconcileTables(a, c), mutating: true },
  stack_sheets: { run: (a, c) => toolStackSheets(a, c), mutating: true },
  split_by_column: { run: (a, c) => toolSplitByColumn(a, c), mutating: true },
  set_data_validation: { run: (a) => setDataValidation(a), mutating: true },
  set_visibility: { run: (a) => setVisibility(a), mutating: true },
  unpivot_columns: { run: (a, c) => toolUnpivotColumns(a, c), mutating: true },
  fill_blank_cells: { run: (a, c) => fillBlankCells(a, c), mutating: true },
  extract_text: { run: (a, c) => extractText(a, c), mutating: true },
  compare_sheets: { run: (a, c) => toolCompareSheets(a, c), mutating: true },
  set_page_layout: { run: (a) => toolSetPageLayout(a), mutating: true },
  create_sheet_index: { run: (a) => toolCreateSheetIndex(a), mutating: true },
  // 不带 output 时只在沙箱里计算、不改工作簿，不必为它建检查点
  run_script: { run: (a, c) => toolRunScript(a, c, (sheet, cell, grid) => writeScriptGrid(sheet, cell, grid, c)), mutating: (a) => !!a?.output },
};

/** 这次调用会不会改工作簿（决定是否先建检查点、成功后是否重置失败计数） */
export function isMutatingCall(name: string, args: unknown): boolean {
  const flag = TOOL_REGISTRY[name]?.mutating;
  return typeof flag === "function" ? flag(args) : !!flag;
}

function writeScriptGrid(sheet: string, startCell: string, grid: unknown[][], ctx: ToolExecContext): Promise<ToolResult> {
  const cells = grid.length * (grid[0]?.length ?? 0);
  if (cells > MAX_SNAPSHOT_CELLS) {
    return Promise.resolve({
      ok: false,
      error: `结果有 ${cells} 个单元格，写入已有位置单次最多 ${MAX_SNAPSHOT_CELLS} 个（需要保留撤销快照）。请改用 output.new_sheet 写到新 Sheet。`,
    });
  }
  return writeCells({ sheet, start_cell: startCell, values: grid }, ctx, false);
}


function checkToolAbort(ctx: ToolExecContext): void {
  if (ctx.signal?.aborted) throw new Error("用户已停止任务，未继续执行操作。");
}

/** 用户拒绝确认时附了说明：替换工具里「向用户询问意图」的通用措辞，让模型按说明调整后继续 */
function withRejectNotes(result: ToolResult, notes: string[]): ToolResult {
  const said = notes.map((note) => `「${note}」`).join("；");
  const instruction = `用户拒绝了这一步，并说明：${said}。请按用户的说明调整做法后继续，不要原样重试，也不必再问一遍意图。`;
  if (!result.ok) return { ...result, error: instruction };
  const data = result.data && typeof result.data === "object" && !Array.isArray(result.data)
    ? { ...(result.data as Record<string, unknown>), user_note: instruction }
    : { result: result.data, user_note: instruction };
  return { ...result, data };
}

export async function executeTool(name: string, argsJson: string, ctx: ToolExecContext): Promise<ToolResult> {
  const parsedArgs = parseToolArgumentsJson(argsJson);
  if (!parsedArgs.ok) return { ok: false, error: `${parsedArgs.error}。本次未执行，请重新发出完整、合法的 JSON 参数。` };
  let args: any = parsedArgs.value;
  try {
    checkToolAbort(ctx);
    const definition = TOOL_DEFINITIONS.find((tool) => tool.function.name === name);
    if (!definition) return { ok: false, error: `未知工具: ${name}` };
    // 先拆开 "表名!地址"，这样只写了带表名地址、漏了 sheet 的调用也能通过必填校验
    if (args && typeof args === "object" && !Array.isArray(args)) {
      splitSheetQualifiedAddresses(args, (definition.function.parameters.properties as Record<string, unknown>) || {});
    }
    try {
      // 顶层拼错字段不应被默默忽略，例如把 source_sheet 写成 sheet。
      args = normalizeToolArguments(args, { ...definition.function.parameters, additionalProperties: false });
    } catch (e: any) {
      const schema = definition.function.parameters;
      return { ok: false, error: `${e?.message || e}。${name} 必填字段：${(schema.required as string[] || []).join("、") || "无"}；允许字段：${Object.keys(schema.properties as object || {}).join("、") || "无"}。请修正参数后重试；本次未执行。` };
    }
    if (!ctx.settings.useMemory && TOOL_REGISTRY[name]?.memory) {
      return { ok: false, error: "长期记忆已关闭，本次不会读取或写入记忆。" };
    }
    const originalCtx = ctx;
    const rejectNotes: string[] = [];
    ctx = {
      ...originalCtx,
      confirm: async (title, detail) => {
        checkToolAbort(originalCtx);
        let accepted: boolean;
        if (originalCtx.confirmDetailed) {
          const answer = await originalCtx.confirmDetailed(title, detail, { tool: name });
          accepted = answer.accepted;
          if (!accepted && answer.note?.trim()) rejectNotes.push(answer.note.trim());
        } else {
          accepted = await originalCtx.confirm(title, detail);
        }
        checkToolAbort(originalCtx);
        return accepted;
      },
      choose: async (...params) => {
        checkToolAbort(originalCtx);
        const answer = await originalCtx.choose(...params);
        checkToolAbort(originalCtx);
        return answer;
      },
    };
    if (ctx.settings.autoCheckpoint !== false && isMutatingCall(name, args)) {
      await ensureCheckpoint();
    }
    checkToolAbort(ctx);
    const spec = TOOL_REGISTRY[name];
    if (!spec) return { ok: false, error: `未知工具: ${name}` };
    const result = await spec.run(args, ctx);
    const final = rejectNotes.length ? withRejectNotes(result, rejectNotes) : result;
    return final.ok ? final : withSheetHint(name, args, final);
  } catch (e: any) {
    const msg = e?.debugInfo ? `${e.message} (${JSON.stringify(e.debugInfo)})` : e?.message || String(e);
    return withSheetHint(name, args, { ok: false, error: `执行 ${name} 出错: ${msg}` });
  }
}

/** 这些工具的 sheet 参数指附加文件里的表，不是当前工作簿 */
const EXTERNAL_SHEET_TOOLS = new Set(["list_external_files", "read_external_sheet", "import_external_sheet"]);
const NOT_FOUND_RE = /ItemNotFound|找不到工作表|does ?n[o']t exist|不存在/i;

/** 表名写错时把现有表名和最接近的候选附在错误里，省掉模型再查一轮结构 */
async function withSheetHint(name: string, args: unknown, result: ToolResult): Promise<ToolResult> {
  if (result.ok || EXTERNAL_SHEET_TOOLS.has(name) || !NOT_FOUND_RE.test(result.error || "")) return result;
  // 工具自己已经带上候选表名时不再重复
  if ((result.error || "").includes("现有工作表")) return result;
  try {
    const names = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((s) => s.name).filter((n) => !isInternalSheet(n));
    });
    const hint = missingSheetHint(args, names);
    return hint ? { ...result, error: `${result.error}\n${hint}` } : result;
  } catch {
    return result;
  }
}

// ---------------- 快速上下文（发消息时附带） ----------------

export interface QuickContext {
  sheetNames: string[];
  activeSheet: string;
  selectionAddress: string;
  selectionPreview?: string;
  selectionHasFormula?: boolean;
  selectionHasError?: boolean;
  selectionFormula?: string;
  selectionErrors?: string[];
  /** 各表已用区域简述（如 "A1:H501，500 行×8 列" / "空"），只在请求时读取 */
  sheetSizes?: Record<string, string>;
}

const MAX_SIZED_SHEETS = 30;

/** 一次 sync 读出每张表的已用区域；失败时返回 undefined，不影响其余上下文。 */
async function loadSheetSizes(context: Excel.RequestContext, sheets: Excel.Worksheet[]): Promise<Record<string, string> | undefined> {
  try {
    const used = sheets.slice(0, MAX_SIZED_SHEETS).map((ws) => {
      const range = ws.getUsedRangeOrNullObject(true);
      range.load("address,rowCount,columnCount,isNullObject");
      return { name: ws.name, range };
    });
    await context.sync();
    const sizes: Record<string, string> = {};
    for (const { name, range } of used) {
      sizes[name] = range.isNullObject
        ? "空"
        : `${range.address.split("!").pop()}，${range.rowCount} 行×${range.columnCount} 列`;
    }
    return sizes;
  } catch {
    return undefined;
  }
}

export async function getQuickContext(includeSelectionValues: boolean, withSheetSizes = false): Promise<QuickContext> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    const sel = context.workbook.getSelectedRange();
    sel.load("address,rowCount,columnCount");
    await context.sync();

    const cells = sel.rowCount * sel.columnCount;
    let selectionPreview: string | undefined;
    let selectionHasFormula: boolean | undefined;
    let selectionHasError: boolean | undefined;
    let selectionFormula: string | undefined;
    let selectionErrors: string[] | undefined;

    if (cells <= 40) {
      sel.load("values,formulas");
      await context.sync();
      const analysis = analyzeGrid(sel.values as any[][], sel.formulas as any[][]);
      selectionHasFormula = analysis.hasFormula;
      selectionHasError = analysis.hasError;
      selectionFormula = analysis.formulaPreview;
      selectionErrors = analysis.errors;
      if (includeSelectionValues) {
        selectionPreview = (sel.values as any[][])
          .map((row) => row.map((v) => (v === "" ? "·" : String(v).slice(0, 20))).join(" | "))
          .join("\n");
      }
    } else if (includeSelectionValues && cells <= 200) {
      sel.load("values");
      await context.sync();
      selectionPreview = (sel.values as any[][])
        .map((row) => row.map((v) => (v === "" ? "·" : String(v).slice(0, 20))).join(" | "))
        .join("\n");
    }
    const visibleSheets = sheets.items.filter((s) => !isInternalSheet(s.name));
    const sheetSizes = withSheetSizes ? await loadSheetSizes(context, visibleSheets) : undefined;
    return {
      sheetNames: visibleSheets.map((s) => s.name),
      activeSheet: active.name,
      selectionAddress: sel.address,
      selectionPreview,
      selectionHasFormula,
      selectionHasError,
      selectionFormula,
      selectionErrors,
      ...(sheetSizes ? { sheetSizes } : {}),
    };
  });
}

/** 切到指定 Sheet 并选中区域（回复里的地址链接用）；不传地址只切换 Sheet */
export async function selectRangeInWorkbook(sheet: string, address?: string): Promise<void> {
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItemOrNullObject(sheet);
    ws.load("isNullObject,visibility");
    await context.sync();
    if (ws.isNullObject) throw new Error(`找不到工作表「${sheet}」，可能已被删除或改名`);
    if (ws.visibility !== Excel.SheetVisibility.visible) throw new Error(`工作表「${sheet}」是隐藏的，取消隐藏后才能定位`);
    ws.activate();
    if (address) ws.getRange(address).select();
    await context.sync();
  });
}

export type ColumnKind = "empty" | "number" | "date" | "text" | "mixed";

export interface SelectionProfile {
  sheet: string;
  address: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  isRange: boolean;
  headers: string[];
  columnKinds: ColumnKind[];
  hasFormula: boolean;
  hasError: boolean;
  emptyRatio: number;
  sheetNames: string[];
}

function guessValueKind(v: unknown): ColumnKind | "skip" {
  if (v === "" || v === null || v === undefined) return "empty";
  if (v instanceof Date) return "date";
  if (typeof v === "number") {
    if (v > 20000 && v < 80000 && Number.isInteger(v)) return "date";
    return "number";
  }
  const s = String(v).trim();
  if (!s) return "empty";
  if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(s) || /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}/.test(s)) return "date";
  if (/^[#](N\/A|REF!|VALUE!|DIV\/0!|NAME\?)/i.test(s)) return "skip";
  if (/^-?\d+(\.\d+)?%?$/.test(s.replace(/,/g, ""))) return "number";
  return "text";
}

function majorityKind(values: unknown[]): ColumnKind {
  const counts: Record<string, number> = {};
  for (const v of values) {
    const k = guessValueKind(v);
    if (k === "skip" || k === "empty") continue;
    counts[k] = (counts[k] || 0) + 1;
  }
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  if (entries.length === 0) return "empty";
  if (entries.length > 1 && entries[0][1] < entries.reduce((s, [, n]) => s + n, 0) * 0.7) return "mixed";
  return entries[0][0] as ColumnKind;
}

/** 读取当前选区的轻量画像，供空状态推荐选项使用 */
export async function getSelectionProfile(): Promise<SelectionProfile> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    const sel = context.workbook.getSelectedRange();
    sel.load("address,rowCount,columnCount");
    await context.sync();

    const rowCount = sel.rowCount;
    const columnCount = sel.columnCount;
    const cellCount = rowCount * columnCount;
    const address = sel.address.split("!").pop() || sel.address;
    const sampleRows = Math.min(rowCount, 6);
    const sampleCols = Math.min(columnCount, 16);
    const sample = sel.getCell(0, 0).getAbsoluteResizedRange(sampleRows, sampleCols);
    sample.load("values,formulas");
    await context.sync();

    const values = (sample.values as any[][]) || [];
    const formulas = (sample.formulas as any[][]) || [];
    const analysis = analyzeGrid(values, formulas);

    let empty = 0;
    let total = 0;
    for (const row of values) {
      for (const v of row) {
        total++;
        if (v === "" || v === null || v === undefined) empty++;
      }
    }

    const first = values[0] || [];
    const rest = values.slice(1);
    const firstTextish = first.filter((v) => guessValueKind(v) === "text").length;
    const hasHeaders = rest.length > 0 && firstTextish >= Math.ceil(first.length * 0.5);

    const headers = (hasHeaders ? first : first.map((_, i) => `第${i + 1}列`)).map((v, i) => {
      const s = String(v ?? "").trim();
      return s || `第${i + 1}列`;
    });

    const columnKinds: ColumnKind[] = [];
    for (let c = 0; c < sampleCols; c++) {
      const colVals = (hasHeaders ? rest : values).map((r) => r[c]);
      columnKinds.push(majorityKind(colVals));
    }

    return {
      sheet: active.name,
      address,
      rowCount,
      columnCount,
      cellCount,
      isRange: cellCount >= 2,
      headers,
      columnKinds,
      hasFormula: analysis.hasFormula,
      hasError: analysis.hasError,
      emptyRatio: total ? empty / total : 0,
      sheetNames: sheets.items.filter((s) => !isInternalSheet(s.name)).map((s) => s.name),
    };
  });
}
