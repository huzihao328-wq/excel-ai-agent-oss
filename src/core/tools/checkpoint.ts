/* global Excel, Office */
/*
 * 任务检查点：每轮任务第一次改表前备份整本，失败或不满意时可「恢复到任务开始」。
 *
 * 两种存储：
 * - "file"（默认）：用 getFileAsync 取整本文件快照存进任务窗格的 IndexedDB，
 *   恢复时用 insertWorksheetsFromBase64 插回。工作簿里不留任何备份数据。
 * - "sheets"（旧版 Excel 不支持 ExcelApi 1.13 时的兜底）：把每张表复制成 VeryHidden 备份表。
 *   备份会随文件保存，所以只在没有别的办法时使用。
 */
import {
  deleteCheckpointFile,
  getCheckpointFile,
  hasCheckpointFile,
  pruneCheckpointFiles,
  putCheckpointFile,
} from "../checkpointStore";

export const CP_PREFIX = "__AI_CP_";
export const RS_PREFIX = "__AI_RS_";
/** 恢复过程中临时改名的用户表，失败时仍保留数据，成功后再删 */
const OLD_PREFIX = "__AI_OLD_";

const SETTINGS_KEY = "excel-ai-agent-checkpoint";

export interface CheckpointSheet {
  originalName: string;
  backupName: string;
  position: number;
  visibility: string;
  tables: string[];
}

export interface TaskCheckpoint {
  id: string;
  /** 备份存放位置；缺省为旧版的隐藏备份表 */
  storage?: "file" | "sheets";
  createdAt: number;
  label: string;
  activeSheet: string;
  sheets: CheckpointSheet[];
  formulas?: CheckpointFormulas[];
}

interface CheckpointFormulas {
  sheet: string;
  range: string;
  formulas: string[][];
}

async function captureFormulas(context: Excel.RequestContext, sheet: Excel.Worksheet, name: string): Promise<CheckpointFormulas[]> {
  const cells = sheet.getRange().getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
  cells.load("isNullObject");
  await context.sync();
  if (cells.isNullObject) return [];
  cells.areas.load("items/address,items/formulas");
  await context.sync();
  return cells.areas.items.map((area) => ({
    sheet: name,
    range: area.address.slice(area.address.lastIndexOf("!") + 1),
    formulas: area.formulas.map((row) => row.map(String)),
  }));
}

export type CheckpointEvent = "hydrated" | "creating" | "created" | "failed" | "cleanup-failed" | "restored" | "discarded";

type Listener = (cp: TaskCheckpoint | null, event: CheckpointEvent) => void;

let current: TaskCheckpoint | null = null;
let sessionLabel = "";
let sessionSnapshotted = false;
let sessionFailed = false;
let listeners: Listener[] = [];

export function isInternalSheet(name: string): boolean {
  return name.startsWith(CP_PREFIX) || name.startsWith(RS_PREFIX) || name.startsWith(OLD_PREFIX);
}

function isDisposableCheckpointSheet(name: string): boolean {
  return name.startsWith(CP_PREFIX) || name.startsWith(RS_PREFIX);
}

export function getCheckpoint(): TaskCheckpoint | null {
  return current;
}

export function onCheckpointChange(fn: Listener): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((x) => x !== fn);
  };
}

function notify(event: CheckpointEvent) {
  listeners.forEach((fn) => fn(current, event));
}

export function beginCheckpointSession(label: string): void {
  sessionLabel = String(label || "任务")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 40);
  sessionSnapshotted = false;
  sessionFailed = false;
}

export function checkpointRestoreHint(enabled: boolean): string {
  return enabled
    ? "单元格级撤销无法还原此次删除。若已创建任务检查点，可用底部「恢复到任务开始」整段回退。"
    : "删除后无法通过本插件撤销。";
}

function makeId(): string {
  return `${Date.now().toString(36).slice(-4)}${Math.random().toString(36).slice(2, 4)}`;
}

function backupSheetName(id: string, index: number): string {
  return `${CP_PREFIX}${id}_${String(index).padStart(2, "0")}`;
}

function stagingSheetName(attemptId: string, index: number): string {
  return `${RS_PREFIX}${attemptId}_${String(index).padStart(2, "0")}`;
}

function parkedSheetName(attemptId: string, index: number): string {
  return `${OLD_PREFIX}${attemptId}_${String(index).padStart(2, "0")}`;
}

function persist(cp: TaskCheckpoint | null) {
  try {
    // 文件快照的公式已随快照存进 IndexedDB，不再塞进工作簿自己的设置里
    const stored = cp && cp.storage === "file" ? { ...cp, formulas: undefined } : cp;
    if (stored) Office.context.document.settings.set(SETTINGS_KEY, JSON.stringify(stored));
    else Office.context.document.settings.remove(SETTINGS_KEY);
    Office.context.document.settings.saveAsync();
  } catch {
    /* 非 Excel 环境 */
  }
}

function readPersisted(): TaskCheckpoint | null {
  try {
    const raw = Office.context.document.settings.get(SETTINGS_KEY);
    if (typeof raw !== "string" || !raw) return null;
    const parsed = JSON.parse(raw) as TaskCheckpoint;
    if (!parsed?.id || !Array.isArray(parsed.sheets)) return null;
    return parsed;
  } catch {
    return null;
  }
}

function asVisibility(value: string): Excel.SheetVisibility {
  if (value === "Hidden" || value === Excel.SheetVisibility.hidden) return Excel.SheetVisibility.hidden;
  if (value === "VeryHidden" || value === Excel.SheetVisibility.veryHidden) return Excel.SheetVisibility.veryHidden;
  return Excel.SheetVisibility.visible;
}

async function deleteSheetsNamed(names: Set<string>): Promise<void> {
  if (names.size === 0) return;
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    for (const ws of sheets.items) {
      if (names.has(ws.name)) {
        // Office.js cannot delete a VeryHidden worksheet directly.
        ws.visibility = Excel.SheetVisibility.hidden;
        ws.delete();
      }
    }
    await context.sync();
  });
}

async function deleteInternalSheets(except: Set<string> = new Set()): Promise<void> {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    for (const ws of sheets.items) {
      if (isDisposableCheckpointSheet(ws.name) && !except.has(ws.name)) {
        ws.visibility = Excel.SheetVisibility.hidden;
        ws.delete();
      }
    }
    await context.sync();
  });
}

/** 当前宿主能否做「文件快照」检查点：需要 getFileAsync 取整本、ExcelApi 1.13 插回。 */
export function supportsFileCheckpoint(): boolean {
  try {
    const doc = Office.context.document as unknown as { getFileAsync?: unknown };
    return (
      typeof doc?.getFileAsync === "function" &&
      typeof indexedDB !== "undefined" &&
      !!Office.context.requirements?.isSetSupported?.("ExcelApi", "1.13")
    );
  } catch {
    return false;
  }
}

const FILE_SLICE_BYTES = 4 * 1024 * 1024;

/** 读出当前工作簿（含未保存改动）的完整 .xlsx 字节。 */
function readWorkbookFile(): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: FILE_SLICE_BYTES }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(res.error?.message || "读取工作簿文件失败"));
        return;
      }
      const file = res.value;
      const parts: Uint8Array[] = [];
      let total = 0;
      const finish = (err?: Error) => {
        file.closeAsync(() => {
          if (err) return reject(err);
          const out = new Uint8Array(total);
          let offset = 0;
          for (const part of parts) {
            out.set(part, offset);
            offset += part.length;
          }
          resolve(out.buffer);
        });
      };
      const next = (index: number) => {
        if (index >= file.sliceCount) return finish();
        file.getSliceAsync(index, (slice) => {
          if (slice.status !== Office.AsyncResultStatus.Succeeded) {
            finish(new Error(slice.error?.message || "读取工作簿文件分片失败"));
            return;
          }
          const bytes = new Uint8Array(slice.value.data as ArrayLike<number>);
          parts.push(bytes);
          total += bytes.length;
          next(index + 1);
        });
      };
      next(0);
    });
  });
}

function bytesToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CHUNK)));
  }
  return btoa(binary);
}

/** 文件快照：工作簿里只读不写，备份整体存进 IndexedDB。 */
async function snapshotWorkbookFile(label: string): Promise<TaskCheckpoint> {
  const id = makeId();
  const meta = await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/position,items/visibility");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();
    const sources = sheets.items
      .filter((s) => !isInternalSheet(s.name))
      .sort((a, b) => a.position - b.position);
    if (sources.length === 0) throw new Error("工作簿里没有可备份的 Sheet");
    for (const src of sources) src.tables.load("items/name");
    await context.sync();
    const formulas: CheckpointFormulas[] = [];
    for (const src of sources) formulas.push(...await captureFormulas(context, src, src.name));
    return {
      activeSheet: active.name,
      formulas,
      sheets: sources.map((src) => ({
        originalName: src.name,
        backupName: "",
        position: src.position,
        visibility: String(src.visibility),
        tables: src.tables.items.map((t) => t.name),
      })),
    };
  });
  const bytes = await readWorkbookFile();
  const createdAt = Date.now();
  await putCheckpointFile({ id, createdAt, bytes, formulas: meta.formulas });
  return {
    id,
    storage: "file",
    createdAt,
    label,
    activeSheet: meta.activeSheet,
    sheets: meta.sheets,
    formulas: meta.formulas,
  };
}

async function snapshotWorkbook(label: string): Promise<TaskCheckpoint> {
  const id = makeId();
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/position,items/visibility");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();

    const sources = sheets.items
      .filter((s) => !isInternalSheet(s.name))
      .sort((a, b) => a.position - b.position);
    if (sources.length === 0) throw new Error("工作簿里没有可备份的 Sheet");

    for (const src of sources) src.tables.load("items/name");
    await context.sync();

    const mapped: CheckpointSheet[] = [];
    const formulas: CheckpointFormulas[] = [];
    for (let i = 0; i < sources.length; i++) {
      const src = sources[i];
      // Capture original expressions before worksheet copies/renames can
      // redirect references, or later task edits delete their source sheets.
      formulas.push(...await captureFormulas(context, src, src.name));
      const backupName = backupSheetName(id, i);
      const copy = src.copy(Excel.WorksheetPositionType.end);
      copy.name = backupName;
      copy.visibility = Excel.SheetVisibility.veryHidden;
      mapped.push({
        originalName: src.name,
        backupName,
        position: src.position,
        visibility: String(src.visibility),
        tables: src.tables.items.map((t) => t.name),
      });
      await context.sync();
    }

    try {
      context.workbook.worksheets.getItem(active.name).activate();
      await context.sync();
    } catch {
      /* 保持当前活动表即可 */
    }

    return {
      id,
      storage: "sheets",
      createdAt: Date.now(),
      label,
      activeSheet: active.name,
      sheets: mapped,
      formulas,
    };
  });
}

/** 首次改表前备份整本；同一轮任务只做一次。失败不阻断任务。 */
export async function ensureCheckpoint(): Promise<TaskCheckpoint | null> {
  if (sessionSnapshotted) return current;
  if (sessionFailed) return current;

  notify("creating");
  const previous = current;
  try {
    const next = supportsFileCheckpoint()
      ? await snapshotWorkbookFile(sessionLabel || "任务")
      : await snapshotWorkbook(sessionLabel || "任务");
    current = next;
    sessionSnapshotted = true;
    persist(next);
    notify("created");
    const keep = new Set(next.sheets.map((s) => s.backupName).filter(Boolean));
    try {
      if (previous?.storage === "file") await deleteCheckpointFile(previous.id);
      if (next.storage === "file") void pruneCheckpointFiles(next.id);
      if (previous && previous.storage !== "file") {
        await deleteSheetsNamed(new Set(previous.sheets.map((s) => s.backupName).filter((n) => !keep.has(n))));
      }
      // 文件快照模式下 keep 为空，这一步会顺带清掉旧版本留在工作簿里的隐藏备份表
      await deleteInternalSheets(keep);
    } catch {
      // The new checkpoint is valid even if removing an older backup fails.
      notify("cleanup-failed");
    }
    return next;
  } catch (e: any) {
    sessionFailed = true;
    notify("failed");
    console.warn("创建任务检查点失败:", e?.message || e);
    return current;
  }
}

async function restoreFromFile(cp: TaskCheckpoint): Promise<void> {
  const record = await getCheckpointFile(cp.id);
  if (!record) throw new Error("检查点的备份文件已丢失（可能被浏览器存储清理），无法整段恢复");
  const base64 = bytesToBase64(record.bytes);
  const formulas = record.formulas || [];
  const attemptId = makeId();

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // 先把现有用户表改成临时名：插回的快照表才能用回原名；插回失败时可以原样改回。
    const parked: { parkedName: string; name: string }[] = [];
    let parkIdx = 0;
    for (const ws of sheets.items) {
      if (isInternalSheet(ws.name)) continue;
      const parkedName = parkedSheetName(attemptId, parkIdx++);
      parked.push({ parkedName, name: ws.name });
      ws.name = parkedName;
    }
    await context.sync();

    try {
      context.workbook.insertWorksheetsFromBase64(base64, {
        positionType: Excel.WorksheetPositionType.end,
        sheetNamesToInsert: cp.sheets.map((s) => s.originalName),
      });
      await context.sync();
    } catch (err) {
      for (const p of parked) context.workbook.worksheets.getItem(p.parkedName).name = p.name;
      await context.sync();
      throw new Error(
        "恢复检查点失败，工作簿未改动。" + (err instanceof Error && err.message ? `（${err.message}）` : "")
      );
    }

    sheets.load("items/name");
    await context.sync();
    const present = new Set(sheets.items.map((s) => s.name));
    const missing = cp.sheets.filter((s) => !present.has(s.originalName));
    if (missing.length) {
      throw new Error(
        `恢复后缺少「${missing.map((s) => s.originalName).join("、")}」，当前表已暂存为 ${OLD_PREFIX}*，数据未删除。请手动检查。`
      );
    }

    // 先保证有可见表，再删暂存表（Excel 不允许删掉最后一张可见表）
    for (const s of cp.sheets) {
      context.workbook.worksheets.getItem(s.originalName).visibility = asVisibility(s.visibility);
    }
    if (!cp.sheets.some((s) => asVisibility(s.visibility) === Excel.SheetVisibility.visible)) {
      context.workbook.worksheets.getItem(cp.sheets[0].originalName).visibility = Excel.SheetVisibility.visible;
    }
    await context.sync();
    for (const p of parked) {
      const ws = context.workbook.worksheets.getItem(p.parkedName);
      // VeryHidden 的表不能直接删
      ws.visibility = Excel.SheetVisibility.hidden;
      ws.delete();
    }
    await context.sync();

    // 暂存表还在时，同名表格会让插回的表格被自动改名；这里改回原名
    for (const s of cp.sheets) {
      const ws = context.workbook.worksheets.getItem(s.originalName);
      ws.tables.load("items/name");
      await context.sync();
      const count = Math.min(ws.tables.items.length, s.tables.length);
      for (let i = 0; i < count; i++) {
        if (s.tables[i] && ws.tables.items[i].name !== s.tables[i]) ws.tables.items[i].name = s.tables[i];
      }
      await context.sync();
    }

    for (const s of [...cp.sheets].sort((a, b) => a.position - b.position)) {
      context.workbook.worksheets.getItem(s.originalName).position = s.position;
    }
    await context.sync();

    // 跨表引用在插回时可能被解析成外部链接或 #REF!，按快照时的原公式重写一遍
    for (const block of formulas) {
      context.workbook.worksheets.getItem(block.sheet).getRange(block.range).formulas = block.formulas;
    }
    await context.sync();

    try {
      context.workbook.worksheets.getItem(cp.activeSheet).activate();
      await context.sync();
    } catch {
      /* 原活动表可能当时就是隐藏表 */
    }
  });
}

export async function restoreCheckpoint(): Promise<string> {
  const cp = current;
  if (!cp) throw new Error("没有可恢复的任务检查点");
  if (cp.storage === "file") {
    await restoreFromFile(cp);
    await discardCheckpoint(false);
    notify("restored");
    return `已恢复到任务开始时的工作簿（${cp.sheets.length} 个 Sheet）`;
  }
  const attemptId = makeId();

  await deleteInternalSheets(new Set(cp.sheets.map((s) => s.backupName)));

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const existing = new Set(sheets.items.map((s) => s.name));
    for (const s of cp.sheets) {
      if (!existing.has(s.backupName)) {
        throw new Error(`检查点里「${s.originalName}」的备份已丢失，无法整段恢复`);
      }
    }

    const formulas = cp.formulas ? [...cp.formulas] : [];
    if (!cp.formulas) {
      // Compatibility with checkpoints created before formula snapshots.
      for (const s of cp.sheets) {
        formulas.push(...await captureFormulas(context, sheets.getItem(s.backupName), s.originalName));
      }
    }

    const staging: { originalName: string; stagingName: string; visibility: string; position: number; tables: string[] }[] =
      [];
    for (let i = 0; i < cp.sheets.length; i++) {
      const s = cp.sheets[i];
      const backup = context.workbook.worksheets.getItem(s.backupName);
      const stagingName = stagingSheetName(attemptId, i);
      // Desktop Excel cannot copy a VeryHidden backup directly.
      try {
        backup.visibility = Excel.SheetVisibility.visible;
        await context.sync();
        const copy = backup.copy(Excel.WorksheetPositionType.end);
        copy.name = stagingName;
        copy.visibility = Excel.SheetVisibility.visible;
        staging.push({ ...s, stagingName });
        await context.sync();
      } finally {
        backup.visibility = Excel.SheetVisibility.veryHidden;
        await context.sync();
      }
    }

    sheets.load("items/name");
    await context.sync();
    const stagingNames = new Set(staging.map((s) => s.stagingName));
    const backupNames = new Set(cp.sheets.map((s) => s.backupName));
    // 先把现有用户表改成临时名，而不是直接删除。
    // 若后面把暂存表改回原名失败，用户数据还在 __AI_OLD_* 里。
    let parkIdx = 0;
    for (const ws of sheets.items) {
      if (stagingNames.has(ws.name) || backupNames.has(ws.name) || isInternalSheet(ws.name)) continue;
      ws.name = parkedSheetName(attemptId, parkIdx);
      parkIdx++;
    }

    await context.sync();

    try {
      for (const s of staging) {
        const ws = context.workbook.worksheets.getItem(s.stagingName);
        ws.name = s.originalName;
        ws.visibility = asVisibility(s.visibility);
      }
      await context.sync();
    } catch (err) {
      throw new Error(
        `恢复检查点时未能把备份改回原名，用户表已暂存为 ${OLD_PREFIX}*，数据未删除。请重试恢复或手动把这些表改回原名。` +
          (err instanceof Error && err.message ? `（${err.message}）` : "")
      );
    }

    sheets.load("items/name");
    await context.sync();
    for (const ws of sheets.items) {
      if (ws.name.startsWith(OLD_PREFIX)) ws.delete();
    }
    await context.sync();

    const visibleCount = staging.filter((s) => asVisibility(s.visibility) === Excel.SheetVisibility.visible).length;
    if (visibleCount === 0 && staging[0]) {
      context.workbook.worksheets.getItem(staging[0].originalName).visibility = Excel.SheetVisibility.visible;
      await context.sync();
    }

    for (const s of staging) {
      const ws = context.workbook.worksheets.getItem(s.originalName);
      ws.tables.load("items/name");
      await context.sync();
      const count = Math.min(ws.tables.items.length, s.tables.length);
      for (let i = 0; i < count; i++) {
        if (s.tables[i] && ws.tables.items[i].name !== s.tables[i]) {
          ws.tables.items[i].name = s.tables[i];
        }
      }
      await context.sync();
    }

    const ordered = [...staging].sort((a, b) => a.position - b.position);
    for (const s of ordered) {
      context.workbook.worksheets.getItem(s.originalName).position = s.position;
    }
    await context.sync();

    // Reapply captured formulas after all original sheet/table names exist.
    // Otherwise Excel's automatic rename tracking leaves references pointing
    // at the parked sheets that were just removed (#REF!). Constants, including
    // formula-like text, are never written through the formulas setter.
    for (const block of formulas) {
      context.workbook.worksheets.getItem(block.sheet).getRange(block.range).formulas = block.formulas;
    }
    await context.sync();

    try {
      context.workbook.worksheets.getItem(cp.activeSheet).activate();
      await context.sync();
    } catch {
      /* 原活动表可能当时就是隐藏表 */
    }
  });

  const count = cp.sheets.length;
  await discardCheckpoint(false);
  notify("restored");
  return `已恢复到任务开始时的工作簿（${count} 个 Sheet）`;
}

export async function discardCheckpoint(emitEvent = true): Promise<void> {
  const previous = current;
  current = null;
  sessionSnapshotted = true;
  persist(null);
  if (previous?.storage === "file") await deleteCheckpointFile(previous.id);
  try {
    const names = new Set(previous?.sheets.map((s) => s.backupName).filter(Boolean) ?? []);
    await deleteInternalSheets();
    if (names.size) await deleteSheetsNamed(names);
  } catch {
    /* 备份表可能已被用户手动删掉 */
  }
  if (emitEvent) notify("discarded");
}

/** 打开窗格时找回上次未丢弃的检查点，并清掉孤儿备份表 */
export async function hydrateCheckpoint(): Promise<TaskCheckpoint | null> {
  const saved = readPersisted();
  try {
    const names = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((s) => s.name);
    });
    const present = new Set(names);
    if (saved?.storage === "file" && saved.sheets.length && (await hasCheckpointFile(saved.id))) {
      current = saved;
      sessionSnapshotted = true;
      notify("hydrated");
      void pruneCheckpointFiles(saved.id);
      // 旧版本留在工作簿里的隐藏备份表一并清掉
      await deleteInternalSheets();
      return current;
    }
    if (saved && saved.storage !== "file" && saved.sheets.length && saved.sheets.every((s) => present.has(s.backupName))) {
      current = saved;
      sessionSnapshotted = true;
      notify("hydrated");
      const keep = new Set(saved.sheets.map((s) => s.backupName));
      await deleteInternalSheets(keep);
      return current;
    }
  } catch {
    /* Excel 尚未就绪 */
  }
  current = null;
  persist(null);
  try {
    await deleteInternalSheets();
  } catch {
    /* ignore */
  }
  notify("hydrated");
  return null;
}
