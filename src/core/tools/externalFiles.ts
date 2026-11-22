/* global Excel */
/*
 * 外部文件支持：用户把本地 Excel/CSV 附加到任务窗格后在内存中解析，
 * 让 AI 可以读取其他文件的数据，并导入当前工作簿处理（绕过 Office.js
 * 只能访问当前工作簿的限制）。
 */
import { ToolResult } from "../types";
import { gridPreview, pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { protectFormulaLikeTextGrid } from "./safeValues";
import {
  deleteStoredAttachment,
  getStoredAttachment,
  listStoredAttachments,
  putStoredAttachment,
} from "../attachmentStore";
import { ensureExtension, sanitizeDownloadName, triggerDownload } from "../exportFile";

let sheetJsPromise: Promise<typeof import("xlsx")> | null = null;

function loadSheetJs(): Promise<typeof import("xlsx")> {
  sheetJsPromise ||= import("xlsx").catch((error) => {
    sheetJsPromise = null;
    throw error;
  });
  return sheetJsPromise;
}

export interface ExternalSheet {
  name: string;
  rows: (string | number | boolean)[][];
}

export interface ExternalFile {
  name: string;
  sheets: ExternalSheet[];
  loadedAt: number;
  /** 当前附加列表中的序号（从 1 起，按添加顺序，随增删重排） */
  slot: number;
  /** 如「文件一」 */
  alias: string;
  /** 如「文件一：金额.xlsx」 */
  caption: string;
}

interface ExternalFileData {
  name: string;
  sheets: ExternalSheet[];
  loadedAt: number;
}

const MAX_FILE_BYTES = 20 * 1024 * 1024;
const MAX_IMPORT_CELLS = 200000;
const SEL_KEY = "excel-ai-agent-ext-selection";

const files = new Map<string, ExternalFileData>();
let listeners: (() => void)[] = [];
let hydratePromise: Promise<number> | null = null;
let persistChain: Promise<void> = Promise.resolve();
const attachmentVersions = new Map<string, number>();
let selectionVersion = 0;

function nextAttachmentVersion(name: string): number {
  const version = (attachmentVersions.get(name) || 0) + 1;
  attachmentVersions.set(name, version);
  return version;
}

function persistAttachment(action: () => Promise<void>): Promise<void> {
  persistChain = persistChain.then(action).catch(() => undefined);
  return persistChain;
}

/** 用户在附加文件预览器中框选的区域 */
export interface ExternalSelection {
  file: string;
  sheet: string;
  range: string; // A1 表示法
}

let extSelection: ExternalSelection | null = null;
let selListeners: (() => void)[] = [];

export function getExternalSelection(): ExternalSelection | null {
  return extSelection;
}

export function setExternalSelection(sel: ExternalSelection | null): void {
  selectionVersion++;
  extSelection = sel;
  try {
    if (sel) localStorage.setItem(SEL_KEY, JSON.stringify(sel));
    else localStorage.removeItem(SEL_KEY);
  } catch {
    /* ignore */
  }
  selListeners.forEach((f) => f());
}

export function onExternalSelectionChange(fn: () => void): () => void {
  selListeners.push(fn);
  return () => {
    selListeners = selListeners.filter((f) => f !== fn);
  };
}

// ---------------- A1 区域解析 ----------------

export function colIndexToLetter(index: number): string {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function colLetterToIndex(letter: string): number {
  let n = 0;
  for (const ch of letter.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export interface A1Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** 解析 "A2:D50" 或 "B3" 为 0 基的行列区间（含端点） */
export function parseA1Range(range: string): A1Rect {
  const parse = (cell: string) => {
    const m = /^([A-Za-z]+)(\d+)$/.exec(cell.trim());
    if (!m) throw new Error(`无法解析单元格地址「${cell}」，请用 A1 表示法（如 B2 或 A2:D50）`);
    return { c: colLetterToIndex(m[1]), r: parseInt(m[2], 10) - 1 };
  };
  const parts = range.split(":");
  const a = parse(parts[0]);
  const b = parts.length > 1 ? parse(parts[1]) : a;
  return {
    r1: Math.min(a.r, b.r),
    c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r),
    c2: Math.max(a.c, b.c),
  };
}

/** 解析单个或逗号分隔的多区域，如 "A2:A50,C2:C50,E2:E50" */
export function parseA1Ranges(range: string): A1Rect[] {
  return range
    .split(/[,;，；]/)
    .map((part) => part.trim())
    .filter(Boolean)
    .map(parseA1Range);
}

function clipRect(rect: A1Rect, rowCount: number, colCount: number): A1Rect {
  return {
    r1: Math.max(0, rect.r1),
    c1: Math.max(0, rect.c1),
    r2: Math.min(rowCount - 1, rect.r2),
    c2: Math.min(colCount - 1, rect.c2),
  };
}

function rectContains(rect: A1Rect, r: number, c: number): boolean {
  return r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2;
}

/** 取外部文件某区域的值（支持 Ctrl 多选的逗号分隔区域；超出实际数据时自动裁剪） */
export function getExternalRangeValues(fileName: string, sheetName: string, range: string): any[][] {
  const { sheet } = findSheet(fileName, sheetName);
  // 按实际数据的行列数裁剪，避免 A1:XFD1048576 之类的区域生成海量空单元格
  const maxCols = sheet.rows.reduce((m, r) => Math.max(m, r.length), 0);
  const rects = parseA1Ranges(range)
    .map((rect) => clipRect(rect, sheet.rows.length, maxCols))
    .filter((rect) => rect.r2 >= rect.r1 && rect.c2 >= rect.c1 && sheet.rows.length > 0);
  if (rects.length === 0) return [];

  if (rects.length === 1) {
    const { r1, c1, r2, c2 } = rects[0];
    return sheet.rows.slice(r1, r2 + 1).map((row) => {
      const out: any[] = [];
      for (let c = c1; c <= c2; c++) out.push(row[c] ?? "");
      return out;
    });
  }

  const colSet = new Set<number>();
  const rowSet = new Set<number>();
  for (const rect of rects) {
    for (let c = rect.c1; c <= rect.c2; c++) colSet.add(c);
    for (let r = rect.r1; r <= rect.r2; r++) rowSet.add(r);
  }
  const cols = [...colSet].sort((a, b) => a - b);
  const rows = [...rowSet].sort((a, b) => a - b);
  return rows.map((r) => {
    const row = sheet.rows[r] || [];
    return cols.map((c) => (rects.some((rect) => rectContains(rect, r, c)) ? row[c] ?? "" : ""));
  });
}

const CN_ONES = "零一二三四五六七八九";

/** 1 → 一，12 → 十二，23 → 二十三 */
export function toChineseNum(n: number): string {
  if (n <= 0) return String(n);
  if (n < 10) return CN_ONES[n];
  if (n === 10) return "十";
  if (n < 20) return `十${CN_ONES[n - 10]}`;
  if (n < 100) {
    const tens = Math.floor(n / 10);
    const ones = n % 10;
    return `${CN_ONES[tens]}十${ones ? CN_ONES[ones] : ""}`;
  }
  return String(n);
}

export function formatFileAlias(slot: number): string {
  return `文件${toChineseNum(slot)}`;
}

function decorateFile(data: ExternalFileData, slot: number): ExternalFile {
  const alias = formatFileAlias(slot);
  return { ...data, slot, alias, caption: `${alias}：${data.name}` };
}

function parseIndexToken(token: string): number | null {
  const t = token.trim();
  if (/^\d+$/.test(t)) {
    const n = parseInt(t, 10);
    return n > 0 ? n : null;
  }
  const onesMap: Record<string, number> = {
    零: 0,
    〇: 0,
    一: 1,
    二: 2,
    两: 2,
    三: 3,
    四: 4,
    五: 5,
    六: 6,
    七: 7,
    八: 8,
    九: 9,
  };
  if (t === "十") return 10;
  if (t.startsWith("十")) {
    const ones = onesMap[t.slice(1)];
    return ones != null && ones > 0 ? 10 + ones : null;
  }
  const shi = t.indexOf("十");
  if (shi === -1) {
    const n = onesMap[t];
    return n != null && n > 0 ? n : null;
  }
  const tens = onesMap[t.slice(0, shi)];
  const onesPart = t.slice(shi + 1);
  const ones = onesPart ? onesMap[onesPart] : 0;
  if (tens == null || tens < 1 || ones == null) return null;
  return tens * 10 + ones;
}

/** 解析「文件一」「文件1」「file 2」「第一个文件」等，返回 1-based 序号 */
export function parseFileSlotRef(ref: string): number | null {
  const s = ref.trim();
  if (!s) return null;
  const ordinal = /^第\s*([0-9]+|[一二三四五六七八九十两]+)\s*个\s*(?:文件|附件|file)?$/i.exec(s);
  if (ordinal) return parseIndexToken(ordinal[1]);
  const named =
    /^(?:文件|附件|file)\s*([0-9]+|[一二三四五六七八九十两]+)(?:\s*[个号])?(?:\s*[：:，,].*)?$/i.exec(s);
  if (named) return parseIndexToken(named[1]);
  return null;
}

function availableFileHint(): string {
  const list = getExternalFiles();
  return list.length ? list.map((f) => f.caption).join("、") : "(无)";
}

export function resolveExternalFile(ref: string): ExternalFile {
  const list = getExternalFiles();
  const raw = ref.trim();
  if (!raw) {
    throw new Error(`未指定外部文件。可用: ${availableFileHint()}`);
  }
  const exact = list.find((f) => f.name === raw);
  if (exact) return exact;

  const slot = parseFileSlotRef(raw);
  if (slot != null) {
    const bySlot = list.find((f) => f.slot === slot);
    if (bySlot) return bySlot;
    throw new Error(`没有${formatFileAlias(slot)}。可用: ${availableFileHint()}`);
  }

  const fuzzy = list.filter((f) => f.name.includes(raw) || f.caption.includes(raw) || f.alias === raw);
  if (fuzzy.length > 1) {
    throw new Error(`「${raw}」匹配到多个附件: ${fuzzy.map((f) => f.caption).join("、")}。请改用文件编号（如文件一）或完整文件名。`);
  }
  if (fuzzy[0]) return fuzzy[0];
  throw new Error(`未找到外部文件「${raw}」。可用: ${availableFileHint()}`);
}

export function getExternalFiles(): ExternalFile[] {
  return [...files.values()].map((f, i) => decorateFile(f, i + 1));
}

export function removeExternalFile(name: string): void {
  nextAttachmentVersion(name);
  files.delete(name);
  if (extSelection?.file === name) setExternalSelection(null);
  listeners.forEach((f) => f());
  void persistAttachment(async () => {
    await deleteStoredAttachment(name);
    await persistOrder();
  });
}

export function onExternalFilesChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((f) => f !== fn);
  };
}

function pad(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function formatDate(d: Date): string {
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  return d.getHours() || d.getMinutes() || d.getSeconds()
    ? `${date} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`
    : date;
}

/** CSV 解码：优先 UTF-8，失败则按 GBK（国内导出的 CSV 常见编码） */
function decodeCsv(buf: ArrayBuffer): string {
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(buf);
  } catch {
    try {
      return new TextDecoder("gbk").decode(buf);
    } catch {
      return new TextDecoder("utf-8").decode(buf);
    }
  }
}

export async function loadExternalFile(file: File): Promise<string> {
  if (!/\.(xlsx|xlsm|xls|csv)$/i.test(file.name)) throw new Error("请附加 xlsx / xls / xlsm / csv 文件");
  if (file.size > MAX_FILE_BYTES) {
    throw new Error(`${file.name} 超过 20MB，请拆分后再附加`);
  }
  const version = nextAttachmentVersion(file.name);
  const buf = await file.arrayBuffer();
  return attachExternalBuffer(file.name, buf, { persist: true, version });
}

export function attachmentKind(name: string): "csv" | "xlsx" {
  return /\.csv$/i.test(name) ? "csv" : "xlsx";
}

export async function attachExternalBuffer(
  name: string,
  buf: ArrayBuffer,
  options?: { persist?: boolean; loadedAt?: number; version?: number }
): Promise<string> {
  const version = options?.version ?? nextAttachmentVersion(name);
  if (buf.byteLength > MAX_FILE_BYTES) {
    throw new Error(`${name} 超过 20MB，请拆分后再附加`);
  }
  const XLSX = await loadSheetJs();
  if (version !== attachmentVersions.get(name)) return `已跳过「${name}」的旧加载请求`;
  let wb: import("xlsx").WorkBook;
  if (attachmentKind(name) === "csv") {
    wb = XLSX.read(decodeCsv(buf), { type: "string", raw: true });
  } else {
    wb = XLSX.read(buf, { type: "array", cellDates: true });
  }
  const sheets: ExternalSheet[] = wb.SheetNames.map((sheetName) => {
    const ws = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as any[][];
    const rows = raw.map((r) => r.map((v) => (v instanceof Date ? formatDate(v) : v)));
    return { name: sheetName, rows };
  });
  const loadedAt = options?.loadedAt ?? Date.now();
  if (extSelection?.file === name) setExternalSelection(null);
  files.set(name, { name, sheets, loadedAt });
  listeners.forEach((f) => f());
  if (options?.persist !== false) {
    const order = getExternalFiles().findIndex((f) => f.name === name);
    try {
      await persistAttachment(() => putStoredAttachment({
        name,
        loadedAt,
        order: order < 0 ? files.size - 1 : order,
        kind: attachmentKind(name),
        bytes: buf.slice(0),
      }));
    } catch {
      /* 无 IndexedDB 时仍可在本会话使用 */
    }
  }
  const totalRows = sheets.reduce((s, x) => s + x.rows.length, 0);
  const attached = getExternalFiles().find((x) => x.name === name);
  const caption = attached?.caption ?? name;
  return `已附加 ${caption}（${sheets.length} 个工作表，共 ${totalRows} 行）。对 AI 说「${attached?.alias ?? "文件一"}」即可，不必念原文件名`;
}

async function persistOrder(): Promise<void> {
  const list = getExternalFiles();
  for (const f of list) {
    const stored = await getStoredAttachment(f.name);
    if (!stored) continue;
    if (stored.order === f.slot - 1) continue;
    await putStoredAttachment({ ...stored, order: f.slot - 1 });
  }
}

function restoreSelection(): void {
  try {
    const raw = localStorage.getItem(SEL_KEY);
    if (!raw) return;
    const sel = JSON.parse(raw) as ExternalSelection;
    if (sel?.file && sel.sheet && sel.range && files.has(sel.file)) {
      extSelection = sel;
      selListeners.forEach((f) => f());
    } else {
      localStorage.removeItem(SEL_KEY);
    }
  } catch {
    /* ignore */
  }
}

/** 从 IndexedDB 恢复上次附加的文件。重复调用会复用同一次 Promise。 */
export function hydrateExternalFiles(): Promise<number> {
  hydratePromise ||= (async () => {
    const selectionAtStart = selectionVersion;
    const stored = await listStoredAttachments();
    let restored = 0;
    for (const item of stored) {
      if (files.has(item.name) || attachmentVersions.has(item.name)) continue;
      try {
        await attachExternalBuffer(item.name, item.bytes, { persist: false, loadedAt: item.loadedAt });
        restored++;
      } catch {
        // A parse/chunk-loading failure must not delete the user's saved file.
      }
    }
    if (selectionAtStart === selectionVersion) restoreSelection();
    return restored;
  })();
  return hydratePromise;
}

export async function downloadAttachedFile(name: string): Promise<string> {
  const file = resolveExternalFile(name);
  const stored = await getStoredAttachment(file.name);
  if (!stored) throw new Error(`没有「${file.caption}」的原始文件副本，请重新附加后再下载`);
  const ext = stored.kind === "csv" ? "csv" : "xlsx";
  const filename = ensureExtension(sanitizeDownloadName(file.name, "附件"), ext);
  const type =
    stored.kind === "csv"
      ? "text/csv;charset=utf-8"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  triggerDownload(filename, new Blob([new Uint8Array(stored.bytes)], { type }));
  return `已开始下载 ${filename}`;
}

// ---------------- 供 Agent 调用的工具 ----------------

export function toolListExternalFiles(): ToolResult {
  const list = getExternalFiles();
  if (list.length === 0) {
    return {
      ok: true,
      data: {
        files: [],
        hint: "当前没有附加的外部文件。请让用户点击输入框左侧的 📎 按钮（或把文件拖进对话窗口）附加 Excel/CSV。附加过的文件会在关闭任务窗格后自动恢复。",
      },
      summary: "没有附加的外部文件",
    };
  }
  const data = list.map((f) => ({
    alias: f.alias,
    caption: f.caption,
    file: f.name,
    sheets: f.sheets.map((s) => ({
      name: s.name,
      row_count: s.rows.length,
      column_count: s.rows[0]?.length ?? 0,
      headers_guess: s.rows[0]?.slice(0, 30) ?? [],
      sample_rows: s.rows.slice(1, 4).map((r) => r.slice(0, 30)),
    })),
  }));
  return { ok: true, data: { files: data }, summary: `共 ${list.length} 个外部文件` };
}

function findSheet(fileName: string, sheetName?: string): { file: ExternalFile; sheet: ExternalSheet } {
  const file = resolveExternalFile(fileName);
  const sheet = sheetName ? file.sheets.find((s) => s.name === sheetName) : file.sheets[0];
  if (!sheet) {
    throw new Error(`「${file.caption}」中没有工作表「${sheetName}」。可用: ${file.sheets.map((s) => s.name).join(", ")}`);
  }
  return { file, sheet };
}

export function toolReadExternalSheet(args: any): ToolResult {
  const { file, sheet } = findSheet(args.file, args.sheet);

  // 指定了 A1 区域时优先按区域读取（对应用户在预览器中的框选）
  if (args.range) {
    const values = getExternalRangeValues(args.file, args.sheet || sheet.name, args.range);
    const cells = values.length * (values[0]?.length ?? 0);
    if (cells > 5000) {
      return { ok: false, error: `区域 ${args.range} 有约 ${cells} 个单元格，超过 5000 上限。请缩小范围或改用 import_external_sheet 导入。` };
    }
    return {
      ok: true,
      data: { range: args.range, total_rows: sheet.rows.length, values },
      summary: `读取 ${file.caption} 的 ${args.range}`,
    };
  }

  const startRow = Math.max(1, args.start_row || 1);
  const maxRows = Math.min(args.max_rows || 100, 1000);
  const slice = sheet.rows.slice(startRow - 1, startRow - 1 + maxRows);
  const cols = Math.max(...slice.map((r) => r.length), 0);
  if (slice.length * cols > 5000) {
    return {
      ok: false,
      error: `本次将返回约 ${slice.length * cols} 个单元格，超过 5000 上限。请减小 max_rows，或直接用 import_external_sheet 导入后用公式处理。`,
    };
  }
  return {
    ok: true,
    data: {
      total_rows: sheet.rows.length,
      returned_rows: slice.length,
      start_row: startRow,
      values: slice,
    },
    summary: `读取 ${file.caption} 第 ${startRow}~${startRow + slice.length - 1} 行`,
  };
}

function sanitizeSheetName(name: string): string {
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_");
  return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned;
}

export async function toolImportExternalSheet(args: any): Promise<ToolResult> {
  const { file, sheet } = findSheet(args.file, args.sheet);
  const rows: (string | number | boolean)[][] = args.range
    ? getExternalRangeValues(args.file, args.sheet || sheet.name, args.range)
    : sheet.rows;
  if (rows.length === 0) {
    return { ok: false, error: args.range ? `区域 ${args.range} 没有数据` : `工作表「${sheet.name}」是空的` };
  }
  const cols = rows.reduce((max, row) => Math.max(max, row.length), 0);
  if (!cols) return { ok: false, error: `工作表「${sheet.name}」没有可导入的列` };
  if (rows.length * cols > MAX_IMPORT_CELLS) {
    return {
      ok: false,
      error: `数据量 ${rows.length}×${cols} 超过 ${MAX_IMPORT_CELLS} 单元格上限，请先在源文件中拆分。`,
    };
  }
  const normalized = rows.map((r) => {
    const row: any[] = [...r];
    while (row.length < cols) row.push("");
    return row;
  });
  const safeValues = protectFormulaLikeTextGrid(normalized);

  const baseName = sanitizeSheetName(
    args.target_sheet_name || `${file.name.replace(/\.[^.]+$/, "")}-${sheet.name}`
  );

  const finalName = await Excel.run(async (context) => {
    const wsCol = context.workbook.worksheets;
    wsCol.load("items/name");
    await context.sync();
    const existing = new Set(wsCol.items.map((w) => w.name));
    let name = baseName;
    let i = 2;
    while (existing.has(name)) {
      name = sanitizeSheetName(`${baseName.slice(0, 27)}(${i})`);
      i++;
    }
    const ws = wsCol.add(name);
    try {
      // 分块写入，避免单次载荷过大
      const chunkRows = Math.max(50, Math.floor(20000 / cols));
      for (let r = 0; r < normalized.length; r += chunkRows) {
        const chunk = safeValues.slice(r, r + chunkRows);
        ws.getRangeByIndexes(r, 0, chunk.length, cols).values = chunk;
        await context.sync();
      }
      ws.activate();
      await context.sync();
      return name;
    } catch (e) {
      try {
        ws.delete();
        await context.sync();
      } catch {
        /* 半成品 Sheet 可能已不可用 */
      }
      throw e;
    }
  });

  pushOpLog({
    description: `从 ${file.caption} / ${sheet.name} 导入为 Sheet「${finalName}」（${rows.length}×${cols}）`,
    kind: "import",
    sheet: finalName,
    after: gridPreview(normalized),
  });
  registerArtifact({
    name: finalName,
    kind: "sheet",
    sheet: finalName,
    address: `A1:${colIndexToLetter(Math.max(0, cols - 1))}${rows.length}`,
    brief: `从 ${file.caption} 导入（${rows.length}×${cols}）`,
  });

  return {
    ok: true,
    data: { imported_sheet: finalName, rows: rows.length, columns: cols },
    summary: `已把 ${file.caption} / ${sheet.name}${args.range ? ` 的 ${args.range}` : ""} 导入为 Sheet「${finalName}」（${rows.length}×${cols}）`,
    undoable: false,
  };
}
