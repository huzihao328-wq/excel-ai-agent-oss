/* global Excel, Office */
import { csvBlob, ensureExtension, sanitizeDownloadName, triggerDownload, xlsxBlob } from "../exportFile";
import { ToolResult } from "../types";
import { isInternalSheet } from "./checkpoint";

const MAX_EXPORT_CELLS = 200000;

function todayStamp(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
}

function getSlice(file: Office.File, index: number): Promise<Office.Slice> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || "读取工作簿分片失败"));
    });
  });
}

/** 导出当前打开的工作簿原件（含格式、图表、透视），走浏览器下载。 */
export async function downloadCurrentWorkbook(filename?: string): Promise<string> {
  if (typeof Office === "undefined" || !Office.context?.document?.getFileAsync) {
    throw new Error("当前环境无法读取工作簿文件。请在 Excel 桌面版任务窗格里导出。");
  }
  const bytes = await new Promise<ArrayBuffer>((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error?.message || "无法读取当前工作簿"));
        return;
      }
      const file = result.value;
      void (async () => {
        try {
          const chunks: Uint8Array[] = [];
          let total = 0;
          for (let i = 0; i < file.sliceCount; i++) {
            const slice = await getSlice(file, i);
            const data = slice.data;
            const part =
              data instanceof ArrayBuffer
                ? new Uint8Array(data)
                : Array.isArray(data)
                  ? Uint8Array.from(data as number[])
                  : new Uint8Array(data);
            chunks.push(part);
            total += part.byteLength;
          }
          const out = new Uint8Array(total);
          let offset = 0;
          for (const part of chunks) {
            out.set(part, offset);
            offset += part.byteLength;
          }
          resolve(out.buffer);
        } catch (e) {
          reject(e);
        } finally {
          file.closeAsync();
        }
      })();
    });
  });
  const name = ensureExtension(sanitizeDownloadName(filename || `工作簿-${todayStamp()}`, "工作簿"), "xlsx");
          triggerDownload(
    name,
    new Blob([new Uint8Array(bytes)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
  );
  return name;
}

async function readSheetGrid(sheetName: string, rangeAddr?: string): Promise<{ name: string; rows: unknown[][]; cells: number }> {
  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(sheetName);
    ws.load("name");
    const range = rangeAddr ? ws.getRange(rangeAddr) : ws.getUsedRangeOrNullObject(true);
    range.load("address,rowCount,columnCount,isNullObject");
    await context.sync();
    if ((range as any).isNullObject) {
      return { name: ws.name, rows: [], cells: 0 };
    }
    const rows = range.rowCount;
    const cols = range.columnCount;
    const cells = rows * cols;
    if (cells > MAX_EXPORT_CELLS) {
      throw new Error(
        `${ws.name}!${range.address.split("!").pop()} 有 ${cells} 个单元格，超过 ${MAX_EXPORT_CELLS} 导出上限。请指定更小的 range。`
      );
    }
    const chunkRows = Math.max(20, Math.floor(20000 / Math.max(cols, 1)));
    const grid: unknown[][] = [];
    for (let r = 0; r < rows; r += chunkRows) {
      const n = Math.min(chunkRows, rows - r);
      const chunk = range.getCell(r, 0).getAbsoluteResizedRange(n, cols);
      chunk.load("formulas");
      await context.sync();
      grid.push(...(chunk.formulas as unknown[][]));
    }
    return { name: ws.name, rows: grid, cells };
  });
}

export async function toolExportFile(args: any): Promise<ToolResult> {
  const format = String(args.format || "xlsx").toLowerCase() === "csv" ? "csv" : "xlsx";
  const sheetArg = String(args.sheet || "").trim();
  const rangeArg = String(args.range || "").trim();
  const wantAll = !sheetArg || /^(all|\*|全部|整个工作簿)$/i.test(sheetArg);

  if (format === "xlsx" && wantAll && !rangeArg) {
    const filename = await downloadCurrentWorkbook(args.filename);
    return {
      ok: true,
      data: { filename, format: "xlsx", scope: "workbook" },
      summary: `已导出当前工作簿为 ${filename}（含格式与图表）。请在浏览器/系统下载里打开或另存。插件无法直接覆盖磁盘上的原路径。`,
    };
  }

  const targetSheets = await Excel.run(async (context) => {
    if (sheetArg && !wantAll) {
      const ws = context.workbook.worksheets.getItem(sheetArg);
      ws.load("name");
      await context.sync();
      if (isInternalSheet(ws.name)) throw new Error("不能导出检查点备份表。");
      return [ws.name];
    }
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();
    if (format === "csv") return [active.name];
    return sheets.items.map((s) => s.name).filter((n) => !isInternalSheet(n));
  });

  if (format === "csv" && targetSheets.length !== 1) {
    return { ok: false, error: "CSV 只能导出单个 Sheet，请指定 sheet。" };
  }

  const grids = [];
  for (const name of targetSheets) {
    grids.push(await readSheetGrid(name, rangeArg || undefined));
  }
  const nonempty = grids.filter((g) => g.rows.length > 0);
  if (nonempty.length === 0) return { ok: false, error: "没有可导出的数据" };

  const fallback = nonempty.length === 1 ? nonempty[0].name : `工作簿-${todayStamp()}`;
  const filename = ensureExtension(sanitizeDownloadName(args.filename || fallback, fallback), format);
  const blob = format === "csv" ? csvBlob(nonempty[0].rows) : await xlsxBlob(nonempty);
  triggerDownload(filename, blob);
  return {
    ok: true,
    data: {
      filename,
      format,
      sheets: nonempty.map((g) => ({ name: g.name, rows: g.rows.length, cells: g.cells })),
    },
    summary: `已导出 ${filename}。请在下载中打开或另存；无法直接改写磁盘上的原文件。`,
  };
}
