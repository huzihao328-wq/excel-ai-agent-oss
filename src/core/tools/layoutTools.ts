/* global Excel */
/** 打印设置与目录页：只改页面布局或新建目录 Sheet，不动单元格数据。 */
import { ToolResult } from "../types";
import { pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { isInternalSheet } from "./checkpoint";
import { uniqueSheetName } from "./tableTools";
import { quoteSheetName } from "./formulaRefs";
import { isEmptyCell } from "./safeValues";

// ---------------- 打印设置 ----------------

const PAPER_SIZES: Record<string, string> = { a4: "A4", a3: "A3", a5: "A5", b4: "B4", b5: "B5", letter: "Letter", legal: "Legal" };

/** 页边距预设（磅），与 Excel「页面布局 → 页边距」的常规 / 窄 / 宽一致 */
export const MARGIN_PRESETS: Record<string, { top: number; bottom: number; left: number; right: number; header: number; footer: number }> = {
  normal: { top: 54, bottom: 54, left: 50.4, right: 50.4, header: 21.6, footer: 21.6 },
  narrow: { top: 54, bottom: 54, left: 18, right: 18, header: 21.6, footer: 21.6 },
  wide: { top: 72, bottom: 72, left: 72, right: 72, header: 36, footer: 36 },
};

export const PAGE_NUMBER_FOOTER = "第 &P 页，共 &N 页";

/** '1' / '1:2' / '$1:$2' → '$1:$2' */
export function normalizeTitleRows(spec: unknown): string {
  const m = /^\$?(\d+)(?::\$?(\d+))?$/.exec(String(spec ?? "").replace(/\s+/g, ""));
  if (!m) throw new Error(`print_title_rows「${spec}」应写成 '1' 或 '1:2'`);
  const a = Number(m[1]);
  const b = Number(m[2] || m[1]);
  if (a < 1 || b < a) throw new Error(`print_title_rows「${spec}」行号无效`);
  return `$${a}:$${b}`;
}

/** 'A' / 'A:B' / '$A:$B' → '$A:$B' */
export function normalizeTitleColumns(spec: unknown): string {
  const m = /^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/.exec(String(spec ?? "").replace(/\s+/g, ""));
  if (!m) throw new Error(`print_title_columns「${spec}」应写成 'A' 或 'A:B'`);
  return `$${m[1].toUpperCase()}:$${(m[2] || m[1]).toUpperCase()}`;
}

/** 页眉页脚里 & 是控制符，普通文字里的 & 要写成 && */
export function headerFooterText(text: unknown): string {
  return String(text ?? "").replace(/&/g, "&&");
}

function describeZoom(zoom: Excel.PageLayoutZoomOptions | undefined): string {
  if (!zoom) return "";
  if (zoom.horizontalFitToPages || zoom.verticalFitToPages) {
    const wide = zoom.horizontalFitToPages || "自动";
    const tall = zoom.verticalFitToPages || "自动";
    return `缩放到 ${wide} 页宽 × ${tall} 页高`;
  }
  return zoom.scale ? `缩放 ${zoom.scale}%` : "";
}

export async function toolSetPageLayout(args: any): Promise<ToolResult> {
  const requested = String(args.sheet || "").trim();
  if (requested && isInternalSheet(requested)) return { ok: false, error: "不能设置检查点备份表。" };
  const fitTo = args.fit_to == null ? "" : String(args.fit_to).toLowerCase();
  if (fitTo && args.scale != null) return { ok: false, error: "fit_to 和 scale 只能二选一" };
  if (fitTo && !["width", "page", "none"].includes(fitTo)) return { ok: false, error: `fit_to 只能是 width / page / none` };
  let titleRows = "";
  let titleColumns = "";
  try {
    if (args.print_title_rows != null && args.print_title_rows !== "") titleRows = normalizeTitleRows(args.print_title_rows);
    if (args.print_title_columns != null && args.print_title_columns !== "") titleColumns = normalizeTitleColumns(args.print_title_columns);
  } catch (e: any) {
    return { ok: false, error: e?.message || String(e) };
  }
  const margins = args.margins == null ? null : MARGIN_PRESETS[String(args.margins).toLowerCase()];
  if (args.margins != null && !margins) return { ok: false, error: "margins 只能是 normal / narrow / wide" };
  const paper = args.paper_size == null ? null : PAPER_SIZES[String(args.paper_size).toLowerCase()];
  if (args.paper_size != null && !paper) return { ok: false, error: `paper_size 只能是 ${Object.values(PAPER_SIZES).join(" / ")}` };
  let scale: number | null = null;
  if (args.scale != null) {
    scale = Math.round(Number(args.scale));
    if (!Number.isFinite(scale) || scale < 10 || scale > 400) return { ok: false, error: "scale 应在 10～400 之间" };
  }

  return Excel.run(async (context) => {
    const ws = requested ? context.workbook.worksheets.getItem(requested) : context.workbook.worksheets.getActiveWorksheet();
    ws.load("name");
    const layout = ws.pageLayout;
    const applied: string[] = [];
    if (args.orientation != null) {
      const landscape = /^land|横/i.test(String(args.orientation));
      layout.orientation = landscape ? "Landscape" : "Portrait";
      applied.push(landscape ? "横向" : "纵向");
    }
    if (paper) {
      layout.paperSize = paper as Excel.PaperType;
      applied.push(`纸张 ${paper}`);
    }
    if (titleRows) {
      layout.setPrintTitleRows(titleRows);
      applied.push(`每页重复第 ${titleRows.replace(/\$/g, "")} 行`);
    }
    if (titleColumns) {
      layout.setPrintTitleColumns(titleColumns);
      applied.push(`每页重复 ${titleColumns.replace(/\$/g, "")} 列`);
    }
    if (args.print_area != null && String(args.print_area).trim()) {
      const area = String(args.print_area).trim().replace(/^.*!/, "");
      layout.setPrintArea(area);
      applied.push(`打印区域 ${area}`);
    }
    if (margins) {
      layout.topMargin = margins.top;
      layout.bottomMargin = margins.bottom;
      layout.leftMargin = margins.left;
      layout.rightMargin = margins.right;
      layout.headerMargin = margins.header;
      layout.footerMargin = margins.footer;
      applied.push(`页边距 ${({ normal: "常规", narrow: "窄", wide: "宽" } as Record<string, string>)[String(args.margins).toLowerCase()]}`);
    }
    if (typeof args.center_horizontally === "boolean") {
      layout.centerHorizontally = args.center_horizontally;
      applied.push(args.center_horizontally ? "水平居中" : "取消水平居中");
    }
    if (typeof args.gridlines === "boolean") {
      layout.printGridlines = args.gridlines;
      applied.push(args.gridlines ? "打印网格线" : "不打印网格线");
    }
    const pages = layout.headersFooters.defaultForAllPages;
    if (args.header_text != null) {
      pages.centerHeader = headerFooterText(args.header_text);
      applied.push(args.header_text ? `页眉「${args.header_text}」` : "清除页眉");
    }
    if (typeof args.page_numbers === "boolean") {
      pages.centerFooter = args.page_numbers ? PAGE_NUMBER_FOOTER : "";
      applied.push(args.page_numbers ? "页码" : "清除页码");
    }
    if (args.footer_text != null) {
      const text = headerFooterText(args.footer_text);
      if (args.page_numbers === true) pages.leftFooter = text;
      else pages.centerFooter = text;
      applied.push(args.footer_text ? `页脚「${args.footer_text}」` : "清除页脚");
    }
    await context.sync();

    if (fitTo || scale != null) {
      // 「一页宽」的高度页数写 0 表示自动；个别版本不接受 0 时退回只限定宽度
      const zoom: Excel.PageLayoutZoomOptions =
        scale != null ? { scale } : fitTo === "none" ? { scale: 100 } : fitTo === "page" ? { horizontalFitToPages: 1, verticalFitToPages: 1 } : { horizontalFitToPages: 1, verticalFitToPages: 0 };
      try {
        layout.zoom = zoom;
        await context.sync();
      } catch (e) {
        if (fitTo !== "width") throw e;
        layout.zoom = { horizontalFitToPages: 1 };
        await context.sync();
      }
      applied.push(scale != null ? `缩放 ${scale}%` : fitTo === "none" ? "取消缩放" : fitTo === "page" ? "整表缩到一页" : "所有列缩到一页宽");
    }
    if (!applied.length) {
      return {
        ok: false,
        error: "没有给出要改的打印设置。可用：orientation、paper_size、fit_to、scale、print_title_rows、print_title_columns、print_area、margins、center_horizontally、gridlines、page_numbers、header_text、footer_text",
      };
    }

    layout.load("orientation,paperSize,zoom,centerHorizontally,printGridlines");
    await context.sync();
    pushOpLog({ description: `打印设置「${ws.name}」：${applied.join("、")}`, kind: "format", sheet: ws.name });
    return {
      ok: true,
      data: {
        sheet: ws.name,
        applied,
        current: {
          orientation: layout.orientation === "Landscape" ? "横向" : "纵向",
          paper_size: layout.paperSize,
          zoom: describeZoom(layout.zoom),
          center_horizontally: layout.centerHorizontally,
          gridlines: layout.printGridlines,
        },
        note: "只改打印/页面布局，不动数据；可在 Excel「文件 → 打印」预览效果。",
      },
      summary: `已设置「${ws.name}」打印格式：${applied.join("、")}`,
    };
  });
}

// ---------------- 目录页 ----------------

export const INDEX_HEADERS = ["序号", "工作表", "数据区域", "行数"];
const INDEX_RANGE_NAME = "_ExcelAIAgent_IndexArea";
const INDEX_COMMENT_PREFIX = "Excel AI Agent: create_sheet_index/";

export interface IndexEntry {
  name: string;
  address: string;
  rows: number;
  hidden: boolean;
}

export function buildIndexGrid(entries: IndexEntry[]): unknown[][] {
  const withStatus = entries.some((e) => e.hidden);
  const header = withStatus ? [...INDEX_HEADERS, "状态"] : INDEX_HEADERS;
  return [
    header,
    ...entries.map((e, i) => {
      const row: unknown[] = [i + 1, e.name, e.address || "（空表）", e.rows];
      if (withStatus) row.push(e.hidden ? "隐藏" : "");
      return row;
    }),
  ];
}

export async function toolCreateSheetIndex(args: any): Promise<ToolResult> {
  const baseName = String(args.name || "").trim() || "目录";
  if (isInternalSheet(baseName)) return { ok: false, error: "不能使用保留的检查点备份表名。" };
  const includeHidden = args.include_hidden === true;
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility");
    await context.sync();

    // 工作表局部名称随文件保存。即使同名表已存在、目录用了后缀，也能找到上次写入的区域。
    const comment = INDEX_COMMENT_PREFIX + baseName.toLowerCase();
    const candidates = sheets.items.filter((sheet) => !isInternalSheet(sheet.name)).map((sheet) => {
      const marker = sheet.names.getItemOrNullObject(INDEX_RANGE_NAME);
      marker.load("comment,type");
      return { sheet, marker };
    });
    await context.sync();
    const owned = candidates.filter(({ marker }) =>
      !marker.isNullObject && marker.type === "Range" && marker.comment.startsWith(INDEX_COMMENT_PREFIX)
    );
    const selected = owned.find(({ sheet }) => sheet.name.toLowerCase() === baseName.toLowerCase())
      ?? owned.find(({ marker }) => marker.comment === comment);
    const target = selected?.sheet;
    const previous = selected?.marker.getRange();
    if (previous) {
      previous.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/name");
    }
    const listed = sheets.items.filter(
      (s) => !isInternalSheet(s.name) && s !== target && (includeHidden || s.visibility === "Visible")
    );
    if (!listed.length) return { ok: false, error: "没有可列入目录的工作表" };
    const used = listed.map((s) => {
      const range = s.getUsedRangeOrNullObject(true);
      range.load("address,rowCount,isNullObject");
      return range;
    });
    await context.sync();
    const entries: IndexEntry[] = listed.map((s, i) => ({
      name: s.name,
      address: used[i].isNullObject ? "" : (used[i].address.split("!").pop() as string),
      rows: used[i].isNullObject ? 0 : used[i].rowCount,
      hidden: s.visibility !== "Visible",
    }));
    const grid = buildIndexGrid(entries);
    const cols = grid[0].length;

    const refreshed = !!target;
    let ws: Excel.Worksheet;
    if (target && previous) {
      ws = target;
      if (previous.worksheet.name !== ws.name || previous.rowIndex !== 0 || previous.columnIndex !== 0 ||
          (previous.columnCount !== 4 && previous.columnCount !== 5)) {
        return { ok: false, error: "原目录区域已移动或改变，请指定新目录名。" };
      }
      if (grid.length > previous.rowCount || cols > previous.columnCount) {
        const next = ws.getRangeByIndexes(0, 0, grid.length, cols);
        next.load("formulas");
        await context.sync();
        if (next.formulas.some((row, r) => row.some((cell, c) =>
          (r >= previous.rowCount || c >= previous.columnCount) && !isEmptyCell(cell)
        ))) {
          return { ok: false, error: "目录扩展区域已有内容。请先移走相邻数据，或指定新目录名。" };
        }
      }
      previous.clear();
    } else {
      ws = sheets.add(await uniqueSheetName(context, baseName));
    }
    ws.load("name");
    ws.position = 0;
    const output = ws.getRangeByIndexes(0, 0, grid.length, cols);
    output.values = grid as any[][];
    entries.forEach((e, i) => {
      ws.getRangeByIndexes(i + 1, 1, 1, 1).hyperlink = {
        documentReference: `${quoteSheetName(e.name)}!A1`,
        textToDisplay: e.name,
        screenTip: `转到「${e.name}」`,
      };
    });
    const header = ws.getRangeByIndexes(0, 0, 1, cols);
    header.format.font.bold = true;
    header.format.fill.color = "#107C41";
    header.format.font.color = "#FFFFFF";
    ws.freezePanes.freezeRows(1);
    output.format.autofitColumns();
    const address = `A1:${String.fromCharCode(64 + cols)}${grid.length}`;
    if (selected) {
      selected.marker.formula = `=${quoteSheetName(ws.name)}!$A$1:$${String.fromCharCode(64 + cols)}$${grid.length}`;
    } else {
      ws.names.add(INDEX_RANGE_NAME, output, comment).visible = false;
    }
    ws.activate();
    await context.sync();

    pushOpLog({
      description: `${refreshed ? "刷新" : "生成"}目录「${ws.name}」（${entries.length} 张表，可点击跳转）`,
      kind: "structure",
      sheet: ws.name,
      address,
    });
    registerArtifact({ name: ws.name, kind: "sheet", sheet: ws.name, address, brief: `工作表目录（${entries.length} 张表，带跳转链接）` });
    return {
      ok: true,
      data: {
        sheet: ws.name,
        refreshed,
        sheets: entries.map((e) => ({ name: e.name, range: e.address, rows: e.rows, ...(e.hidden ? { hidden: true } : {}) })),
        note: "「工作表」列是超链接，点一下跳到对应表；增删表后再调用一次即可原地刷新。",
      },
      summary: `已${refreshed ? "刷新" : "生成"}目录「${ws.name}」，列出 ${entries.length} 张表`,
    };
  });
}
