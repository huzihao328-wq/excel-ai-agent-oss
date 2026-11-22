/* global Excel */
import { ToolResult } from "../types";
import { pushOpLog } from "../oplog";
import { registerArtifact } from "../artifacts";
import { isInternalSheet } from "./checkpoint";

function sanitizeName(name: string, fallback: string): string {
  const cleaned = name.replace(/[\\/?*[\]:]/g, "_").replace(/\s+/g, "_").trim();
  const base = cleaned || fallback;
  return base.length > 31 ? base.slice(0, 31) : base;
}

function sanitizeTableName(name: string): string {
  let s = name.replace(/[^\w\u4e00-\u9fff]/g, "_");
  if (!s || !/^[A-Za-z_\u4e00-\u9fff]/.test(s)) s = `T_${s}`;
  return s.slice(0, 200);
}

async function uniqueSheetName(context: Excel.RequestContext, baseName: string): Promise<string> {
  const wsCol = context.workbook.worksheets;
  wsCol.load("items/name");
  await context.sync();
  const existing = new Set(wsCol.items.map((w) => w.name.toLowerCase()));
  let name = sanitizeName(baseName, "透视");
  let i = 2;
  while (existing.has(name.toLowerCase())) {
    name = sanitizeName(`${baseName.slice(0, 27)}(${i})`, `透视${i}`);
    i++;
  }
  return name;
}

export { uniqueSheetName };

async function uniqueTableName(context: Excel.RequestContext, baseName: string): Promise<string> {
  const tables = context.workbook.tables;
  tables.load("items/name");
  await context.sync();
  const existing = new Set(tables.items.map((t) => t.name));
  let name = sanitizeTableName(baseName);
  let i = 2;
  while (existing.has(name)) {
    name = sanitizeTableName(`${baseName}_${i}`);
    i++;
  }
  return name;
}

export async function toolListTables(): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    for (const ws of sheets.items) ws.tables.load("items/name");
    await context.sync();

    const pending: { sheet: string; name: string; range: Excel.Range; columns: Excel.TableColumnCollection }[] = [];
    for (const ws of sheets.items) {
      if (isInternalSheet(ws.name)) continue;
      for (const t of ws.tables.items) {
        const range = t.getRange();
        range.load("address");
        t.columns.load("items/name");
        pending.push({ sheet: ws.name, name: t.name, range, columns: t.columns });
      }
    }
    await context.sync();

    const tables = pending.map((t) => ({
      sheet: t.sheet,
      name: t.name,
      range: t.range.address.split("!").pop(),
      columns: t.columns.items.map((c) => c.name),
    }));
    return {
      ok: true,
      data: { tables },
      summary: tables.length ? `共 ${tables.length} 个 Excel 表` : "当前工作簿没有 Excel 表（ListObject）",
    };
  });
}

export async function toolCreateTable(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,isNullObject,rowCount,columnCount");
    await context.sync();
    if ((range as any).isNullObject) {
      return { ok: false, error: `${args.sheet} 是空表，没有可转换的数据区域` };
    }
    const name = await uniqueTableName(context, args.name || `${args.sheet}_表`);
    const table = sheet.tables.add(range, args.has_headers !== false);
    table.name = name;
    table.getRange().format.autofitColumns();
    await context.sync();
    const tableAddr = range.address.split("!").pop() as string;
    pushOpLog({
      description: `把 ${args.sheet}!${tableAddr} 转成 Excel 表「${name}」`,
      kind: "table",
      sheet: args.sheet,
      address: tableAddr,
    });
    registerArtifact({
      name,
      kind: "table",
      sheet: args.sheet,
      address: tableAddr,
      brief: `Excel 表（${args.sheet}!${tableAddr}）`,
    });
    return {
      ok: true,
      data: { table: name, sheet: args.sheet, range: tableAddr },
      summary: `已把 ${args.sheet}!${tableAddr} 转成 Excel 表「${name}」`,
    };
  });
}

function mapFilterOn(kind: string | undefined): Excel.FilterOn {
  const k = String(kind || "values").toLowerCase();
  if (k === "custom") return Excel.FilterOn.custom;
  if (k === "topitems") return Excel.FilterOn.topItems;
  if (k === "bottomitems") return Excel.FilterOn.bottomItems;
  return Excel.FilterOn.values;
}

export async function toolApplyFilter(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    const rawIndex = Number(args.column_index ?? 0);
    const columnIndex = Math.trunc(rawIndex);
    if (!Number.isFinite(rawIndex) || columnIndex < 0 || columnIndex !== rawIndex) {
      return { ok: false, error: `column_index 必须是 ≥0 的整数，收到: ${JSON.stringify(args.column_index)}` };
    }
    const range = args.range ? sheet.getRange(args.range) : sheet.getUsedRangeOrNullObject(true);
    range.load("address,isNullObject,columnCount");
    await context.sync();
    if ((range as any).isNullObject) {
      return { ok: false, error: `${args.sheet} 没有可筛选的数据` };
    }
    if (columnIndex >= range.columnCount) {
      return { ok: false, error: `column_index ${columnIndex} 超出区域列数 ${range.columnCount}` };
    }

    const criteria: Excel.FilterCriteria = { filterOn: mapFilterOn(args.filter_on) };
    if (Array.isArray(args.values) && args.values.length) {
      criteria.filterOn = Excel.FilterOn.values;
      criteria.values = args.values.map((v: unknown) => String(v));
    }
    if (args.criterion1 != null) {
      criteria.criterion1 = String(args.criterion1);
      if (!Array.isArray(args.values) || !args.values.length) criteria.filterOn = Excel.FilterOn.custom;
    }
    if (args.criterion2 != null) criteria.criterion2 = String(args.criterion2);

    sheet.autoFilter.apply(range, columnIndex, criteria);
    await context.sync();
    const addr = range.address.split("!").pop();
    return {
      ok: true,
      data: { sheet: args.sheet, range: addr, column_index: columnIndex },
      summary: `已在 ${args.sheet}!${addr} 第 ${columnIndex + 1} 列加上筛选`,
    };
  });
}

export async function toolClearFilter(args: any): Promise<ToolResult> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(args.sheet);
    if (args.remove) sheet.autoFilter.remove();
    else sheet.autoFilter.clearCriteria();
    await context.sync();
    return {
      ok: true,
      data: { sheet: args.sheet, removed: !!args.remove },
      summary: args.remove ? `已移除 ${args.sheet} 的自动筛选` : `已清除 ${args.sheet} 的筛选条件`,
    };
  });
}

function mapAggregation(name: string | undefined): Excel.AggregationFunction {
  switch (String(name || "sum").toLowerCase()) {
    case "count":
      return Excel.AggregationFunction.count;
    case "average":
    case "avg":
      return Excel.AggregationFunction.average;
    case "max":
      return Excel.AggregationFunction.max;
    case "min":
      return Excel.AggregationFunction.min;
    case "countnumbers":
      return Excel.AggregationFunction.countNumbers;
    default:
      return Excel.AggregationFunction.sum;
  }
}

export async function toolCreatePivot(args: any): Promise<ToolResult> {
  const values = Array.isArray(args.values) ? args.values : [];
  const valueFields = values
    .filter((raw: unknown) => raw != null)
    .map((raw: any) => (typeof raw === "string" ? raw : String(raw.field || raw.name || "")))
    .filter((field: string) => field.length > 0);
  if (valueFields.length === 0) {
    return { ok: false, error: "create_pivot_table 至少需要一个 values 字段" };
  }

  return Excel.run(async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(args.source_sheet);
    const sourceRange = sourceSheet.getRange(args.source_range);
    sourceRange.load("address");
    await context.sync();

    const destName = args.target_sheet
      ? String(args.target_sheet)
      : await uniqueSheetName(context, args.name || "透视");
    // 注意：新建分支里 destSheet.name 未 load，后续一律用已知的 destName，
    // 避免读取未加载属性抛 PropertyNotLoaded（透视表建好了却报失败）。
    let destSheet: Excel.Worksheet;
    let createdDest = false;
    if (args.target_sheet) {
      try {
        destSheet = context.workbook.worksheets.getItem(destName);
        destSheet.load("name");
        await context.sync();
      } catch {
        destSheet = context.workbook.worksheets.add(destName);
        createdDest = true;
      }
    } else {
      destSheet = context.workbook.worksheets.add(destName);
      createdDest = true;
    }

    const destCell = destSheet.getRange(args.target_cell || "A3");
    const pivotName = sanitizeTableName(args.name || `Pivot_${Date.now().toString(36)}`);
    const pivotTable = destSheet.pivotTables.add(pivotName, sourceRange, destCell);
    await context.sync();

    try {
      pivotTable.hierarchies.load("items/name");
      await context.sync();
      const available = pivotTable.hierarchies.items.map((h) => h.name);
      const findHier = (field: string) => {
        const exact = pivotTable.hierarchies.items.find((h) => h.name === field);
        if (exact) return exact;
        throw new Error(`透视字段「${field}」不存在。可用字段: ${available.join("、") || "(无)"}`);
      };

      for (const row of args.rows || []) pivotTable.rowHierarchies.add(findHier(String(row)));
      for (const col of args.columns || []) pivotTable.columnHierarchies.add(findHier(String(col)));

      for (const raw of values) {
        if (raw == null) continue;
        const field = typeof raw === "string" ? raw : String(raw.field || raw.name || "");
        const agg = typeof raw === "string" ? args.aggregation : raw.aggregation;
        if (!field) continue;
        const dh = pivotTable.dataHierarchies.add(findHier(field));
        dh.summarizeBy = mapAggregation(agg);
      }

      destSheet.activate();
      await context.sync();
      pushOpLog({
        description: `在 Sheet「${destName}」创建透视表「${pivotName}」（源 ${args.source_sheet}）`,
        kind: "table",
        sheet: destName,
        address: args.target_cell || "A3",
      });
      registerArtifact({
        name: args.name || pivotName,
        kind: "pivot",
        sheet: destName,
        address: args.target_cell || "A3",
        brief: `透视表（行: ${(args.rows || []).join("、") || "无"}，值: ${valueFields.length} 个字段）`,
      });
      return {
        ok: true,
        data: {
          pivot: pivotName,
          sheet: destName,
          source: `${args.source_sheet}!${sourceRange.address.split("!").pop()}`,
          fields: available,
        },
        summary: `已在 Sheet「${destName}」创建透视表「${pivotName}」`,
      };
    } catch (e) {
      try {
        destSheet.pivotTables.getItem(pivotName).delete();
        if (createdDest) destSheet.delete();
        await context.sync();
      } catch {
        /* 尽力清理半成品 */
      }
      throw e;
    }
  });
}
