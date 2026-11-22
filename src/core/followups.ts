import { FollowupAction } from "./types";

export function heuristicFollowups(toolNames: string[]): FollowupAction[] {
  const used = new Set(toolNames.filter((n) => n !== "ask_user_choice" && n !== "suggest_followups"));
  if (used.size === 0) return [];
  const actions: FollowupAction[] = [];
  const add = (label: string, prompt: string) => {
    if (actions.length >= 3 || actions.some((a) => a.label === label)) return;
    actions.push({ label, prompt });
  };
  if (used.has("create_pivot_table") || used.has("group_summary") || used.has("write_range") || used.has("create_table") || used.has("create_sheet")) {
    add("做成图表", "根据刚才的结果选择最合适的图表类型画出来，并说明理由。");
  }
  if (used.has("reconcile_tables")) {
    add("汇总差异", "把刚才对账结果里不一致的记录按差异列汇总条数和差额，列出差额最大的几条。");
  }
  if (used.has("stack_sheets")) {
    add("按来源汇总", "对刚才合并的表按来源分组汇总主要指标，写到新 Sheet。");
  }
  if (used.has("unpivot_columns")) {
    add("按属性汇总", "对刚才转好的长表按属性列汇总值，写到新 Sheet，再画一张趋势或对比图。");
  }
  if (used.has("compare_sheets")) {
    add("汇总改动", "按列统计刚才版本差异里的改动次数，列出改动最多的列和行，并说明新增、删除了多少行。");
  }
  if (used.has("fill_blank_cells")) {
    add("按分类汇总", "刚才已经把分类列补齐了，现在按这些分类汇总主要指标，写到新 Sheet。");
  }
  if (!used.has("create_pivot_table") && (used.has("write_range") || used.has("create_table") || used.has("import_external_sheet"))) {
    add("做透视表", "根据刚才产出的数据做一个透视表，放到新 Sheet。行/列/值字段按数据含义选择。");
  }
  if (used.has("create_chart")) {
    add("解读图表", "根据刚才的图表写一段简短的数据解读，标出最值得注意的点。");
  }
  if (!used.has("apply_autofilter") && (used.has("create_table") || used.has("write_range") || used.has("get_workbook_overview"))) {
    add("加上筛选", "把当前相关数据转成 Excel 表（如果还不是）并加上自动筛选。");
  }
  if (used.has("read_range") || used.has("inspect_range") || used.has("get_workbook_overview")) {
    add("清洗数据", "检查刚才涉及的数据有没有空值、重复或格式问题，清洗后写入新 Sheet。");
  }
  if (used.has("column_stats") || used.has("get_workbook_overview")) {
    add("做成图表", "根据刚才的统计选择最合适的图表类型画出来，并说明理由。");
  }
  if (used.has("copy_range") || used.has("fill_down") || used.has("add_lookup_column") || used.has("write_range") || used.has("create_sheet")) {
    add("整理格式", "用 polish_range 把刚才的结果整理成正规表：表头加粗、边框、列宽自适应、冻结首行。");
  }
  if (used.has("add_lookup_column") || used.has("set_formulas") || used.has("fill_down")) {
    add("转成数值", "把刚才写入的公式转成当前显示值，不要再保留公式。");
  }
  if (used.has("split_column") || used.has("combine_columns") || used.has("clean_column") || used.has("extract_text")) {
    add("再做个体检", "对刚才清洗过的表做一次数据体检，报告写到新 Sheet。");
  }
  if (used.has("remove_duplicates") || used.has("find_replace") || used.has("delete_blank_rows")) {
    add("再做个体检", "对刚才处理过的表做一次数据体检，报告写到新 Sheet。");
  }
  if (!used.has("export_file") && (used.has("write_range") || used.has("create_sheet") || used.has("create_pivot_table") || used.has("import_external_sheet") || used.has("add_lookup_column") || used.has("split_column") || used.has("polish_range") || used.has("group_summary") || used.has("reconcile_tables") || used.has("stack_sheets") || used.has("split_by_column") || used.has("unpivot_columns") || used.has("compare_sheets"))) {
    add("导出文件", "把刚才的结果导出成 Excel 文件供我下载。");
  }
  if (actions.length < 2) {
    add("做成汇总表", "把当前活动表按合适的维度汇总到新 Sheet。");
    add("生成图表", "根据当前活动表做一张合适的图表。");
  }
  return actions.slice(0, 3);
}

