export const QUICK_ACTIONS: { icon: string; label: string; prompt: string }[] = [
  { icon: "📊", label: "分析当前表", prompt: "分析当前活动 Sheet 的数据：概述数据结构，给出关键统计指标和你发现的规律或异常。" },
  { icon: "🩺", label: "数据体检", prompt: "对当前活动 Sheet 做一次数据体检：检查缺失值、重复行、类型不一致、异常值和多余空格，把报告写到新 Sheet，并在回复里总结主要问题。" },
  { icon: "🧹", label: "数据清洗", prompt: "检查当前活动 Sheet 的数据质量问题（空值、重复行、格式不一致、异常值），列出发现的问题，然后把清洗后的数据写入新 Sheet。" },
  { icon: "🔗", label: "跨表合并", prompt: "先查看各 Sheet 结构，按关联键把相关表合并到新 Sheet（用查找公式匹配，不要把整表读出来手填）。告诉我你计划用哪个字段关联，确认后执行。" },
  { icon: "📈", label: "生成图表", prompt: "根据当前活动 Sheet 的数据，选择最合适的图表类型创建一个图表，并说明你的选择理由。" },
  { icon: "ƒx", label: "写公式", prompt: "我需要一个公式来解决以下问题（请根据我的描述在合适的位置写入公式并验证结果）：" },
  { icon: "⊞", label: "做透视表", prompt: "根据当前活动 Sheet 的数据做一个透视表，放到新 Sheet。先看表头，选合适的行/列/值字段，告诉我你的选择后执行。" },
  { icon: "∑", label: "列统计", prompt: "不要把整表读进对话。请用 column_stats 统计当前活动 Sheet 各列的非空数、空值、数值个数、合计/最小/最大/平均和唯一值个数，金额类列重点汇报。" },
  { icon: "⬇", label: "导出文件", prompt: "把当前工作簿导出成 Excel 文件供我下载。" },
  { icon: "📌", label: "冻结表头", prompt: "把当前活动 Sheet 的首行冻结，往下滚动时仍能看到表头。" },
  { icon: "✨", label: "整理格式", prompt: "把当前活动 Sheet 整理成正规表：表头加粗配色、边框、列宽自适应、冻结首行，金额/日期列套上合适的数字格式。" },
];

