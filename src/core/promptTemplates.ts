import { TEMPLATES_STORAGE_KEY, schedulePersistUserData } from "./durableStore";

export interface PromptTemplate {
  id: string;
  name: string;
  prompt: string;
  category: string;
  icon: string;
  builtIn: boolean;
  createdAt: number;
  updatedAt: number;
}

const MAX_CUSTOM_TEMPLATES = 30;

const BUILTIN_TEMPLATES: PromptTemplate[] = [
  {
    id: "builtin-weekly-report",
    name: "周报汇总",
    prompt:
      "根据当前活动 Sheet 生成一份周报：汇总本周核心指标、与上周对比、列出变化最大的项目和异常，并把汇总表与合适的图表放到新 Sheet。",
    category: "汇总",
    icon: "📅",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-management-summary",
    name: "管理层摘要",
    prompt:
      "分析当前活动 Sheet，用简洁的管理层语言给出：3 个最重要结论、关键数字、主要风险和建议动作；需要的汇总数据写到新 Sheet。",
    category: "分析",
    icon: "🧭",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-monthly-compare",
    name: "月度环比",
    prompt:
      "识别当前表中的日期、维度和指标列，按月汇总主要指标，计算环比变化和变化率，标出增长与下降最大的项目，并在新 Sheet 生成汇总表和趋势图。",
    category: "分析",
    icon: "📈",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-reconcile",
    name: "两表核对",
    prompt:
      "先查看各 Sheet 结构，找出最适合的唯一键，核对两张相关表的全部记录：列出仅在左表、仅在右表、关键字段不一致的记录，结果写到新的核对 Sheet，并总结差异数量和金额差。",
    category: "核对",
    icon: "⚖",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-invoice-check",
    name: "金额核对",
    prompt:
      "检查当前活动 Sheet 的金额列：统计总额、空值、负数、重复单号、金额为零和明显异常值，把可疑记录写到新 Sheet，并按问题类型汇总数量与金额。",
    category: "核对",
    icon: "💰",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-clean-roster",
    name: "名单清洗",
    prompt:
      "清洗当前活动 Sheet 的名单数据：统一姓名和联系方式的空格与格式，识别重复人员，保留原始数据不变，把清洗结果和重复项清单分别放到新 Sheet。",
    category: "清洗",
    icon: "👥",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-split-master",
    name: "按类别拆表",
    prompt:
      "查看当前活动 Sheet 的表头，询问我或判断最合适的分类列，然后按该列的不同值拆分到多个新 Sheet；保留表头、格式和列宽，并汇报每个 Sheet 的记录数。",
    category: "整理",
    icon: "🗂",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-unmerge-fill",
    name: "拆合并补空白",
    prompt:
      "当前活动 Sheet 是从系统或报表里导出的，分类列用了合并单元格、下面留着空白。请取消合并并把每一格补上所属的值，让每行都完整，方便筛选、排序和做透视；改之前告诉我会动哪些列。",
    category: "整理",
    icon: "🧩",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-unpivot",
    name: "宽表转长表",
    prompt:
      "当前活动 Sheet 是横着排的宽表（如 1月～12月、各门店各占一列）。请先看表头，判断哪些是标识列、哪些是要转成行的数值列，转成「标识列 + 属性 + 值」的长表放到新 Sheet，并给属性列和值列起合适的名字。",
    category: "整理",
    icon: "🔃",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-id-card",
    name: "身份证信息",
    prompt:
      "找出当前活动 Sheet 里的身份证号列，在右侧新增出生日期、性别、年龄和身份证校验四列；汇报有多少个号码校验不通过、有没有被存成数字丢了位数的。",
    category: "清洗",
    icon: "🪪",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-compare-versions",
    name: "版本比对",
    prompt:
      "比较工作簿里同一张表的两个版本（旧版和新版，不确定是哪两张就先问我），找出改了哪些格、新增和删除了哪些行；有唯一编号列就按编号对齐。差异清单写到新 Sheet，并在新版本上把改动标出颜色。",
    category: "核对",
    icon: "🔍",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-print-ready",
    name: "打印排版",
    prompt:
      "把当前活动 Sheet 设置成适合打印的样子：列多就横向、所有列缩到一页宽，每页重复打印表头行，A4 纸、窄边距、水平居中，页脚加「第 X 页，共 Y 页」；表格本身没整理过的话顺便整理格式。",
    category: "整理",
    icon: "🖨",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
  {
    id: "builtin-dashboard-source",
    name: "仪表盘底表",
    prompt:
      "把当前活动 Sheet 整理成适合制作仪表盘的汇总底表：识别日期、分类和指标字段，按合理粒度聚合到新 Sheet，再创建能展示趋势与构成的图表。",
    category: "汇总",
    icon: "📊",
    builtIn: true,
    createdAt: 0,
    updatedAt: 0,
  },
];

let listeners: Array<() => void> = [];

function notify(): void {
  listeners.forEach((fn) => fn());
}

export function onPromptTemplatesChange(fn: () => void): () => void {
  listeners.push(fn);
  return () => {
    listeners = listeners.filter((item) => item !== fn);
  };
}

export function loadCustomPromptTemplates(): PromptTemplate[] {
  try {
    const raw = JSON.parse(localStorage.getItem(TEMPLATES_STORAGE_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return raw
      .filter((item) => item && typeof item.id === "string" && typeof item.name === "string" && typeof item.prompt === "string")
      .map((item) => ({
        ...item,
        category: "我的",
        icon: typeof item.icon === "string" && item.icon ? item.icon : "★",
        builtIn: false,
      }));
  } catch {
    return [];
  }
}

function saveCustomPromptTemplates(items: PromptTemplate[]): void {
  localStorage.setItem(TEMPLATES_STORAGE_KEY, JSON.stringify(items.slice(0, MAX_CUSTOM_TEMPLATES)));
  schedulePersistUserData();
  notify();
}

export function listPromptTemplates(): PromptTemplate[] {
  return [...loadCustomPromptTemplates(), ...BUILTIN_TEMPLATES];
}

export function upsertPromptTemplate(input: { id?: string; name: string; prompt: string }): PromptTemplate {
  const name = input.name.trim().slice(0, 30);
  const prompt = input.prompt.trim().slice(0, 3000);
  if (!name || !prompt) throw new Error("请填写指令名称和完整内容");

  const items = loadCustomPromptTemplates();
  const now = Date.now();
  const found = input.id ? items.find((item) => item.id === input.id) : undefined;
  if (found) {
    found.name = name;
    found.prompt = prompt;
    found.updatedAt = now;
    saveCustomPromptTemplates([found, ...items.filter((item) => item.id !== found.id)]);
    return found;
  }

  const item: PromptTemplate = {
    id: `tpl_${now.toString(36)}_${Math.random().toString(36).slice(2, 7)}`,
    name,
    prompt,
    category: "我的",
    icon: "★",
    builtIn: false,
    createdAt: now,
    updatedAt: now,
  };
  saveCustomPromptTemplates([item, ...items]);
  return item;
}

export function removePromptTemplate(id: string): void {
  saveCustomPromptTemplates(loadCustomPromptTemplates().filter((item) => item.id !== id));
}
