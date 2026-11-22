/** 统一的线性 SVG 图标，替代 emoji，保证各平台渲染一致 */
import { ReactNode } from "react";

interface IconProps {
  size?: number;
  className?: string;
}

function Svg({ size = 16, className, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.7}
      className={className}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  );
}

export function IconPlus(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 5v14M5 12h14" />
    </Svg>
  );
}

/** 历史对话：对话气泡 */
export function IconChats(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M21 11.5a8.4 8.4 0 0 1-9 8.4 8.9 8.9 0 0 1-3.8-.8L3 20l1-4.1a8.3 8.3 0 0 1-1-4.4 8.4 8.4 0 0 1 9-8.4 8.4 8.4 0 0 1 9 8.4Z" />
      <path d="M8.5 10h7M8.5 13.5h4.5" />
    </Svg>
  );
}

/** 操作记录：带回退箭头的时钟 */
export function IconHistory(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M3 12a9 9 0 1 0 2.6-6.3" />
      <path d="M3 4v4h4" />
      <path d="M12 8v4l2.8 2.2" />
    </Svg>
  );
}

/** 批量任务：层叠文件 */
export function IconLayers(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m12 3 9 5-9 5-9-5 9-5Z" />
      <path d="m3.6 12.6 8.4 4.7 8.4-4.7" />
      <path d="m3.6 16.6 8.4 4.7 8.4-4.7" />
    </Svg>
  );
}

/** 记忆：书签 */
export function IconBookmark(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M6 4.5A1.5 1.5 0 0 1 7.5 3h9A1.5 1.5 0 0 1 18 4.5V21l-6-4-6 4V4.5Z" />
    </Svg>
  );
}

/** 技能库：闪电 */
/** 点了直接执行（区别于只填入输入框） */
export function IconPlay(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M7 4.5v15l12.5-7.5L7 4.5Z" />
    </Svg>
  );
}

export function IconZap(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M13 2 4.5 13.5H11L10 22l8.5-11.5H12L13 2Z" />
    </Svg>
  );
}

/** 常用指令：带星标的卡片 */
export function IconTemplate(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="3" y="4" width="18" height="16" rx="2.5" />
      <path d="M7 8h6M7 12h4" />
      <path d="m16.5 11 .8 1.7 1.7.8-1.7.8-.8 1.7-.8-1.7-1.7-.8 1.7-.8.8-1.7Z" />
    </Svg>
  );
}

export function IconGear(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="12" cy="12" r="3.2" />
      <path d="M19.4 15a1.7 1.7 0 0 0 .34 1.87l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.7 1.7 0 0 0-1.87-.34 1.7 1.7 0 0 0-1.03 1.56V21a2 2 0 1 1-4 0v-.09A1.7 1.7 0 0 0 8.9 19.4a1.7 1.7 0 0 0-1.87.34l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.7 1.7 0 0 0 .34-1.87 1.7 1.7 0 0 0-1.56-1.03H3a2 2 0 1 1 0-4h.09A1.7 1.7 0 0 0 4.6 8.9a1.7 1.7 0 0 0-.34-1.87l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.7 1.7 0 0 0 1.87.34H9a1.7 1.7 0 0 0 1.03-1.56V3a2 2 0 1 1 4 0v.09c0 .68.4 1.3 1.03 1.56a1.7 1.7 0 0 0 1.87-.34l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.7 1.7 0 0 0-.34 1.87V9c.26.63.88 1.03 1.56 1.03H21a2 2 0 1 1 0 4h-.09a1.7 1.7 0 0 0-1.56 1.03Z" />
    </Svg>
  );
}

export function IconDots(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="5" cy="12" r="1" fill="currentColor" />
      <circle cx="12" cy="12" r="1" fill="currentColor" />
      <circle cx="19" cy="12" r="1" fill="currentColor" />
    </Svg>
  );
}

export function IconSend(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M22 2 11 13" />
      <path d="M22 2 15 22l-4-9-9-4 20-7Z" />
    </Svg>
  );
}

export function IconStop(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="6.5" y="6.5" width="11" height="11" rx="1.8" fill="currentColor" stroke="none" />
    </Svg>
  );
}

export function IconClip(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m21.4 11.05-8.9 8.9a5.5 5.5 0 0 1-7.78-7.78l8.9-8.9a3.67 3.67 0 0 1 5.19 5.19l-8.92 8.9a1.83 1.83 0 0 1-2.59-2.6l8.22-8.2" />
    </Svg>
  );
}

export function IconDownload(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 3v12" />
      <path d="m8 11 4 4 4-4" />
      <path d="M5 19h14" />
    </Svg>
  );
}

export function IconClose(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M18 6 6 18M6 6l12 12" />
    </Svg>
  );
}

export function IconSparkle(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 3.5c.5 3.9 2.6 6 6.5 6.5-3.9.5-6 2.6-6.5 6.5-.5-3.9-2.6-6-6.5-6.5 3.9-.5 6-2.6 6.5-6.5Z" fill="currentColor" stroke="none" />
      <path d="M18.5 15.5c.2 1.4.9 2.1 2.3 2.3-1.4.2-2.1.9-2.3 2.3-.2-1.4-.9-2.1-2.3-2.3 1.4-.2 2.1-.9 2.3-2.3Z" fill="currentColor" stroke="none" />
    </Svg>
  );
}

/** 工作表：网格 */
export function IconSheet(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="3.5" y="4" width="17" height="16" rx="2.5" />
      <path d="M3.5 9.5h17M3.5 14.5h17M9.5 9.5V20" />
    </Svg>
  );
}

/** 选区：虚线框 */
export function IconSelection(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M5 4h2M11 4h2M17 4h2M5 20h2M11 20h2M17 20h2M4 5v2M4 11v2M4 17v2M20 5v2M20 11v2M20 17v2" />
    </Svg>
  );
}

export function IconChart(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M4 20h16" />
      <rect x="6" y="11" width="3" height="6" rx="1" />
      <rect x="11" y="6" width="3" height="11" rx="1" />
      <rect x="16" y="9" width="3" height="8" rx="1" />
    </Svg>
  );
}

export function IconTrend(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m3.5 16.5 5.5-5.5 4 4 7.5-7.5" />
      <path d="M15.5 7.5h5v5" />
    </Svg>
  );
}

export function IconPivot(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="3.5" y="3.5" width="17" height="17" rx="2.5" />
      <path d="M3.5 9h17M9 3.5v17" />
      <path d="m13 15 2 2 2-2M15 12.5V17" />
    </Svg>
  );
}

export function IconFile(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M14 3.5H7.5a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h9a2 2 0 0 0 2-2V8Z" />
      <path d="M14 3.5V8h4.5" />
    </Svg>
  );
}

export function IconCheck(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m5 12.5 4.5 4.5L19 7.5" />
    </Svg>
  );
}

export function IconAlert(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M10.3 4.2 2.9 17a2 2 0 0 0 1.7 3h14.8a2 2 0 0 0 1.7-3L13.7 4.2a2 2 0 0 0-3.4 0Z" />
      <path d="M12 9.5v4M12 16.8v.2" />
    </Svg>
  );
}

export function IconInfo(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="12" cy="12" r="8.5" />
      <path d="M12 11v5M12 7.8v.2" />
    </Svg>
  );
}

export function IconCompass(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="12" cy="12" r="8.5" />
      <path d="m15.5 8.5-2 5-5 2 2-5 5-2Z" />
    </Svg>
  );
}

/** 数据体检：心电脉搏 */
export function IconPulse(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M3 12h4l2.5-6 5 12 2.5-6h4" />
    </Svg>
  );
}

/** 清洗：魔术棒 */
export function IconWand(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m4 20 11-11" />
      <path d="m13.5 7.5 3 3" />
      <path d="M18 3v3M16.5 4.5h3M20 9v2M19 10h2M9.5 3v2M8.5 4h2" />
    </Svg>
  );
}

export function IconLink(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1" />
      <path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1" />
    </Svg>
  );
}

export function IconFunction(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M15 4.5c-2.5-.5-4 .5-4.5 3L8.5 18c-.5 2.2-1.8 2.8-3.5 2.3" />
      <path d="M7 10h7" />
      <path d="m14 13 5 5M19 13l-5 5" />
    </Svg>
  );
}

export function IconPin(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M9 3.5h6l-1 5 3.5 3.5v1.5h-11V12L10 8.5Z" />
      <path d="M12 13.5V21" />
    </Svg>
  );
}

export function IconBrush(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M18.5 3.5 10 12l2 2 8.5-8.5a1.4 1.4 0 0 0-2-2Z" />
      <path d="M10 12c-2 0-3.5 1.5-3.5 3.5 0 1.5-1 2.5-3 3 4.5 1.5 8.5 0 8.5-4.5Z" />
    </Svg>
  );
}

export function IconCalendar(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="3.5" y="5" width="17" height="15.5" rx="2.5" />
      <path d="M3.5 10h17M8 3v4M16 3v4" />
    </Svg>
  );
}

export function IconScale(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 4v16M7 20h10M5 7h14" />
      <path d="m5 7-2.5 6a2.5 2.5 0 0 0 5 0Z" />
      <path d="m19 7-2.5 6a2.5 2.5 0 0 0 5 0Z" />
    </Svg>
  );
}

export function IconCoin(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="12" cy="12" r="8.5" />
      <path d="M14.8 9.2c-.5-.9-1.5-1.4-2.8-1.4-1.6 0-2.8.8-2.8 2.1 0 3 5.8 1.4 5.8 4.3 0 1.3-1.3 2.2-3 2.2-1.4 0-2.5-.6-3-1.6M12 6v1.8M12 16.4V18" />
    </Svg>
  );
}

export function IconUsers(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="9" cy="8.5" r="3.5" />
      <path d="M3 19.5c.6-3.2 3-5 6-5s5.4 1.8 6 5" />
      <path d="M16 5.2a3.5 3.5 0 0 1 0 6.6M18 14.8c1.6.7 2.7 2.3 3 4.7" />
    </Svg>
  );
}

export function IconFolder(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M3.5 7.5a2 2 0 0 1 2-2h4l2 2.5h7a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2Z" />
    </Svg>
  );
}

export function IconStar(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m12 3.8 2.5 5.1 5.6.8-4 4 .9 5.6-5-2.7-5 2.7.9-5.6-4-4 5.6-.8Z" />
    </Svg>
  );
}

export function IconBulb(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M9 17.5h6M10 21h4" />
      <path d="M12 3a6 6 0 0 0-3.6 10.8c.6.5 1 1.2 1 2V17.5h5.2v-1.7c0-.8.4-1.5 1-2A6 6 0 0 0 12 3Z" />
    </Svg>
  );
}

export function IconCopy(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="8.5" y="8.5" width="12" height="12" rx="2.5" />
      <path d="M15.5 8.5V6a2.5 2.5 0 0 0-2.5-2.5H6A2.5 2.5 0 0 0 3.5 6v7A2.5 2.5 0 0 0 6 15.5h2.5" />
    </Svg>
  );
}

export function IconChevronDown(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m6 9.5 6 6 6-6" />
    </Svg>
  );
}

export function IconChevronRight(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="m9.5 6 6 6-6 6" />
    </Svg>
  );
}

export function IconArrowDown(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 5v14M6 13l6 6 6-6" />
    </Svg>
  );
}

export function IconArrowUpRight(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M7 17 17 7M8.5 7H17v8.5" />
    </Svg>
  );
}

export function IconArrowRight(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M5 12h14M13 6l6 6-6 6" />
    </Svg>
  );
}

export function IconUndo(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M9 14 4 9l5-5" />
      <path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11" />
    </Svg>
  );
}

export function IconSearch(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="11" cy="11" r="6.5" />
      <path d="m20 20-4.2-4.2" />
    </Svg>
  );
}

export function IconTrash(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M4 7h16M10 11v6M14 11v6" />
      <path d="M5.5 7 6.5 19a2 2 0 0 0 2 1.8h7a2 2 0 0 0 2-1.8L18.5 7" />
      <path d="M9 7V4.5h6V7" />
    </Svg>
  );
}

export function IconUpload(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 16V4M7.5 8.5 12 4l4.5 4.5" />
      <path d="M4 15v3.5A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5V15" />
    </Svg>
  );
}

export function IconClock(p: IconProps) {
  return (
    <Svg {...p}>
      <circle cx="12" cy="12" r="8.5" />
      <path d="M12 7.5V12l3 2" />
    </Svg>
  );
}

export function IconSliders(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M4 7h9M17 7h3M4 17h3M11 17h9" />
      <circle cx="15" cy="7" r="2" />
      <circle cx="9" cy="17" r="2" />
    </Svg>
  );
}

export function IconShield(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M12 3.5 5 6v5.5c0 4.3 2.9 7.8 7 9 4.1-1.2 7-4.7 7-9V6Z" />
      <path d="m9 12 2 2 4-4" />
    </Svg>
  );
}

export function IconServer(p: IconProps) {
  return (
    <Svg {...p}>
      <rect x="3.5" y="4" width="17" height="7" rx="2" />
      <rect x="3.5" y="13" width="17" height="7" rx="2" />
      <path d="M7.5 7.5h.01M7.5 16.5h.01" />
    </Svg>
  );
}

export function IconSigma(p: IconProps) {
  return (
    <Svg {...p}>
      <path d="M17.5 5.5V4h-11l6 8-6 8h11v-1.5" />
    </Svg>
  );
}

type IconComponent = (p: IconProps) => JSX.Element;

/**
 * 数据里保存的是 emoji（内置 / 用户自建的常用指令、快捷任务、旧对话里的提示），
 * 渲染时统一映射为线性图标；未知字符退回通用图标，避免各平台 emoji 风格不一。
 */
const GLYPHS: Record<string, IconComponent> = {
  "📊": IconChart, "📈": IconTrend, "📉": IconTrend, "⊞": IconPivot, "⧉": IconPivot, "📄": IconFile, "📎": IconClip,
  "🩺": IconPulse, "🧹": IconWand, "✨": IconBrush, "🔗": IconLink, "📌": IconPin, "⬇": IconDownload, "⬚": IconSelection,
  "📅": IconCalendar, "🗓": IconCalendar, "🧭": IconCompass, "⚖": IconScale, "💰": IconCoin, "👥": IconUsers,
  "🗂": IconFolder, "★": IconStar, "⭐": IconStar, "🧠": IconBulb, "🕘": IconClock, "ƒx": IconFunction, "ƒ": IconFunction, "∑": IconSigma,
  "⚠": IconAlert, "✓": IconCheck, "✅": IconCheck,
};
const stripVariation = (text: string) => text.replace(/️/g, "");

export function GlyphIcon({ glyph, size, className }: IconProps & { glyph?: string }) {
  const Icon = (glyph && GLYPHS[stripVariation(glyph).trim()]) || IconStar;
  return <Icon size={size} className={className} />;
}

/** 拆出文本开头的 emoji：返回对应图标和去掉 emoji 的正文（兼容旧对话里保存的提示文字） */
export function splitLeadingGlyph(text: string): { Icon?: IconComponent; rest: string } {
  const clean = stripVariation(text);
  for (const [glyph, Icon] of Object.entries(GLYPHS)) {
    if (clean.startsWith(glyph)) return { Icon, rest: clean.slice(glyph.length).trimStart() };
  }
  return { rest: text };
}
