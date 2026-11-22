/*
 * 让工具调用「一次就对」的两件小事：
 * 1. 模型常把地址写成 "Sheet1!A1:C9" 而 sheet 参数另给或漏给——能确定意图时就地拆开；
 * 2. 表名写错时 Office 只报 ItemNotFound，模型得再跑一轮 overview——这里直接给出现有表名和最接近的那个。
 */

/** 地址参数 → 对应的 Sheet 参数 */
const ADDRESS_SHEET_KEYS: Record<string, string> = {
  range: "sheet",
  start_cell: "sheet",
  cell: "sheet",
  address: "sheet",
  source_range: "source_sheet",
  dest_range: "dest_sheet",
  dest_cell: "dest_sheet",
  lookup_range: "lookup_sheet",
};

function splitQualified(value: string): { sheet: string; address: string } | null {
  const m = /^\s*(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(value);
  if (!m) return null;
  const sheet = (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]).trim();
  const address = m[3].trim();
  // 多区域（A1:A5,C1:C5）或再次带表名的写法不动，交给工具报错
  if (!sheet || !address || address.includes("!")) return null;
  return { sheet, address };
}

/**
 * 拆开 "表名!地址"：sheet 参数为空时补上，与表名一致时去掉前缀；两者矛盾时保持原样（由工具报错，不猜）。
 * 只处理 schema 里确有对应 sheet 字段的工具。返回是否改动过。
 */
export function splitSheetQualifiedAddresses(args: Record<string, unknown>, schemaProps: Record<string, unknown>): boolean {
  let changed = false;
  for (const [addrKey, sheetKey] of Object.entries(ADDRESS_SHEET_KEYS)) {
    const value = args[addrKey];
    if (typeof value !== "string" || !(sheetKey in schemaProps)) continue;
    const parts = splitQualified(value);
    if (!parts) continue;
    const current = typeof args[sheetKey] === "string" ? String(args[sheetKey]).trim() : "";
    if (!current) args[sheetKey] = parts.sheet;
    else if (current !== parts.sheet) continue;
    args[addrKey] = parts.address;
    changed = true;
  }
  return changed;
}

/** 参数里引用到的所有 Sheet 名（sheet、source_sheet、sheets[]、inputs[].sheet、output.sheet…） */
export function referencedSheets(args: unknown): string[] {
  const found = new Set<string>();
  const visit = (value: unknown, key: string, depth: number) => {
    if (depth > 3 || value == null) return;
    if (typeof value === "string") {
      // new_sheet 是要新建的表名，不是引用
      if (/(^|_)sheets?$/.test(key) && !key.startsWith("new_") && value.trim()) found.add(value);
      return;
    }
    if (Array.isArray(value)) {
      for (const item of value) visit(item, key, depth + 1);
      return;
    }
    if (typeof value === "object") {
      for (const [k, v] of Object.entries(value as Record<string, unknown>)) visit(v, k, depth + 1);
    }
  };
  visit(args, "", 0);
  return Array.from(found);
}

function normalize(name: string): string {
  return name
    .replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
    .replace(/\s+/g, "")
    .toLowerCase();
}

function editDistance(a: string, b: string): number {
  const dp = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = dp[j];
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = tmp;
    }
  }
  return dp[b.length];
}

/** 最接近的现有表名：忽略大小写/空格/全半角后相同 > 互相包含 > 编辑距离足够小 */
export function closestSheetName(name: string, existing: string[]): string | null {
  const target = normalize(name);
  if (!target) return null;
  const exact = existing.find((s) => normalize(s) === target);
  if (exact) return exact;
  const contains = existing.filter((s) => normalize(s).includes(target) || target.includes(normalize(s)));
  if (contains.length === 1) return contains[0];
  let best: string | null = null;
  let bestScore = Infinity;
  for (const s of existing) {
    const d = editDistance(target, normalize(s));
    if (d < bestScore) {
      bestScore = d;
      best = s;
    }
  }
  return best && bestScore <= Math.max(1, Math.floor(target.length / 3)) ? best : null;
}

const MAX_LISTED = 30;

/** 参数里有不存在的表名时返回提示文字，否则返回空串 */
export function missingSheetHint(args: unknown, existing: string[]): string {
  const have = new Set(existing);
  const missing = referencedSheets(args).filter((s) => !have.has(s));
  if (!missing.length) return "";
  const parts = missing.map((s) => {
    const guess = closestSheetName(s, existing);
    return `工作表「${s}」不存在${guess ? `，你可能想用「${guess}」` : ""}`;
  });
  const listed = existing.slice(0, MAX_LISTED).map((s) => `「${s}」`).join("、");
  const more = existing.length > MAX_LISTED ? ` 等 ${existing.length} 张` : "";
  return `${parts.join("；")}。现有工作表：${listed}${more}。请用准确的表名重试（区分空格和全半角）。`;
}
