/**
 * 回复里的结果位置变成可点击的定位：`汇总!A1:D12`、`'Q1 销售'!B2`、「汇总」。
 * 只认当前工作簿里真实存在的 Sheet：已删除或改名的表不出链接，免得点了没反应。
 */

export interface RangeRef {
  start: number;
  end: number;
  sheet: string;
  address?: string;
}

// 单元格或区域（A1、$A$1:D12）、整列（A:C）、整行（2:5）
const ADDRESS = String.raw`\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d{0,6}:\$?[1-9]\d{0,6}`;
const ADDRESS_END = String.raw`(?![A-Za-z0-9_])`;
const WORD_CHAR = /[A-Za-z0-9_]/;

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

let cached: { key: string; pattern: RegExp; byLower: Map<string, string> } | null = null;

function matcherFor(sheetNames: string[]) {
  const key = sheetNames.join("\u0000");
  if (cached?.key === key) return cached;
  const names = sheetNames.filter(Boolean).sort((a, b) => b.length - a.length);
  const bare = names.map(escapeRegExp).join("|");
  const quoted = names.map((n) => escapeRegExp(n.replace(/'/g, "''"))).join("|");
  const addr = `(${ADDRESS})${ADDRESS_END}`;
  // 三种写法依次是：'带空格的表'!A1、「表」（可跟地址）、表!A1
  const pattern = new RegExp(
    `'(${quoted})'!${addr}|「(${bare})」(?:(?:!|\\s)?${addr})?|(${bare})!${addr}`,
    "gi"
  );
  const byLower = new Map(names.map((n) => [n.toLowerCase(), n]));
  cached = { key, pattern, byLower };
  return cached;
}

export function findRangeRefs(text: string, sheetNames: string[]): RangeRef[] {
  if (!text || sheetNames.length === 0) return [];
  const { pattern, byLower } = matcherFor(sheetNames);
  const refs: RangeRef[] = [];
  pattern.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = pattern.exec(text))) {
    const quotedName = m[1]?.replace(/''/g, "'");
    const name = quotedName ?? m[3] ?? m[5];
    const address = m[2] ?? m[4] ?? m[6];
    // 英文表名前面紧挨着字母数字时是别的词的一部分（MySheet1!A1 不是 Sheet1）
    if (m[5] && WORD_CHAR.test(m[5][0]) && m.index > 0 && WORD_CHAR.test(text[m.index - 1])) {
      pattern.lastIndex = m.index + 1;
      continue;
    }
    const sheet = byLower.get(name.toLowerCase());
    if (!sheet) continue;
    refs.push({ start: m.index, end: m.index + m[0].length, sheet, address: address?.toUpperCase() });
  }
  return refs;
}

/** 把 root 里文本中的地址替换成定位按钮；链接、按钮和代码块里的内容不动 */
export function linkRangeRefs(root: Node, sheetNames: string[]): void {
  if (sheetNames.length === 0) return;
  const doc = root.ownerDocument || (root as Document);
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const targets: Text[] = [];
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!node.parentElement?.closest("a, pre, button")) targets.push(node);
  }
  for (const node of targets) {
    const text = node.data;
    const refs = findRangeRefs(text, sheetNames);
    if (refs.length === 0) continue;
    const frag = doc.createDocumentFragment();
    let pos = 0;
    for (const ref of refs) {
      if (ref.start > pos) frag.append(text.slice(pos, ref.start));
      const link = doc.createElement("button");
      link.type = "button";
      link.className = "range-link";
      link.dataset.sheet = ref.sheet;
      if (ref.address) link.dataset.address = ref.address;
      link.title = ref.address ? `在 Excel 中选中 ${ref.sheet}!${ref.address}` : `切换到工作表「${ref.sheet}」`;
      link.textContent = text.slice(ref.start, ref.end);
      frag.append(link);
      pos = ref.end;
    }
    if (pos < text.length) frag.append(text.slice(pos));
    node.replaceWith(frag);
  }
}
