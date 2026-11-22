import { createContext, MouseEvent, useContext, useMemo } from "react";
import { Marked } from "marked";
import DOMPurify from "dompurify";
import { linkRangeRefs } from "../rangeLinks";
import { copyText } from "../clipboard";

/*
 * 模型的回复会受单元格、附件里的文字影响，这里按不可信内容渲染。
 * 远程图片（以及 style、video 等）一渲染就自动发请求，URL 里可以夹带表格数据，不需要用户点任何东西；
 * 所以只放行 Markdown 用得到的标签，图片降级成要手点才打开的链接。taskpane.html 里的 CSP 是第二道闸。
 */
const md = new Marked({
  breaks: true,
  walkTokens(token) {
    if (token.type !== "image") return;
    const label = token.text || "图片";
    Object.assign(token, { type: "link", tokens: [{ type: "text", raw: label, text: label }] });
  },
});

const ALLOWED_TAGS = [
  "a", "b", "blockquote", "br", "code", "del", "div", "em", "h1", "h2", "h3", "h4", "h5", "h6", "hr", "i", "input",
  "kbd", "li", "mark", "ol", "p", "pre", "s", "span", "strong", "sub", "sup", "table", "tbody", "td", "tfoot", "th",
  "thead", "tr", "u", "ul",
];
const ALLOWED_ATTR = ["href", "title", "align", "start", "class", "colspan", "rowspan", "type", "checked", "disabled"];

DOMPurify.addHook("afterSanitizeAttributes", (node) => {
  // class 只留代码块的语言标记，回复不能借用界面自己的样式类冒充按钮
  const cls = node.getAttribute("class");
  if (cls && !/^language-[\w+#.-]+$/.test(cls)) node.removeAttribute("class");
  // input 只为任务列表的勾选框保留，且不可操作
  if (node.tagName === "INPUT") {
    if (node.getAttribute("type") !== "checkbox") node.remove();
    else node.setAttribute("disabled", "");
    return;
  }
  // 任务窗格里点普通链接会把整个窗格导航走，插件界面随之消失；统一改为在系统浏览器里打开。
  if (node.tagName !== "A") return;
  const href = node.getAttribute("href") || "";
  if (!href || href.startsWith("#")) return;
  node.setAttribute("target", "_blank");
  node.setAttribute("rel", "noopener noreferrer");
});

/** 回复里的地址链接需要知道有哪些 Sheet，点了之后怎么在 Excel 里定位 */
export interface WorkbookLinks {
  sheetNames: string[];
  locate: (sheet: string, address?: string) => void;
}
export const WorkbookLinksContext = createContext<WorkbookLinks | null>(null);

function renderMarkdown(text: string, sheetNames: string[]): string {
  const raw = md.parse(text, { async: false }) as string;
  const fragment = DOMPurify.sanitize(raw, {
    ALLOWED_TAGS,
    ALLOWED_ATTR,
    ALLOW_DATA_ATTR: false,
    ALLOW_ARIA_ATTR: false,
    RETURN_DOM_FRAGMENT: true,
  });
  linkRangeRefs(fragment, sheetNames);
  // 代码块（多半是公式）右上角放一个复制按钮，不用手动框选
  fragment.querySelectorAll("pre").forEach((pre) => {
    const wrap = document.createElement("div");
    wrap.className = "code-block";
    const copy = document.createElement("button");
    copy.type = "button";
    copy.className = "code-copy";
    copy.title = "复制代码";
    copy.textContent = "复制";
    pre.replaceWith(wrap);
    wrap.append(copy, pre);
  });
  const container = document.createElement("div");
  container.append(fragment);
  return container.innerHTML;
}

const NO_SHEETS: string[] = [];

export default function Markdown({ text }: { text: string }) {
  const links = useContext(WorkbookLinksContext);
  // 不含地址写法的回复不依赖 Sheet 列表：新建工作表时长对话里的其他回复不用重新解析
  const sheetNames = /[!「]/.test(text) ? links?.sheetNames ?? NO_SHEETS : NO_SHEETS;
  const html = useMemo(() => renderMarkdown(text, sheetNames), [text, sheetNames]);

  const onClick = async (e: MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const link = target.closest<HTMLElement>(".range-link");
    if (link?.dataset.sheet) {
      links?.locate(link.dataset.sheet, link.dataset.address);
      return;
    }
    const copy = target.closest<HTMLButtonElement>(".code-copy");
    const code = copy?.parentElement?.querySelector("pre")?.textContent;
    if (!copy || code == null) return;
    if (!(await copyText(code.replace(/\n$/, "")))) return;
    copy.textContent = "已复制";
    copy.classList.add("copied");
    window.setTimeout(() => {
      copy.textContent = "复制";
      copy.classList.remove("copied");
    }, 1500);
  };

  return <div className="markdown" onClick={onClick} dangerouslySetInnerHTML={{ __html: html }} />;
}
