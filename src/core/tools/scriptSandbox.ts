/*
 * run_script 的执行沙箱：模型写的 JS 在一次性的 Web Worker 里跑。
 * - Worker 本身拿不到 Office.js、DOM、localStorage，也看不到 API Key；
 * - 编译完脚本后把全局对象上除纯计算内置对象以外的一切（fetch、XHR、WebSocket、importScripts、
 *   eval、Function 构造器、定时器……）全部移除，脚本无法联网或动态执行字符串；
 * - 超时或用户停止时直接 terminate，死循环也能收回。
 * 单元格内容可能带有提示注入，所以这里按「脚本不可信」设计，而不是只防手滑。
 */

export interface ScriptJob {
  code: string;
  /** 输入名 → 二维数组（已从工作簿读好） */
  inputs: Record<string, unknown[][]>;
  timeoutMs: number;
}

export interface ScriptOutcome {
  ok: boolean;
  value?: unknown;
  error?: string;
  logs: string[];
  ms: number;
}

export type ScriptRunner = (job: ScriptJob, signal?: AbortSignal) => Promise<ScriptOutcome>;

export const MAX_SCRIPT_CHARS = 20_000;

/** 在进入沙箱前就拒绝的写法：动态 import 是语法，无法靠移除全局对象拦住。 */
export function checkScriptSource(code: unknown): string | null {
  if (typeof code !== "string" || !code.trim()) return "code 必须是非空的 JS 函数体文本";
  if (code.length > MAX_SCRIPT_CHARS) return `脚本超过 ${MAX_SCRIPT_CHARS} 字符，请精简逻辑或拆成多步`;
  if (/\bimport\b/.test(code)) return "脚本里不能出现 import（沙箱不允许加载外部代码）；变量名或字符串里用到这个词时请换一个写法";
  if (!/\breturn\b/.test(code)) return "脚本必须用 return 返回结果（例如 return rows;）";
  return null;
}

// 纯计算需要的内置对象；其余全局名一律移除。
const KEEP = [
  "Object", "Array", "Number", "String", "Boolean", "Symbol", "BigInt", "Math", "JSON", "Date", "RegExp",
  "Map", "Set", "WeakMap", "WeakSet", "Promise", "Reflect", "Proxy", "Intl",
  "Error", "TypeError", "RangeError", "SyntaxError", "ReferenceError", "EvalError", "URIError", "AggregateError",
  "ArrayBuffer", "DataView", "Int8Array", "Uint8Array", "Uint8ClampedArray", "Int16Array", "Uint16Array",
  "Int32Array", "Uint32Array", "Float32Array", "Float64Array", "BigInt64Array", "BigUint64Array",
  "parseInt", "parseFloat", "isNaN", "isFinite", "encodeURIComponent", "decodeURIComponent", "encodeURI", "decodeURI",
  "NaN", "Infinity", "undefined", "globalThis", "TextEncoder", "TextDecoder", "structuredClone", "queueMicrotask",
];

// 这些必须确认已不可达，否则拒绝运行（失败即关闭）。
const MUST_BLOCK = [
  "fetch", "XMLHttpRequest", "WebSocket", "WebSocketStream", "EventSource", "importScripts", "Worker", "SharedWorker",
  "BroadcastChannel", "WebTransport", "FontFace", "fonts", "caches", "indexedDB", "Notification", "navigator",
  "eval", "Function", "setTimeout", "setInterval", "postMessage", "location",
];

/** Worker 源码。保持为不依赖打包器的纯 ES2019 文本。 */
export const WORKER_SOURCE = `"use strict";
// 包一层函数作用域：脚本里拿不到 send、helpers 实现等内部绑定
(function () {
const KEEP = new Set(${JSON.stringify(KEEP)});
const MUST_BLOCK = ${JSON.stringify(MUST_BLOCK)};
const send = self.postMessage.bind(self);
const MAX_LOGS = 50;

function chainOf(o) {
  const list = [];
  for (let p = o; p && p !== Object.prototype; p = Object.getPrototypeOf(p)) list.push(p);
  return list;
}

function lockdown() {
  const chain = chainOf(self);
  const names = new Set();
  for (const o of chain) for (const k of Object.getOwnPropertyNames(o)) names.add(k);
  for (const k of names) {
    if (KEEP.has(k)) continue;
    for (const o of chain.slice(1)) { try { delete o[k]; } catch (e) {} }
    try {
      Object.defineProperty(self, k, { value: undefined, writable: false, configurable: false, enumerable: false });
    } catch (e) {
      try { delete self[k]; } catch (e2) {}
    }
  }
  const fnProtos = [function () {}, async function () {}, function* () {}, async function* () {}].map(Object.getPrototypeOf);
  for (const p of fnProtos) {
    Object.defineProperty(p, "constructor", { value: undefined, writable: false, configurable: false });
  }
  for (const k of MUST_BLOCK) {
    for (const o of chain) {
      let v;
      try { v = o[k]; } catch (e) { v = undefined; }
      if (v !== undefined) throw new Error("沙箱未能移除 " + k);
    }
  }
  if ((function () {}).constructor !== undefined) throw new Error("沙箱未能移除 Function 构造器");
}

function toText(v) {
  if (typeof v === "string") return v;
  try { return JSON.stringify(v); } catch (e) { return String(v); }
}

const helpers = Object.freeze({
  /** 第一行当表头，其余行转成对象数组 */
  toObjects(grid) {
    if (!Array.isArray(grid) || !grid.length) return [];
    const headers = grid[0].map((h, i) => (String(h == null ? "" : h).trim() || "列" + (i + 1)));
    return grid.slice(1).map((row) => {
      const o = {};
      headers.forEach((h, i) => { o[h] = row[i] === undefined ? "" : row[i]; });
      return o;
    });
  },
  /** 对象数组 → 带表头的二维数组；headers 省略时取所有对象键的并集 */
  toGrid(objects, headers) {
    const list = Array.isArray(objects) ? objects : [];
    const cols = Array.isArray(headers) && headers.length ? headers.slice() : [];
    if (!cols.length) {
      const seen = new Set();
      for (const o of list) for (const k of Object.keys(o || {})) if (!seen.has(k)) { seen.add(k); cols.push(k); }
    }
    return [cols].concat(list.map((o) => cols.map((k) => (o && o[k] != null ? o[k] : ""))));
  },
  /** Excel 日期序列号 → "yyyy-mm-dd"；不是序列号时原样返回 */
  excelDate(serial) {
    if (typeof serial !== "number" || !isFinite(serial)) return serial;
    const d = new Date(Math.round((serial - 25569) * 86400000));
    return d.toISOString().slice(0, 10);
  },
  /** "2024-03-05" / "2024/3/5" / "2024年3月5日" → Excel 日期序列号；无法识别返回 null */
  toSerial(text) {
    if (typeof text === "number") return text;
    const m = /^(\\d{4})[-/.年](\\d{1,2})[-/.月](\\d{1,2})/.exec(String(text == null ? "" : text).trim());
    if (!m) return null;
    return Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569;
  },
  /** "¥1,200" "800元" "(35)" "12%" → 数字；无法识别返回 null */
  num(v) {
    if (typeof v === "number") return isFinite(v) ? v : null;
    let s = String(v == null ? "" : v).trim();
    if (!s) return null;
    let neg = false;
    if (/^\\(.*\\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    const pct = /%$/.test(s);
    s = s.replace(/[,，\\s¥￥$€£元%]/g, "");
    if (!/^[-+]?\\d*\\.?\\d+(e[-+]?\\d+)?$/i.test(s)) return null;
    let n = Number(s);
    if (pct) n /= 100;
    return neg ? -n : n;
  },
});

self.onmessage = function (ev) {
  self.onmessage = null;
  const job = ev.data || {};
  const logs = [];
  const log = function () {
    if (logs.length >= MAX_LOGS) return;
    logs.push(Array.prototype.map.call(arguments, toText).join(" ").slice(0, 500));
  };
  let fn;
  try {
    fn = new Function("inputs", "helpers", "log", '"use strict";\\n' + job.code);
  } catch (e) {
    send({ ok: false, error: "脚本语法错误: " + (e && e.message), logs: logs });
    return;
  }
  try {
    lockdown();
  } catch (e) {
    send({ ok: false, error: "脚本沙箱初始化失败，已拒绝运行: " + (e && e.message), logs: logs });
    return;
  }
  const fail = function (e) {
    const where = e && e.stack ? String(e.stack).split("\\n").slice(0, 3).join(" | ") : "";
    send({ ok: false, error: "脚本运行出错: " + (e && e.message ? e.message : toText(e)) + (where ? "（" + where + "）" : ""), logs: logs });
  };
  let result;
  try {
    result = fn(Object.freeze(job.inputs || {}), helpers, log);
  } catch (e) {
    fail(e);
    return;
  }
  Promise.resolve(result).then(function (value) {
    try {
      send({ ok: true, value: value === undefined ? null : value, logs: logs });
    } catch (e) {
      send({ ok: false, error: "返回值无法传回（只能是数字、文本、布尔、数组或普通对象，不能含函数）: " + (e && e.message), logs: logs });
    }
  }, fail);
};
})();
`;

function workerRunner(job: ScriptJob, signal?: AbortSignal): Promise<ScriptOutcome> {
  return new Promise((resolve) => {
    const started = Date.now();
    let url = "";
    let worker: Worker;
    try {
      url = URL.createObjectURL(new Blob([WORKER_SOURCE], { type: "text/javascript" }));
      worker = new Worker(url);
    } catch (e: any) {
      if (url) URL.revokeObjectURL(url);
      resolve({ ok: false, error: `当前环境无法创建脚本沙箱（Web Worker）：${e?.message || e}。请改用其他工具完成。`, logs: [], ms: 0 });
      return;
    }
    let done = false;
    const finish = (outcome: Omit<ScriptOutcome, "ms">) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
      worker.terminate();
      URL.revokeObjectURL(url);
      resolve({ ...outcome, ms: Date.now() - started });
    };
    const timer = setTimeout(() => finish({
      ok: false,
      error: `脚本运行超过 ${Math.round(job.timeoutMs / 1000)} 秒被终止。请检查是否有死循环，或减少输入数据量、降低算法复杂度（避免嵌套遍历全表）。`,
      logs: [],
    }), job.timeoutMs);
    const onAbort = () => finish({ ok: false, error: "用户已停止任务，脚本已终止。", logs: [] });
    if (signal?.aborted) return onAbort();
    signal?.addEventListener("abort", onAbort);
    worker.onmessage = (ev: MessageEvent) => {
      const data = (ev.data || {}) as Partial<ScriptOutcome>;
      finish({
        ok: data.ok === true,
        value: data.value,
        error: data.ok === true ? undefined : String(data.error || "脚本没有返回结果"),
        logs: Array.isArray(data.logs) ? data.logs.map(String) : [],
      });
    };
    worker.onerror = (ev: ErrorEvent) => {
      ev.preventDefault();
      finish({ ok: false, error: `脚本运行出错: ${ev.message || "未知错误"}`, logs: [] });
    };
    try {
      worker.postMessage({ code: job.code, inputs: job.inputs });
    } catch (e: any) {
      finish({ ok: false, error: `输入数据无法传入沙箱: ${e?.message || e}`, logs: [] });
    }
  });
}

let runner: ScriptRunner = workerRunner;

export function runScript(job: ScriptJob, signal?: AbortSignal): Promise<ScriptOutcome> {
  return runner(job, signal);
}

/** 测试用：Node 里没有 Web Worker，替换成可控的执行器 */
export function setScriptRunner(next: ScriptRunner | null): void {
  runner = next ?? workerRunner;
}
