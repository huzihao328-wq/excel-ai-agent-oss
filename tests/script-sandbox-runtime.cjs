const assert = require('node:assert/strict');
const { test } = require('node:test');
const vm = require('node:vm');
const { sourceLoader } = require('./load-source.cjs');

const { WORKER_SOURCE, checkScriptSource } = sourceLoader()('src/core/tools/scriptSandbox.ts');

// 在独立 V8 上下文里模拟 Worker 全局：带上真实 Worker 里会有的联网/动态执行入口，确认沙箱把它们都拿掉了。
function runInSandbox(code, inputs = {}) {
  return new Promise((resolve, reject) => {
    const leaked = [];
    const sandbox = vm.createContext({});
    // 假的联网入口都在沙箱自己的 realm 里创建（真实 Worker 也是单 realm），通过闭包把「泄漏」报回宿主。
    vm.runInContext(`(function (leak, reply) {
      Object.setPrototypeOf(globalThis, { fetch() { leak('proto.fetch'); } });
      globalThis.fetch = function () { leak('fetch'); };
      globalThis.XMLHttpRequest = function () { leak('xhr'); };
      globalThis.importScripts = function () { leak('importScripts'); };
      globalThis.setTimeout = function () { leak('setTimeout'); };
      globalThis.navigator = { sendBeacon() { leak('beacon'); } };
      globalThis.postMessage = function (message) { reply(message); };
      globalThis.self = globalThis;
    })`, sandbox)(name => leaked.push(name), message => resolve({ message: JSON.parse(JSON.stringify(message)), leaked }));
    vm.runInContext(WORKER_SOURCE, sandbox);
    const timer = setTimeout(() => reject(new Error('sandbox never answered')), 2000);
    sandbox.onmessage({ data: { code, inputs } });
    setTimeout(() => clearTimeout(timer), 0);
  });
}

test('sandbox: pure computation over inputs with helpers returns structured results and logs', async () => {
  const code = `
    const rows = helpers.toObjects(inputs.t);
    log('rows', rows.length);
    const total = rows.reduce((s, r) => s + (helpers.num(r['金额']) || 0), 0);
    return { total, date: helpers.excelDate(45000), serial: helpers.toSerial('2023年3月15日'), grid: helpers.toGrid(rows, ['名称']) };`;
  const { message, leaked } = await runInSandbox(code, { t: [['名称', '金额'], ['甲', '¥1,200'], ['乙', '(35)'], ['丙', '12%']] });
  assert.equal(message.ok, true, message.error);
  assert.equal(message.value.total, 1200 - 35 + 0.12);
  assert.equal(message.value.date, '2023-03-15');
  assert.equal(message.value.serial, 45000);
  assert.deepEqual(message.value.grid, [['名称'], ['甲'], ['乙'], ['丙']]);
  assert.deepEqual(message.logs, ['rows 3']);
  assert.deepEqual(leaked, []);
});

for (const [label, code] of [
  ['global fetch', 'return fetch("https://x")'],
  ['globalThis.fetch', 'return globalThis.fetch("https://x")'],
  ['prototype fetch', 'return Object.getPrototypeOf(globalThis).fetch("https://x")'],
  ['XMLHttpRequest', 'return new XMLHttpRequest()'],
  ['importScripts', 'return importScripts("https://x")'],
  ['navigator.sendBeacon', 'return navigator.sendBeacon("https://x")'],
  ['timer string eval', 'return setTimeout("fetch(1)")'],
  ['eval', 'return eval("fetch")("https://x")'],
  ['Function constructor', 'return (function(){}).constructor("return fetch")()("https://x")'],
  ['async Function constructor', 'return (async function(){}).constructor("return fetch")()'],
  ['arrow constructor via prototype', 'return Object.getPrototypeOf(() => 0).constructor("return this")().fetch("x")'],
  ['re-posting results', 'return postMessage({ ok: true, value: "forged" })'],
]) {
  test(`sandbox: ${label} is unreachable`, async () => {
    const { message, leaked } = await runInSandbox(code);
    assert.equal(message.ok, false, `expected failure, got ${JSON.stringify(message)}`);
    assert.match(message.error, /脚本运行出错/);
    assert.deepEqual(leaked, []);
  });
}

test('sandbox: syntax errors and non-cloneable results are reported, not thrown', async () => {
  const syntax = await runInSandbox('return (');
  assert.equal(syntax.message.ok, false);
  assert.match(syntax.message.error, /语法错误/);
  const asyncResult = await runInSandbox('return Promise.resolve([[1, 2]])');
  assert.deepEqual(asyncResult.message.value, [[1, 2]]);
});

test('source check rejects dynamic import, missing return and oversized scripts before running', () => {
  assert.match(checkScriptSource('return import("https://x")'), /import/);
  assert.match(checkScriptSource('const x = 1;'), /return/);
  assert.match(checkScriptSource('return 1;' + ' '.repeat(20001)), /20000/);
  assert.match(checkScriptSource(''), /非空/);
  assert.equal(checkScriptSource('return inputs.t.length'), null);
});

test('sandbox: internal bindings (send, lockdown, allowlists) are not visible to the script', async () => {
  const { message } = await runInSandbox('return [typeof send, typeof KEEP, typeof lockdown, typeof helpers.toObjects]');
  assert.deepEqual(message.value, ['undefined', 'undefined', 'undefined', 'function']);
});
