const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const { sourceLoader } = require('./load-source.cjs');
const { DEFAULT_SETTINGS } = sourceLoader()('src/core/settings.ts');
const root = path.resolve(__dirname, '..');
const frame = value => `data: ${JSON.stringify(value)}\n\n`;
const reply = text => frame({ choices: [{ delta: { content: text }, finish_reason: 'stop' }] }) + 'data: [DONE]\n\n';
const choice = id => frame({ choices: [{ delta: { tool_calls: [{ index: 0, id, type: 'function', function: {
  name: 'ask_user_choice', arguments: JSON.stringify({ title: '选择处理方向', options: [{ id: 'summary', label: '汇总数据' }, { id: 'clean', label: '清洗数据' }] }),
} }] }, finish_reason: 'tool_calls' }] }) + 'data: [DONE]\n\n';
const localSettings = { ...DEFAULT_SETTINGS, baseUrl: 'http://localhost:1234/v1', model: 'test-model', apiKey: '', autoCheckpoint: false, useMemory: false, autoLearnSkills: false };
const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };
const json = (route, data) => route.fulfill({ contentType: 'application/json', headers: { 'Access-Control-Allow-Origin': '*' }, body: JSON.stringify(data) });
const stream = (route, body) => route.fulfill({ contentType: 'text/event-stream', headers: { 'Access-Control-Allow-Origin': '*' }, body });

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  const errors = [];
  async function pageFor(settings = DEFAULT_SETTINGS, intercept = async () => false) {
    const context = await browser.newContext({ viewport: { width: 380, height: 850 } });
    const page = await context.newPage();
    page.setDefaultTimeout(8000);
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(settings => localStorage.setItem('excel-ai-agent-settings', JSON.stringify(settings)), settings);
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (await intercept(route, url)) return;
      if (url.origin !== 'https://excel-agent.test') {
        if (url.pathname === '/v1/models') return json(route, { data: [{ id: 'test-model', context_length: 32000 }] });
        return json(route, {});
      }
      if (url.pathname.endsWith('/office-js/office.js')) {
        return route.fulfill({ contentType: 'application/javascript', body: 'window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:"selection"},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:(_event,fn)=>{window.selectionChanged=fn},removeHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error("Office mocked"))};' });
      }
      if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
      const asset = path.resolve(root, 'dist', '.' + url.pathname);
      if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
      return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
    });
    await page.goto('https://excel-agent.test/taskpane.html');
    await page.locator('.welcome-panel').waitFor();
    return { page, context };
  }
  try {
    // The real startup entrypoint must render while a profile PUT is stalled.
    const backup = deferred();
    const startup = await pageFor(DEFAULT_SETTINGS, async (route, url) => {
      if (url.pathname !== '/api/user-data' || route.request().method() !== 'PUT') return false;
      await backup.promise;
      await json(route, {});
      return true;
    });
    backup.resolve();
    await startup.context.close();
    console.log('PASS: startup renders without waiting for backup PUT');

    const oldModels = deferred();
    const requestedModels = deferred();
    const modelResponseDone = deferred();
    const settingsTest = await pageFor(DEFAULT_SETTINGS, async (route, url) => {
      if (url.port !== '1234' || url.pathname !== '/v1/models') return false;
      requestedModels.resolve();
      await oldModels.promise;
      await json(route, { data: [{ id: 'obsolete-model', context_length: 999000 }] });
      modelResponseDone.resolve();
      return true;
    });
    const p = settingsTest.page;
    await p.getByTitle('模型设置', { exact: true }).click();
    await p.getByRole('button', { name: 'LM Studio 本地', exact: true }).click();
    await p.getByRole('button', { name: '获取模型', exact: true }).click();
    await requestedModels.promise;
    await p.getByRole('button', { name: 'Ollama 本地', exact: true }).click();
    oldModels.resolve();
    await modelResponseDone.promise;
    await p.getByPlaceholder('deepseek-chat').fill('different-model');
    assert.equal(await p.locator('.model-list').count(), 0, 'old endpoint models reappeared');
    const windowInput = p.getByPlaceholder('接入后自动探测，也可填 128k');
    await windowInput.fill('128k');
    await p.getByPlaceholder('deepseek-chat').fill('next-model');
    assert.equal(await windowInput.inputValue(), '', 'previous model context window leaked');
    await p.getByPlaceholder('deepseek-chat').fill('');
    assert.equal(await p.getByRole('button', { name: '保存', exact: true }).isDisabled(), true);
    await p.getByPlaceholder('deepseek-chat').fill('saved-model');
    await p.evaluate(() => {
      const original = Storage.prototype.setItem;
      Storage.prototype.setItem = function (key, value) {
        if (key === 'excel-ai-agent-settings') throw new DOMException('full', 'QuotaExceededError');
        return original.call(this, key, value);
      };
    });
    await p.getByRole('button', { name: '保存', exact: true }).click();
    await p.getByText('设置未保存，本机存储不可用或空间不足。请释放空间后重试。').waitFor();
    assert.equal(await p.locator('.header-model > span').first().innerText(), DEFAULT_SETTINGS.model);
    for (const width of [300, 380, 600]) {
      await p.setViewportSize({ width, height: 650 });
      assert.equal(await p.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    }
    await settingsTest.context.close();
    console.log('PASS: stale model requests, context reset, empty model validation, failed saves, narrow settings');

    let toolProbeRequests = 0;
    const capabilities = await pageFor(localSettings, async (route, url) => {
      if (url.pathname !== '/v1/chat/completions') return false;
      const body = route.request().postDataJSON();
      if (!body.tools) await json(route, { choices: [{ message: { content: 'pong' }, finish_reason: 'stop' }] });
      else {
        toolProbeRequests++;
        assert.equal(body.tools[0].function.name, 'connection_probe');
        if (body.messages.at(-1).role === 'tool') {
          await stream(route, reply(JSON.parse(body.messages.at(-1).content).receipt));
        } else {
          await stream(route, frame({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'probe', type: 'function', function: { name: 'connection_probe', arguments: '{"value":7}' } }] }, finish_reason: 'tool_calls' }] }) + 'data: [DONE]\n\n');
        }
      }
      return true;
    });
    const cp = capabilities.page;
    await cp.getByTitle('模型设置', { exact: true }).click();
    const manualWindow = cp.getByPlaceholder('接入后自动探测，也可填 128k');
    await manualWindow.fill('16k');
    await cp.getByRole('button', { name: '测试连接', exact: true }).click();
    await cp.locator('.test-result').filter({ hasText: '原生工具调用、参数和结果回填测试通过' }).waitFor();
    assert.equal(toolProbeRequests, 2);
    assert.equal(await manualWindow.inputValue(), '16K', 'connection test replaced the manual context limit');
    assert.match(await cp.locator('.test-result').innerText(), /32K（接口返回）/);
    await manualWindow.fill('');
    await cp.getByRole('button', { name: '探测', exact: true }).click();
    await cp.locator('.test-result').filter({ hasText: '已用于自动预算' }).waitFor();
    await cp.getByRole('button', { name: '保存', exact: true }).click();
    assert.equal(await cp.evaluate(() => JSON.parse(localStorage.getItem('excel-ai-agent-settings')).contextWindowTokens), 0, 'auto detection was persisted as a manual override');
    await capabilities.context.close();
    console.log('PASS: capability round trip, context provenance and manual override preservation');

    let requests = [];
    const queuedRequest = deferred();
    const finishQueue = deferred();
    const chat = await pageFor(localSettings, async (route, url) => {
      if (url.pathname !== '/v1/chat/completions') return false;
      requests.push(route.request().postDataJSON());
      if (requests.length === 1) await stream(route, choice('queue-choice'));
      else if (requests.length === 2) await stream(route, reply('第一项已完成。'));
      else {
        queuedRequest.resolve();
        await finishQueue.promise;
        await stream(route, reply('排队图片任务完成。'));
      }
      return true;
    });
    const input = chat.page.getByRole('textbox', { name: '任务指令' });
    await input.fill('第一项任务');
    await chat.page.getByTitle('发送', { exact: true }).click();
    await chat.page.locator('.choice-card:not(.resolved)').waitFor();
    await chat.page.locator('.run-status').getByText('等待你的确认或选择').waitFor();
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=', 'base64');
    await chat.page.locator('.composer input[type=file]').setInputFiles({ name: 'queued.png', mimeType: 'image/png', buffer: png });
    await chat.page.locator('.image-chip').waitFor();
    await input.fill('处理排队图片');
    await chat.page.getByRole('button', { name: '排队发送', exact: true }).click();
    await chat.page.locator('.queue-chip').waitFor();
    await chat.page.locator('.image-chip').waitFor({ state: 'detached' });
    assert.equal(await chat.page.locator('.image-chip').count(), 0);
    await input.fill('这是仍在编辑的新草稿');
    await chat.page.getByRole('button', { name: '汇总数据', exact: true }).click();
    await queuedRequest.promise;
    assert.equal(await input.inputValue(), '这是仍在编辑的新草稿');
    const user = requests[2].messages.filter(m => m.role === 'user').at(-1);
    assert.equal(user.content.filter(part => part.type === 'image_url').length, 1);
    assert.ok(user.content[0].text.includes('处理排队图片'));
    finishQueue.resolve();
    await chat.page.getByText('排队图片任务完成。', { exact: true }).waitFor();
    await chat.context.close();
    console.log('PASS: waiting status, queued image snapshot, preservation of new draft');

    const failResponse = deferred();
    let failedRequests = 0;
    const failed = await pageFor(localSettings, async (route, url) => {
      if (url.pathname !== '/v1/chat/completions') return false;
      failedRequests++;
      await failResponse.promise;
      await stream(route, failedRequests === 1 ? frame({ error: { message: 'rejected task' } }) + 'data: [DONE]\n\n' : reply('后续任务成功。'));
      return true;
    });
    const failedInput = failed.page.getByRole('textbox', { name: '任务指令' });
    await failedInput.fill('将失败的任务');
    await failed.page.getByTitle('发送', { exact: true }).click();
    await failedInput.fill('依赖上一项的任务');
    await failed.page.getByRole('button', { name: '排队发送', exact: true }).click();
    failResponse.resolve();
    await failed.page.locator('.queue-chip').getByText('上一任务未完成，已暂停排队：', { exact: false }).waitFor();
    assert.equal(failedRequests, 1);
    await failed.page.getByTitle('取回指令').click();
    assert.equal(await failedInput.inputValue(), '依赖上一项的任务');
    await failed.page.getByTitle('发送', { exact: true }).click();
    await failed.page.getByText('后续任务成功。', { exact: true }).waitFor();
    assert.equal(await failed.page.getByRole('button', { name: '重试这条指令' }).count(), 0, 'old retry button could rerun a newer task');
    await failed.context.close();
    console.log('PASS: failed tasks pause queued work and allow recovering the prompt');

    const quota = await pageFor(localSettings, async (route, url) => {
      if (url.pathname !== '/v1/chat/completions') return false;
      await stream(route, reply('存储不可用时仍能完成回复。'));
      return true;
    });
    await quota.page.evaluate(() => { Storage.prototype.setItem = () => { throw new DOMException('full', 'QuotaExceededError'); }; });
    await quota.page.getByRole('textbox', { name: '任务指令' }).fill('检查存储满的情况');
    await quota.page.getByTitle('发送', { exact: true }).click();
    await quota.page.getByText('存储不可用时仍能完成回复。', { exact: true }).waitFor();
    await quota.page.locator('.storage-warning').waitFor();
    await quota.page.getByTitle('发送', { exact: true }).waitFor();
    await quota.context.close();
    console.log('PASS: full storage shows a save warning without stranding the task');

    const workbook = await pageFor(localSettings);
    await workbook.page.evaluate(() => {
      window.selectionReads = [];
      Excel.run = () => new Promise(resolve => window.selectionReads.push(resolve));
      window.selectionChanged();
      window.selectionChanged();
      window.selectionReads[1]({ sheetNames: ['最新工作表'], activeSheet: '最新工作表', selectionAddress: '最新工作表!B2' });
    });
    await workbook.page.locator('.ctx-chip.sheet').getByText('最新工作表', { exact: false }).waitFor();
    await workbook.page.evaluate(() => window.selectionReads[0]({ sheetNames: ['过期工作表'], activeSheet: '过期工作表', selectionAddress: '过期工作表!A1' }));
    assert.ok((await workbook.page.locator('.ctx-chip.sheet').innerText()).includes('最新工作表'));
    await workbook.page.evaluate(() => {
      Office.FileType = { Compressed: 'compressed' };
      Office.AsyncResultStatus = { Succeeded: 'succeeded' };
      Office.context.document.getFileAsync = (_type, _options, cb) => {
        window.finishExport = () => cb({ status: 'failed', error: { message: 'controlled export failure' } });
      };
    });
    await workbook.page.getByTitle('更多功能', { exact: true }).click();
    await workbook.page.getByRole('button', { name: '导出工作簿', exact: false }).click();
    await workbook.page.getByText('正在处理工作簿，请稍候…', { exact: true }).waitFor();
    const workbookInput = workbook.page.getByRole('textbox', { name: '任务指令' });
    await workbookInput.fill('导出期间保留的草稿');
    assert.equal(await workbook.page.getByTitle('发送', { exact: true }).isDisabled(), true);
    await workbookInput.press('Enter');
    assert.equal(await workbookInput.inputValue(), '导出期间保留的草稿');
    await workbook.page.evaluate(() => window.finishExport());
    await workbook.page.getByText('导出失败: controlled export failure', { exact: true }).waitFor();
    assert.equal(await workbook.page.getByTitle('发送', { exact: true }).isEnabled(), true);
    await workbook.context.close();
    console.log('PASS: selection responses cannot regress context; export locks workbook actions and releases after failure');

    let batchCalls = 0;
    const batch = await pageFor(localSettings, async (route, url) => {
      if (url.pathname !== '/v1/chat/completions') return false;
      batchCalls++;
      await stream(route, batchCalls === 1 ? choice('batch-choice') : reply('批量流程完成。'));
      return true;
    });
    await batch.page.getByTitle('更多功能', { exact: true }).click();
    await batch.page.getByRole('button', { name: '批量任务', exact: true }).click();
    await batch.page.locator('.batch-panel input[type=file]').setInputFiles({ name: 'sample.csv', mimeType: 'text/csv', buffer: Buffer.from('id,value\n001,12') });
    await batch.page.locator('.batch-file').waitFor();
    await batch.page.locator('.batch-instruction textarea').fill('汇总文件');
    await batch.page.getByRole('button', { name: '开始批量处理', exact: false }).click();
    await batch.page.locator('.choice-card:not(.resolved)').waitFor();
    // 面板收起有退场动画，等待其真正隐藏
    await batch.page.locator('.batch-panel').waitFor({ state: 'hidden' });
    assert.equal(await batch.page.getByTitle('新对话', { exact: true }).isDisabled(), true);
    await batch.page.getByRole('button', { name: '查看进度', exact: true }).click();
    await batch.page.getByTitle('收起，任务继续执行').click();
    await batch.page.getByRole('button', { name: '汇总数据', exact: true }).click();
    await batch.page.getByText('批量任务执行结束：1 个完成。各文件结果和执行详情见批量面板。', { exact: true }).waitFor();
    await batch.page.getByTitle('更多功能', { exact: true }).click();
    await batch.page.getByRole('button', { name: '批量任务', exact: true }).click();
    await batch.page.locator('.batch-run-item.done').waitFor();
    assert.equal(batchCalls, 2);
    await batch.page.setViewportSize({ width: 300, height: 650 });
    assert.equal(await batch.page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    await batch.page.screenshot({ path: path.join(__dirname, 'review-batch-narrow.png'), animations: 'disabled' });
    await batch.page.setViewportSize({ width: 300, height: 320 });
    await batch.page.locator('.batch-run-item.done').scrollIntoViewIfNeeded();
    const batchRect = await batch.page.locator('.batch-run-item.done').boundingBox();
    assert.ok(batchRect.y >= 0 && batchRect.y + batchRect.height <= 320, 'batch results cannot be reached in a short pane');
    await batch.context.close();
    console.log('PASS: batch confirmations remain reachable, minimizing keeps work alive, results survive reopening');
    assert.deepEqual(errors, []);
    console.log('PASS: all stability UI scenarios; model and Office mocked');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
