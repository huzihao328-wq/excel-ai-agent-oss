const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const { sourceLoader } = require('./load-source.cjs');
const { DEFAULT_SETTINGS } = sourceLoader()('src/core/settings.ts');
const root = path.resolve(__dirname, '..');
const frame = data => `data: ${JSON.stringify(data)}\n\n`;
const completion = text => [...text].map(content => frame({ choices: [{ delta: { content } }] })).join('') + frame({ choices: [{ delta: {}, finish_reason: 'stop' }] }) + 'data: [DONE]\n\n';
const choice = id => frame({ choices: [{ delta: { tool_calls: [{ index: 0, id, type: 'function', function: {
  name: 'ask_user_choice', arguments: JSON.stringify({ title: '选择处理方向', options: [{ id: 'summary', label: '汇总数据' }, { id: 'clean', label: '清洗数据' }] }),
} }] }, finish_reason: 'tool_calls' }] }) + 'data: [DONE]\n\n';

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 380, height: 850 } });
    const errors = [];
    const requests = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(settings => {
      if (!localStorage.getItem('excel-ai-agent-settings')) localStorage.setItem('excel-ai-agent-settings', JSON.stringify(settings));
    }, { ...DEFAULT_SETTINGS, baseUrl: 'http://localhost:11434/v1', model: 'ui-test', apiKey: '', useMemory: false, autoLearnSkills: false, autoCheckpoint: false });
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname === 'localhost' && url.pathname === '/v1/chat/completions') {
        const body = route.request().postDataJSON();
        requests.push(body);
        let response;
        switch (requests.length) {
          case 1: response = choice('choice-1'); break;
          case 2:
            assert.equal(body.messages.at(-1).tool_call_id, 'choice-1');
            assert.equal(JSON.parse(body.messages.at(-1).content).ok, true);
            response = completion('已按选择完成，金额111元。'); break;
          case 3: response = completion('排队任务已完成。'); break;
          case 4: response = choice('choice-stop'); break;
          case 5: {
            const stopped = body.messages.find(message => message.tool_call_id === 'choice-stop');
            assert.ok(stopped, 'Missing result for stopped tool');
            assert.equal(JSON.parse(stopped.content).ok, false);
            response = completion('会话已恢复。'); break;
          }
          case 6: response = frame({ error: { message: 'test gateway rejected' } }) + 'data: [DONE]\n\n'; break;
          case 7: response = completion('重试已完成。'); break;
          default: throw new Error(`Unexpected model request ${requests.length}`);
        }
        return route.fulfill({ contentType: 'text/event-stream', headers: { 'Access-Control-Allow-Origin': '*' }, body: response });
      }
      if (url.origin !== 'https://excel-agent.test') return route.abort();
      if (url.pathname.endsWith('/office-js/office.js')) {
        return route.fulfill({ contentType: 'application/javascript', body: 'window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:"selection"},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error("Office mocked"))};' });
      }
      if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
      const asset = path.resolve(root, 'dist', '.' + url.pathname);
      if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
      return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
    });
    await page.goto('https://excel-agent.test/taskpane.html');
    const input = page.getByRole('textbox', { name: '任务指令' });
    const send = async text => { await input.fill(text); await page.getByTitle('发送', { exact: true }).click(); };
    await send('处理这份数据');
    await page.locator('.choice-card:not(.resolved)').waitFor();
    await input.fill('接着整理结果');
    await page.getByRole('button', { name: '排队发送', exact: true }).click();
    await page.locator('.queue-chip').waitFor();
    assert.equal(requests.length, 1);
    await page.getByRole('button', { name: '汇总数据', exact: true }).click();
    await page.getByText('排队任务已完成。', { exact: true }).waitFor();
    await page.getByText('已按选择完成，金额111元。', { exact: true }).waitFor();
    assert.equal(requests.length, 3);
    await send('中止测试');
    await page.locator('.choice-card:not(.resolved)').waitFor();
    await page.getByTitle('停止', { exact: true }).click();
    await page.getByTitle('发送', { exact: true }).waitFor();
    await send('恢复会话');
    await page.getByText('会话已恢复。', { exact: true }).waitFor();
    await send('失败后重试');
    await page.getByRole('button', { name: '重试这条指令' }).click();
    await page.getByText('重试已完成。', { exact: true }).waitFor();
    assert.equal(requests.length, 7);
    assert.deepEqual(errors, []);
    console.log('PASS: choice tool, queued prompt, repeated numeric output, stop/resume and retry; model and Office mocked.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
