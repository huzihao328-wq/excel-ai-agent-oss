const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const { sourceLoader } = require('./load-source.cjs');
const { DEFAULT_SETTINGS } = sourceLoader()('src/core/settings.ts');
const root = path.resolve(__dirname, '..');
const frame = data => `data: ${JSON.stringify(data)}\n\n`;
const done = 'data: [DONE]\n\n';
const completion = text => frame({ choices: [{ delta: { content: text } }] }) + frame({ choices: [{ delta: {}, finish_reason: 'stop' }] }) + done;
const toolCalls = (calls, content = '') =>
  (content ? frame({ choices: [{ delta: { content } }] }) : '') +
  frame({ choices: [{ delta: { tool_calls: calls.map(([id, name, args], index) => ({ index, id, type: 'function', function: { name, arguments: JSON.stringify(args) } })) }, finish_reason: 'tool_calls' }] }) + done;
const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function until(check, label) {
  for (let i = 0; i < 400; i++) { if (check()) return; await sleep(25); }
  throw new Error(`Timed out waiting for ${label}`);
}
const textOf = message => typeof message.content === 'string' ? message.content : (message.content || []).map(p => p.text || '').join('');
const lastUserText = body => textOf([...body.messages].reverse().find(m => m.role === 'user'));
const toolResult = (body, id) => JSON.parse(body.messages.find(m => m.role === 'tool' && m.tool_call_id === id).content);

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  try {
    // 用最窄的任务窗格尺寸：布局挤压、滚动夹紧这类问题在这里最容易出现（窄屏规则曾把提示行隐藏）
    const page = await browser.newPage({ viewport: { width: 300, height: 700 } });
    const errors = [];
    const requests = [];
    // 每个模型请求按顺序取一个脚本；脚本可以先等待测试放行（模拟模型还在思考）
    const script = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.addInitScript(settings => {
      if (!localStorage.getItem('excel-ai-agent-settings')) localStorage.setItem('excel-ai-agent-settings', JSON.stringify(settings));
    }, { ...DEFAULT_SETTINGS, baseUrl: 'http://localhost:11434/v1', model: 'ui-test', apiKey: '', useMemory: false, autoLearnSkills: false, autoCheckpoint: false, askDirection: false });
    await page.route('**/*', async route => {
      const url = new URL(route.request().url());
      if (url.hostname === 'localhost' && url.pathname === '/v1/chat/completions') {
        const body = route.request().postDataJSON();
        requests.push(body);
        const step = script.shift();
        if (!step) throw new Error(`Unexpected model request ${requests.length}`);
        const response = await step(body);
        try {
          await route.fulfill({ contentType: 'text/event-stream', headers: { 'Access-Control-Allow-Origin': '*' }, body: response });
        } catch { /* 请求已被停止任务中断 */ }
        return;
      }
      if (url.origin !== 'https://excel-agent.test') return route.abort();
      if (url.pathname.endsWith('/office-js/office.js')) {
        return route.fulfill({ contentType: 'application/javascript', body: 'window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:"selection"},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:()=>{},removeHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error("Office mocked"))};' });
      }
      if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
      const asset = path.resolve(root, 'dist', '.' + url.pathname);
      if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
      return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
    });
    await page.goto('https://excel-agent.test/taskpane.html');
    const input = page.getByRole('textbox', { name: '任务指令' });
    const send = async text => { await input.fill(text); await page.getByTitle('发送', { exact: true }).click(); };

    // 1. 快捷任务：会改表的只填入输入框，不直接发请求；两类样式可区分
    const firstShortcut = page.locator('.launcher-scroll .chip').first();
    assert.ok((await firstShortcut.getAttribute('class')).includes('fill'));
    await firstShortcut.click();
    assert.ok((await input.inputValue()).length > 0);
    assert.equal(requests.length, 0);
    await input.fill('');
    console.log('PASS: shortcuts that may write only fill the composer');

    // 2. 任务执行中补充说明：下一次模型请求就带上；结束后执行过程折叠，失败步骤仍露在外面
    const thinking = deferred();
    script.push(
      () => toolCalls([['files', 'list_external_files', {}], ['r1', 'read_range', { sheet: 'S', range: 'A1:B2' }], ['r2', 'create_sheet', { name: '结果' }]], '先读取两块数据。'),
      async () => { await thinking.promise; return completion('初步完成。'); },
      body => {
        assert.ok(lastUserText(body).includes('[任务进行中用户补充] 改用 B 列'), 'steering note must reach the model');
        return completion('已按 B 列调整。');
      },
    );
    await send('处理数据');
    await until(() => requests.length === 2, 'second model request');
    // 前面的读取失败后，同批里会改表的调用不执行（只读调用照常执行）
    assert.match(toolResult(requests[1], 'r2').error, /本工具未执行/, 'a failed prerequisite must skip the writes left in the batch');
    await input.fill('改用 B 列');
    await input.press('Enter');
    await page.locator('.user-bubble.steer.pending').getByText('改用 B 列').waitFor();
    assert.equal(await input.inputValue(), '');
    thinking.resolve();
    await page.getByText('已按 B 列调整。', { exact: true }).waitFor();
    assert.equal(requests.length, 3);
    await page.locator('.user-bubble.steer:not(.pending)').waitFor();
    const group = page.locator('.process-group');
    await group.getByText('已执行 2 步 · 1 步未成功').waitFor();
    assert.equal(await page.locator('.process-group-failed .tool-card.error').count(), 1);
    await sleep(600);
    assert.equal(await page.locator('.jump-latest').count(), 0, 'finishing a task must keep the list pinned to the bottom');
    // 展开时失败步骤不能瞬间消失：内容高度骤减会把滚动位置夹紧，点的那一行随之跳走
    // 场景：停在底部、折叠头就在视野里（窗格稍高一点即如此），这时内容骤减会直接夹紧滚动位置
    await page.setViewportSize({ width: 300, height: 820 });
    await page.locator('.chat-list').evaluate(el => { el.scrollTop = el.scrollHeight; });
    await sleep(300);
    const head = group.locator('.process-group-head');
    const listTop = (await page.locator('.chat-list').boundingBox()).y;
    const before = (await head.boundingBox()).y;
    assert.ok(before > listTop, 'group header should be visible while pinned to the bottom');
    await head.click();
    await group.locator('.process-group-body').getByText('先读取两块数据。').waitFor();
    await sleep(700);
    assert.ok(Math.abs((await head.boundingBox()).y - before) < 2, `expanding a group must not move the header the user clicked (${before} -> ${(await head.boundingBox()).y})`);
    await page.setViewportSize({ width: 300, height: 700 });
    console.log('PASS: mid-task note is delivered on the next request; finished steps collapse with failures visible');

    // 3. 后续建议默认只填入输入框
    const followup = page.locator('.followup-chip').first();
    const followupPrompt = (await followup.getAttribute('title')).replace(/^填入输入框，可修改后发送：/, '');
    await followup.click();
    assert.equal(await input.inputValue(), followupPrompt);
    assert.equal(requests.length, 3);
    await input.fill('');
    console.log('PASS: follow-up suggestions fill the composer instead of running');

    // 4. 选择卡片可以写自己的做法
    script.push(
      () => toolCalls([['c1', 'ask_user_choice', { title: '选择方向', options: [{ id: 'a', label: '汇总' }, { id: 'b', label: '清洗' }] }]]),
      body => {
        const result = toolResult(body, 'c1');
        assert.equal(result.ok, true);
        assert.equal(result.data.custom, '按客户汇总，只看 2024 年');
        return completion('已按客户汇总。');
      },
    );
    await send('选方向');
    await page.getByRole('button', { name: '其他做法…' }).click();
    await sleep(500);
    assert.equal(await page.locator('.jump-latest').count(), 0, 'opening the custom-direction note must keep the list pinned to the bottom');
    await page.getByRole('textbox', { name: '其他做法' }).fill('按客户汇总，只看 2024 年');
    await page.getByRole('button', { name: '按这个做' }).click();
    await page.getByText('已按客户汇总。', { exact: true }).waitFor();
    await page.locator('.choice-resolved').getByText('你的做法：「按客户汇总，只看 2024 年」').waitFor();
    console.log('PASS: choice card accepts a custom direction');

    // 5. 本次任务中同类操作不再询问；下一次任务重新询问，并可拒绝并说明
    script.push(
      () => toolCalls([['d1', 'delete_sheet', { name: '旧表A' }]]),
      () => toolCalls([['d2', 'delete_sheet', { name: '旧表B' }]]),
      () => completion('两张表都处理了。'),
      () => toolCalls([['d3', 'delete_sheet', { name: '旧表C' }]]),
      body => {
        const result = toolResult(body, 'd3');
        assert.equal(result.ok, false);
        assert.match(result.error, /只删空表/);
        return completion('好的，只删空表。');
      },
    );
    await send('删掉旧表');
    const firstConfirm = page.locator('.confirm-card:not(.accepted):not(.rejected)');
    await firstConfirm.getByText('删除 Sheet「旧表A」').waitFor();
    await firstConfirm.getByLabel('本次任务中同类操作不再询问').check();
    await firstConfirm.getByRole('button', { name: '允许', exact: true }).click();
    await sleep(300);
    // 点过的按钮被移除时 WebView 会挪动滚动位置，不能被当成用户上翻
    assert.equal(await page.locator('.jump-latest').count(), 0, 'resolving a card must keep the list pinned to the bottom');
    await page.getByText('两张表都处理了。', { exact: true }).waitFor();
    // 任务结束后，答复过的确认随执行过程一起折叠；展开后两条记录都在，各自收成一行
    const deletions = page.locator('.process-group').last();
    await deletions.locator('.process-group-head').click();
    await deletions.locator('.confirm-card.auto').getByText('删除 Sheet「旧表B」').waitFor();
    assert.equal(await deletions.locator('.confirm-card.auto .confirm-auto-tag').innerText(), '已自动允许');
    assert.deepEqual(await deletions.locator('.confirm-card.settled .confirm-auto-tag').allInnerTexts(), ['已允许', '已自动允许']);
    assert.equal(await deletions.locator('.confirm-card .confirm-detail:visible').count(), 0, 'settled confirmations stay collapsed to one line');
    await deletions.locator('.process-group-head').click();
    await send('再删一张');
    const nextConfirm = page.locator('.confirm-card:not(.accepted):not(.rejected)');
    await nextConfirm.getByText('删除 Sheet「旧表C」').waitFor();
    await nextConfirm.getByRole('button', { name: '拒绝并说明…' }).click();
    await sleep(500);
    assert.equal(await page.locator('.jump-latest').count(), 0, 'opening the rejection note must keep the list pinned to the bottom');
    await page.getByRole('textbox', { name: '拒绝原因或新的要求' }).fill('只删空表');
    await page.getByRole('button', { name: '拒绝并发送说明' }).click();
    await page.getByText('好的，只删空表。', { exact: true }).waitFor();
    await page.locator('.confirm-resolved').getByText('已拒绝：只删空表').waitFor();
    console.log('PASS: allow-similar lasts one task; rejecting with a note reaches the model');

    // 6. 另一会话在执行时，本会话回车会排队并给出提示，而不是静默吞掉
    const background = deferred();
    script.push(
      async () => { await background.promise; return completion('长任务完成。'); },
      body => { assert.ok(lastUserText(body).includes('后台期间的指令')); return completion('排队指令完成。'); },
    );
    await send('长任务');
    await until(() => requests.length === 11, 'background request');
    await page.getByTitle('新对话', { exact: true }).click();
    await input.fill('后台期间的指令');
    await input.press('Enter');
    await page.locator('.queue-chip').getByText('后台期间的指令', { exact: false }).waitFor();
    await page.locator('.composer-flash').getByText('已排队', { exact: false }).waitFor();
    assert.equal(await input.inputValue(), '');
    background.resolve();
    await page.getByText('排队指令完成。', { exact: true }).waitFor();
    assert.equal(requests.length, 12);
    console.log('PASS: Enter during another conversation\'s run queues with feedback and sends afterwards');

    // 7. 停止任务时没送达的补充说明放回输入框
    const never = deferred();
    script.push(async () => { await never.promise; return completion('不会显示'); });
    await send('会被停止');
    await until(() => requests.length === 13, 'stoppable request');
    await input.fill('这条来不及');
    await input.press('Enter');
    await page.locator('.user-bubble.steer.pending').waitFor();
    await page.getByTitle('停止', { exact: true }).click();
    await page.getByText('任务已结束，补充说明没来得及送达，已放回输入框。', { exact: true }).waitFor();
    assert.equal(await input.inputValue(), '这条来不及');
    assert.equal(await page.locator('.user-bubble.steer').count(), 0);
    never.resolve();
    console.log('PASS: undelivered notes return to the composer when a task stops');

    // 8. 在空白新对话里排队后又离开：该对话不会保存、也回不去，排队的指令必须放回输入框
    await input.fill('');
    const held = deferred();
    script.push(async () => { await held.promise; return completion('后台任务结束。'); });
    await send('再跑一个长任务');
    await until(() => requests.length === 14, 'held request');
    await page.getByTitle('新对话', { exact: true }).click();
    await input.fill('在空白对话里排队');
    await input.press('Enter');
    await page.locator('.queue-chip').waitFor();
    await page.locator('.bg-run-banner').getByRole('button', { name: '返回' }).click();
    await page.locator('.composer-flash').getByText('排队的指令已放回输入框。').waitFor();
    assert.equal(await input.inputValue(), '在空白对话里排队');
    held.resolve();
    await page.getByText('后台任务结束。', { exact: true }).waitFor();
    assert.equal(requests.length, 14, 'reclaimed instruction must not be sent on its own');
    console.log('PASS: queued instruction in an abandoned blank chat returns to the composer');

    // 9. 补充已进入历史后请求失败：重试继续同一任务，不能把原指令重新排到补充之后。
    const beforeRetry = requests.length;
    const retryThinking = deferred();
    let retriedBody;
    script.push(
      async () => { await retryThinking.promise; return completion('初步结果。'); },
      () => frame({ error: { message: 'steering request rejected' } }) + done,
      body => {
        retriedBody = body;
        return completion('已按补充要求完成分析。');
      },
    );
    await send('生成并写入结果');
    await until(() => requests.length === beforeRetry + 1, 'request before steering retry');
    await input.fill('改为只分析，不要写入');
    await input.press('Enter');
    retryThinking.resolve();
    await page.getByRole('button', { name: '重试这条指令' }).click();
    await page.getByText('已按补充要求完成分析。', { exact: true }).waitFor();
    assert.equal(requests.length, beforeRetry + 3);
    assert.match(lastUserText(retriedBody), /改为只分析，不要写入/);
    assert.equal(retriedBody.messages.filter(m => m.role === 'user' && textOf(m).includes('生成并写入结果')).length, 1);
    console.log('PASS: retry after delivered steering preserves the latest instruction without duplicating the original');

    assert.equal(script.length, 0, 'every scripted model response was used');
    assert.deepEqual(errors, []);
    console.log('PASS: interaction flows; model and Office mocked.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
