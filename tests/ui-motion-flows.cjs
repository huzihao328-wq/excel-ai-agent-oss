const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const XLSX = require('xlsx');
const root = path.resolve(__dirname, '..');

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  const context = await browser.newContext({ viewport: { width: 380, height: 850 }, hasTouch: true });
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.setDefaultTimeout(10000);
  await context.route('**/*', route => {
    const url = new URL(route.request().url());
    if (url.origin !== 'https://excel-agent.test') return route.abort();
    if (url.pathname.endsWith('/office-js/office.js')) return route.fulfill({ contentType: 'application/javascript', body: 'window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:"selection"},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error("Office mocked"))};' });
    if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
    const asset = path.resolve(root, 'dist', '.' + url.pathname);
    if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
    return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
  });
  const settle = () => page.evaluate(() => document.getAnimations().forEach(animation => {
    if (animation.effect?.getTiming().iterations !== Infinity) animation.finish();
  }));
  try {
    await page.goto('https://excel-agent.test/taskpane.html');
    await page.locator('.welcome-panel').waitFor();
    await page.getByTitle('模型设置', { exact: true }).click();
    const dialog = page.getByRole('dialog', { name: '模型设置' });
    await dialog.waitFor();
    assert.ok(await dialog.evaluate(element => element.contains(document.activeElement)));
    await dialog.evaluate(element => { element.querySelector('button').focus(); });
    await page.keyboard.press('Shift+Tab');
    assert.ok(await dialog.evaluate(element => element.contains(document.activeElement)));
    await page.keyboard.press('Tab');
    assert.equal(await page.evaluate(() => document.activeElement.title), '关闭');
    await page.keyboard.press('Escape');
    await dialog.waitFor({ state: 'detached' });
    assert.equal(await page.evaluate(() => document.activeElement.title), '模型设置');
    console.log('PASS: modal Tab loop and focus returns to its opener');

    await page.getByRole('button', { name: '全部任务', exact: false }).click();
    await page.getByRole('group', { name: '任务分类' }).getByRole('button', { name: '分析', exact: true }).click();
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.getByRole('button', { name: '整理', exact: true }).getAttribute('aria-pressed'), 'true');
    await page.keyboard.press('Home');
    assert.equal(await page.getByRole('group', { name: '任务分类' }).getByRole('button', { name: '全部', exact: true }).getAttribute('aria-pressed'), 'true');
    const reuse = await page.evaluate(async () => {
      const input = document.querySelector('[aria-label="搜索快捷任务"]');
      const first = document.querySelector('.task-results .motion-list-item');
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      set.call(input, '不存在的任务'); input.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(resolve => requestAnimationFrame(resolve));
      const inert = first.inert;
      set.call(input, ''); input.dispatchEvent(new Event('input', { bubbles: true }));
      await new Promise(resolve => setTimeout(resolve, 350));
      return { inert, same: first === document.querySelector('.task-results .motion-list-item'), live: !first.inert };
    });
    assert.ok(reuse.inert && reuse.same && reuse.live, 'filter reversal replaced or disabled a live item');
    await settle();
    await page.screenshot({ path: path.join(__dirname, 'motion-tasks.png') });
    await page.keyboard.press('Escape');
    await page.locator('.task-catalog').waitFor({ state: 'detached' });
    console.log('PASS: keyboard category switching and interrupted list removal reuses live nodes');

    const input = page.getByRole('textbox', { name: '任务指令' });
    await input.fill('多行草稿\n'.repeat(20));
    await page.waitForFunction(() => document.querySelector('.input-row textarea').clientHeight >= 150);
    await input.fill('短草稿');
    const heightReversal = await input.evaluate(async element => {
      const frame = () => new Promise(resolve => requestAnimationFrame(resolve));
      const current = element.getAnimations().find(animation => animation.effect.getKeyframes().some(keyframe => 'height' in keyframe));
      if (!current) throw Error('No textarea height transition');
      current.pause(); current.currentTime = 60;
      const before = element.getBoundingClientRect().height;
      const set = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set;
      set.call(element, '继续输入\n'.repeat(20)); element.dispatchEvent(new Event('input', { bubbles: true }));
      await frame();
      const next = element.getAnimations().find(animation => animation !== current && animation.effect.getKeyframes().some(keyframe => 'height' in keyframe));
      if (!next) throw Error('Interrupted resize did not retarget');
      next.pause(); next.currentTime = 0;
      const after = element.getBoundingClientRect().height;
      next.finish();
      return { before, after };
    });
    assert.ok(Math.abs(heightReversal.before - heightReversal.after) < 1, 'textarea resize jumped on reversal');
    await input.fill('');
    console.log('PASS: textarea resize reverses from its current height');

    const book = XLSX.utils.book_new();
    for (const name of ['销售', '库存', '费用']) XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(Array.from({ length: 180 }, (_, r) => Array.from({ length: 12 }, (_, c) => `${name}-${r}-${c}`))), name);
    await page.locator('.composer input[type=file]').setInputFiles({ name: 'motion-sample.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }) });
    await page.locator('.file-chip-name').getByText('motion-sample.xlsx', { exact: false }).click();
    await page.getByRole('tab', { name: '库存' }).click();
    await page.keyboard.press('ArrowRight');
    assert.equal(await page.getByRole('tab', { name: '费用' }).getAttribute('aria-selected'), 'true');
    await page.setViewportSize({ width: 600, height: 850 });
    const a = await page.locator('[data-hit="cell"][data-r="1"][data-c="0"]').boundingBox();
    const b = await page.locator('[data-hit="cell"][data-r="5"][data-c="2"]').boundingBox();
    await page.mouse.move(a.x + 5, a.y + 5); await page.mouse.down();
    await page.mouse.move(b.x + 5, b.y + 5, { steps: 30 }); await page.mouse.up();
    await page.locator('.preview-sel-addr').getByText('已选: A2:C6', { exact: true }).waitFor();
    assert.equal(await page.locator('.preview-grid-wrap').evaluate(element => getComputedStyle(element).touchAction), 'pan-x pan-y');
    const grid = await page.locator('.preview-grid-wrap').boundingBox();
    const cdp = await context.newCDPSession(page);
    await cdp.send('Input.synthesizeScrollGesture', { x: grid.x + 120, y: grid.y + grid.height - 35, yDistance: -180, speed: 600, gestureSourceType: 'touch' });
    const nativeTouchScrollObserved = await page.locator('.preview-grid-wrap').evaluate(element => element.scrollTop > 30);
    assert.equal(await page.locator('.preview-sel-addr').innerText(), '已选: A2:C6', 'touch scrolling changed the selection');
    assert.equal(await page.locator('.data-cell').first().evaluate(element => getComputedStyle(element).touchAction), 'pan-x pan-y');
    await page.locator('.preview-grid-wrap').evaluate(element => { element.scrollTop = 0; });
    const tap = await page.locator('[data-hit="cell"][data-r="2"][data-c="1"]').boundingBox();
    await page.touchscreen.tap(tap.x + 5, tap.y + 5);
    await page.locator('.preview-sel-addr').getByText('已选: B3', { exact: true }).waitFor();
    await page.mouse.move(grid.x + 120, grid.y + 120);
    await page.mouse.wheel(0, 180);
    await page.waitForFunction(() => document.querySelector('.preview-grid-wrap').scrollTop > 30);
    await settle();
    await page.screenshot({ path: path.join(__dirname, 'motion-preview.png') });
    await page.keyboard.press('Escape');
    await page.locator('.preview-panel').waitFor({ state: 'detached' });
    console.log('PASS: sheet tabs, keyboard navigation, drag selection, touch taps, and wheel scrolling');
    if (!nativeTouchScrollObserved) console.log('LIMIT: native touch scrolling was not observed in headless Edge; also reproduced on a plain scroll container');

    // Seed only this isolated test origin; no personal conversations or Excel data are touched.
    await page.evaluate(() => {
      const meta = { id: 'motion-chat', title: '长对话', createdAt: Date.now(), updatedAt: Date.now() };
      localStorage.setItem('excel-ai-agent-conv-index', JSON.stringify({ currentId: meta.id, items: [meta] }));
      localStorage.setItem('excel-ai-agent-conv:' + meta.id, JSON.stringify({ ...meta, messages: [], items: [
        ...Array.from({ length: 400 }, (_, i) => ({ kind: 'assistant', id: `message-${i}`, text: `历史消息 ${i}：| 数据 | 结果 |\n| --- | --- |\n| 合计 | 100 |` })),
        { kind: 'tool', id: 'tool', label: '检查结果', name: 'read_range', status: 'ok', argsSummary: 'A1:C6', detail: '可展开的结果\n'.repeat(12) },
      ] }));
    });
    await page.reload();
    await page.locator('.tool-card-head').waitFor();
    await page.locator('.tool-card-head').click();
    await page.locator('.tool-detail').waitFor();
    await page.evaluate(async () => {
      for (let i = 0; i < 20; i++) {
        document.querySelector('.tool-card-head').click();
        await new Promise(resolve => setTimeout(resolve, 18));
      }
    });
    await settle();
    assert.equal(await page.locator('.tool-card-head').getAttribute('aria-expanded'), 'true');
    assert.ok(await page.locator('.tool-detail').evaluate(element => element.getBoundingClientRect().height > 0));
    await page.locator('.tool-card-head').focus();
    await page.keyboard.press('Enter');
    await page.locator('.tool-detail').waitFor({ state: 'detached' });
    // Browse earlier messages: input resizing must not pull the user back to the bottom.
    await page.locator('.chat-list').evaluate(element => { element.dispatchEvent(new WheelEvent('wheel', { deltaY: -400 })); element.scrollTop = 200; });
    await input.fill('输入不会抢走历史滚动位置\n'.repeat(20));
    await page.waitForFunction(() => document.querySelector('.input-row textarea').clientHeight >= 150);
    assert.ok(await page.locator('.chat-list').evaluate(element => Math.abs(element.scrollTop - 200) < 2));
    console.log('PASS: repeated disclosure reversal, Enter control, and 400-message scroll preservation');

    await page.evaluate(() => {
      const stored = JSON.parse(localStorage.getItem('excel-ai-agent-conv-index'));
      stored.items = Array.from({ length: 500 }, (_, i) => ({ id: `stress-${i}`, title: `历史任务 ${String(i).padStart(3, '0')}`, createdAt: Date.now(), updatedAt: Date.now() }));
      localStorage.setItem('excel-ai-agent-conv-index', JSON.stringify(stored));
    });
    await page.getByTitle('历史对话', { exact: true }).click();
    await page.waitForFunction(() => document.querySelectorAll('.history-item').length === 500);
    await settle();
    const perf = await page.evaluate(async () => {
      const frames = [], longTasks = [];
      let previous = performance.now(), running = true, peakAnimations = 0;
      const observer = new PerformanceObserver(list => longTasks.push(...list.getEntries().map(entry => entry.duration)));
      observer.observe({ type: 'longtask' });
      const sample = now => {
        frames.push(now - previous); previous = now;
        peakAnimations = Math.max(peakAnimations, document.getAnimations().length);
        if (running) requestAnimationFrame(sample);
      };
      requestAnimationFrame(sample);
      const search = document.querySelector('.history-search input');
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
      for (let i = 0; i < 24; i++) {
        set.call(search, i % 2 ? '' : `历史任务 ${i % 5}`);
        search.dispatchEvent(new Event('input', { bubbles: true }));
        await new Promise(resolve => setTimeout(resolve, 65));
      }
      await new Promise(resolve => setTimeout(resolve, 400));
      running = false; observer.disconnect();
      const sorted = frames.filter(value => value > 0).sort((a, b) => a - b);
      return { historyRows: 500, chatMessages: 400, filterChanges: 24, frames: frames.length,
        frameP95Ms: sorted[Math.floor(sorted.length * 0.95)], maxFrameMs: Math.max(...sorted),
        longTasks: longTasks.length, maxLongTaskMs: Math.max(0, ...longTasks), peakAnimations,
        remainingAnimations: document.getAnimations().filter(animation => animation.effect.getTiming().iterations !== Infinity).length,
        exitingRows: document.querySelectorAll('.is-removing').length,
      };
    });
    assert.equal(perf.exitingRows, 0); assert.equal(perf.remainingAnimations, 0);
    assert.ok(perf.peakAnimations <= 80, 'long list created unbounded simultaneous animations');
    fs.writeFileSync(path.join(__dirname, 'motion-performance.json'), JSON.stringify({ environment: 'Headless Edge; Office mocked; wall-clock observation, not real Excel frame-rate guarantee', nativeTouchScrollObserved, ...perf }, null, 2));
    console.log('PERFORMANCE:', JSON.stringify(perf));
    await page.locator('.history-item').nth(1).getByTitle('删除对话').click();
    await page.locator('.history-item-confirm').getByRole('button', { name: '删除', exact: true }).click();
    await page.waitForFunction(() => document.querySelectorAll('.history-item').length === 499);
    console.log('PASS: confirmed history deletion updates business data and completes its exit');

    for (const width of [300, 380, 600]) {
      await page.setViewportSize({ width, height: 650 });
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
    }
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.getByPlaceholder('搜索对话…').fill('没有匹配');
    assert.equal(await page.locator('.is-removing').count(), 0);
    await page.keyboard.press('Escape');
    await page.locator('.history-panel').waitFor({ state: 'detached' });
    assert.deepEqual(errors, []);
    console.log('PASS: narrow layouts and live reduced-motion changes cancel layout animation');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
