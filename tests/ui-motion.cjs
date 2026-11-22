const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '..');

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 380, height: 850 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.setDefaultTimeout(8000);
  await page.route('**/*', route => {
    const url = new URL(route.request().url());
    if (url.origin !== 'https://excel-agent.test') return route.abort();
    if (url.pathname.endsWith('/office-js/office.js')) {
      return route.fulfill({ contentType: 'application/javascript', body: 'window.Office={onReady:fn=>fn(),EventType:{DocumentSelectionChanged:"selection"},context:{document:{settings:{get:()=>null,set:()=>{},remove:()=>{},saveAsync:()=>{}},addHandlerAsync:()=>{}}}};window.Excel={run:()=>Promise.reject(new Error("Office mocked"))};' });
    }
    if (url.pathname.startsWith('/api/')) return route.fulfill({ status: 404, body: '{}' });
    const asset = path.resolve(root, 'dist', '.' + url.pathname);
    if (!asset.startsWith(path.join(root, 'dist') + path.sep) || !fs.existsSync(asset)) return route.abort();
    return route.fulfill({ body: fs.readFileSync(asset), contentType: asset.endsWith('.js') ? 'application/javascript' : asset.endsWith('.html') ? 'text/html' : 'application/octet-stream' });
  });
  try {
    await page.goto('https://excel-agent.test/taskpane.html');
    await page.locator('.welcome-panel').waitFor();
    await page.getByTitle('更多功能', { exact: true }).click();
    await page.waitForFunction(() => Number(getComputedStyle(document.querySelector('.header-menu')).opacity) === 1);

    // Seek the real CSS transitions so reversal checks do not depend on machine speed.
    const reversal = await page.evaluate(async () => {
      const frame = () => new Promise(resolve => requestAnimationFrame(resolve));
      const menu = document.querySelector('.header-menu');
      const presence = menu.closest('.presence');
      const toggle = document.querySelector('[title="更多功能"]');
      const opacityTransition = () => menu.getAnimations().find(animation => animation.transitionProperty === 'opacity');
      const waitForTransition = async previous => {
        for (let i = 0; i < 12; i++) {
          await frame();
          const animation = opacityTransition();
          if (animation && animation !== previous) return animation;
        }
        throw Error('Expected a continuous opacity transition');
      };
      toggle.click();
      const closing = await waitForTransition();
      closing.pause();
      closing.currentTime = closing.effect.getTiming().duration * 0.45;
      const beforeReopen = Number(getComputedStyle(menu).opacity);
      menu.querySelector('button').focus();
      const exitInert = presence.inert && !presence.contains(document.activeElement);
      const focusAtExit = { inert: presence.inert, active: document.activeElement.outerHTML.slice(0, 280), trigger: toggle.outerHTML.slice(0, 280) };
      toggle.click();
      const opening = await waitForTransition(closing);
      opening.pause();
      opening.currentTime = 0;
      const afterReopen = Number(getComputedStyle(menu).opacity);
      opening.currentTime = opening.effect.getTiming().duration * 0.35;
      const beforeClose = Number(getComputedStyle(menu).opacity);
      toggle.click();
      const closingAgain = await waitForTransition(opening);
      closingAgain.pause();
      closingAgain.currentTime = 0;
      const afterClose = Number(getComputedStyle(menu).opacity);
      // A further reopen must cancel the obsolete unmount timer.
      toggle.click();
      await frame();
      await frame();
      await new Promise(resolve => setTimeout(resolve, 350));
      return {
        beforeReopen, afterReopen, beforeClose, afterClose, exitInert, focusAtExit,
        sameNode: document.querySelector('.header-menu') === menu,
        open: !presence.inert && Number(getComputedStyle(menu).opacity) === 1,
      };
    });
    assert.ok(reversal.exitInert, 'closing menu still accepts keyboard focus: ' + JSON.stringify(reversal.focusAtExit));
    assert.ok(Math.abs(reversal.beforeReopen - reversal.afterReopen) < 0.02, 'reopening jumps to an endpoint');
    assert.ok(Math.abs(reversal.beforeClose - reversal.afterClose) < 0.02, 'interrupted entry jumps before closing');
    assert.ok(reversal.sameNode && reversal.open, 'stale exit timer removes reopened menu');
    await page.keyboard.press('Escape');
    await page.locator('.header-menu').waitFor({ state: 'detached' });
    console.log('PASS: continuous reversals in both directions, inert exit, obsolete timer cancellation');

    await page.getByTitle('更多功能', { exact: true }).click();
    await page.getByRole('button', { name: '批量任务', exact: true }).click();
    await page.locator('.batch-instruction textarea').fill('保留这条批量任务草稿');
    await page.locator('.batch-panel .icon-btn').first().click();
    await page.locator('.batch-panel').waitFor({ state: 'hidden' });
    await page.getByTitle('更多功能', { exact: true }).click();
    await page.getByRole('button', { name: '批量任务', exact: true }).click();
    assert.equal(await page.locator('.batch-instruction textarea').inputValue(), '保留这条批量任务草稿');
    await page.keyboard.press('Escape');
    await page.locator('.batch-panel').waitFor({ state: 'hidden' });
    console.log('PASS: retained panel state survives animated close and reopen');

    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.reload();
    await page.locator('.welcome-panel').waitFor();
    const reduced = await page.locator('.welcome-tasks > button').evaluateAll(elements => elements.every(element => {
      const style = getComputedStyle(element);
      return style.animationDelay === '0s' && parseFloat(style.animationDuration) < 0.001;
    }));
    assert.ok(reduced, 'reduced motion still delays welcome content');
    await page.getByRole('button', { name: '全部任务', exact: false }).click();
    await page.getByRole('textbox', { name: '搜索快捷任务' }).fill('公式');
    await page.keyboard.press('Escape');
    assert.equal(await page.locator('.task-catalog').count(), 0, 'reduced motion waits for exit animation');
    assert.deepEqual(errors, []);
    console.log('PASS: reduced motion has no stagger or delayed dismissal; Office mocked');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
