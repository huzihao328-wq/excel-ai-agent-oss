const fs = require('node:fs');
const path = require('node:path');
const ts = require('typescript');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const root = path.resolve(__dirname, '..');
const compile = file => ts.transpileModule(fs.readFileSync(path.join(root, file), 'utf8'), {
  compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2020, jsx: ts.JsxEmit.ReactJSX, esModuleInterop: true },
}).outputText;

(async () => {
  const browser = await chromium.launch({ channel: process.env.UI_BROWSER_CHANNEL || 'msedge', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 380, height: 650 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.setContent('<div id="root"></div>');
    for (const file of ['react/umd/react.production.min.js', 'react-dom/umd/react-dom.production.min.js']) {
      await page.addScriptTag({ content: fs.readFileSync(path.join(root, 'node_modules', file), 'utf8') });
    }
    await page.addStyleTag({ content: fs.readFileSync(path.join(root, 'src/taskpane/motion.css'), 'utf8') + '\nbutton { display:block; width:100%; height:60px; }' });
    await page.addScriptTag({ content: `
      const motion = {}, list = {};
      const jsx = (type, props, key) => React.createElement(type, {...props, key});
      const require = name => name === 'react' ? React : name === 'react/jsx-runtime' ? {jsx, jsxs:jsx} : motion;
      (function(require, exports) { ${compile('src/taskpane/motion.ts')} })(require, motion);
      (function(require, exports) { ${compile('src/taskpane/components/MotionList.tsx')} })(require, list);
      motion.installMotionTokens();
      const root = ReactDOM.createRoot(document.querySelector('#root'));
      window.renderList = keys => ReactDOM.flushSync(() => root.render(React.createElement(list.default, {}, keys.map(key => React.createElement('button', {key, 'data-key':key}, key)))));
      window.renderList(['A', 'B', 'C']);
    ` });
    const result = await page.evaluate(async () => {
      const frame = () => new Promise(resolve => requestAnimationFrame(resolve));
      const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
      const node = key => document.querySelector(`[data-key="${key}"]`).parentElement;
      const a = node('A');
      renderList(['B', 'C', 'A']);
      await frame();
      const moving = a.getAnimations()[0];
      moving.pause(); moving.currentTime = 100;
      const before = a.getBoundingClientRect().top;
      renderList(['A', 'B', 'C']);
      const reversing = a.getAnimations()[0];
      reversing.pause(); reversing.currentTime = 0;
      const after = a.getBoundingClientRect().top;
      reversing.finish();
      const b = node('B');
      renderList(['A', 'C']);
      const exiting = b.inert && b.getAttribute('aria-hidden') === 'true';
      await wait(50);
      renderList(['A', 'B', 'C']);
      const reused = b === node('B') && !b.inert;
      await wait(350);
      const finalOrder = [...document.querySelectorAll('[data-key]')].map(element => element.dataset.key).join('');
      renderList(['A', 'C']);
      await wait(350);
      return { before, after, exiting, reused, finalOrder, removed: !document.querySelector('[data-key="B"]'), animations: document.getAnimations().length };
    });
    assert.ok(Math.abs(result.before - result.after) < 1, 'FLIP reorder jumps when reversed');
    assert.ok(result.exiting && result.reused && result.removed);
    assert.equal(result.finalOrder, 'ABC');
    assert.equal(result.animations, 0);
    await page.evaluate(() => renderList(['C', 'A']));
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.waitForFunction(() => document.getAnimations().length === 0);
    assert.deepEqual(errors, []);
    console.log('PASS: real shared component reorder reversal, inert deletion, reinsertion before exit, and animation cleanup');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
