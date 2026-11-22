const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const { DEFAULT_SETTINGS } = sourceLoader()('src/core/settings.ts');
const settings = { ...DEFAULT_SETTINGS, autoCheckpoint: false, useMemory: false };
const clone = value => JSON.parse(JSON.stringify(value));

// 只模拟本组回归需要的 Office 对象行为；不宣称替代桌面 Excel 验收。
function excelMock(t, rows = [['名称', '金额'], ['甲', 11], ['乙', 22]]) {
  const previous = global.Excel;
  const state = { rows: clone(rows), writes: 0, deleted: [], deduplicated: 0, frozen: [], sort: null, syncHook: null, active: null, validation: null, hidden: [], merged: [] };
  const worksheets = [];
  function makeSheet(name) {
    const sheet = { name, load() {}, activate() { state.active = sheet; },
      copy() { const copy = makeSheet(`${sheet.name} (2)`); worksheets.push(copy); return copy; },
      freezePanes: {
        unfreeze() { state.frozen.push(['none']); },
        freezeRows(n) { state.frozen.push(['rows', n]); },
        freezeColumns(n) { state.frozen.push(['columns', n]); },
        freezeAt(range) { state.frozen.push(['range', range.rowIndex, range.columnIndex, range.rowCount, range.columnCount]); },
      },
    };
    function range(addr = 'A1:B3', rowIndex = 0, columnIndex = 0, rowCount = state.rows.length, columnCount = state.rows[0]?.length || 1) {
      const r = {
        address: `${sheet.name}!${addr}`, isNullObject: false, rowIndex, columnIndex, rowCount, columnCount,
        load() {},
        get values() { return clone(state.rows); },
        get formulas() { return clone(state.formulaRows ?? state.rows); },
        get numberFormat() { return state.rows.map(row => row.map(() => 'General')); },
        set values(grid) { state.writes++; state.rows = clone(grid); },
        set formulas(grid) { state.writes++; if (state.calc) { state.formulaRows = clone(grid); state.rows = state.calc(grid); } else state.rows = clone(grid); },
        getCell(row, col) { return range(addr, rowIndex + row, columnIndex + col, 1, 1); },
        getAbsoluteResizedRange(nr, nc) { return range(addr, rowIndex, columnIndex, nr, nc); },
        removeDuplicates(columns, headers) { state.deduplicated++; return { removed: 1, uniqueRemaining: 1, load() {} }; },
        delete() { state.deleted.push(addr); },
        sort: { apply(fields, matchCase, headers) { state.sort = { fields, headers }; } },
        format: { font: {}, fill: {}, borders: { getItem() { return {}; } }, autofitColumns() {} },
        dataValidation: { clear() { state.validation = { cleared: addr }; }, set rule(rule) { state.validation.rule = rule; }, set ignoreBlanks(v) { state.validation.ignoreBlanks = v; }, set errorAlert(v) { state.validation.errorAlert = v; }, set prompt(v) { state.validation.prompt = v; } },
        set rowHidden(v) { state.hidden.push(['rows', addr, v]); },
        set columnHidden(v) { state.hidden.push(['columns', addr, v]); },
        merge() { state.merged.push(addr); },
        unmerge() {},
        getIntersectionOrNullObject() { return r; },
      };
      return r;
    }
    sheet.getRange = addr => {
      const cell = /^([A-Z])(\d+)$/.exec(addr);
      return cell ? range(addr, Number(cell[2]) - 1, cell[1].charCodeAt(0) - 65, 1, 1) : range(addr);
    };
    sheet.getUsedRangeOrNullObject = () => range(`A1:B${state.rows.length}`);
    sheet.getUsedRange = sheet.getUsedRangeOrNullObject;
    sheet.getRangeByIndexes = (r, c, nr, nc) => range('byIndexes', r, c, nr, nc);
    return sheet;
  }
  worksheets.push(makeSheet('S'), makeSheet('Other'));
  state.active = worksheets[1];
  const collection = {
    items: worksheets, load() {},
    getItem(name) { const sheet = worksheets.find(x => x.name === name); if (!sheet) throw Object.assign(new Error("The requested resource doesn't exist."), { code: 'ItemNotFound' }); return sheet; },
    getItemOrNullObject(name) { return worksheets.find(x => x.name === name) || { isNullObject: true, load() {} }; },
    getActiveWorksheet() { return state.active; },
  };
  const context = { workbook: { worksheets: collection }, sync: async () => { await state.syncHook?.(); } };
  global.Excel = { run: fn => fn(context), WorksheetPositionType: { after: 'after' }, DeleteShiftDirection: { up: 'up' }, HorizontalAlignment: { center: 'center' }, BorderLineStyle: { continuous: 'continuous' }, DataValidationAlertStyle: { stop: 'Stop', warning: 'Warning' }, SheetVisibility: { visible: 'Visible', hidden: 'Hidden' } };
  t.after(() => { global.Excel = previous; });
  return { state, worksheets };
}
function runtime(overrides) {
  return sourceLoader(overrides)('src/core/tools/excelTools.ts');
}
function ctx(overrides = {}) { return { settings, confirm: async () => true, choose: async () => ({ kind: 'skipped' }), ...overrides }; }
const execute = (rt, name, args, context = ctx()) => rt.executeTool(name, JSON.stringify(args), context);

test('misspelled tool fields return actionable schema feedback before any workbook access', async t => {
  const previous = global.Excel;
  t.after(() => { global.Excel = previous; });
  let accessed = false;
  global.Excel = { run() { accessed = true; throw new Error('should not run'); } };
  const result = await execute(runtime(), 'read_range', { sheet: 'S', ranges: 'A1:B3' });
  assert.equal(result.ok, false);
  assert.match(result.error, /ranges/);
  assert.match(result.error, /允许字段.*range/);
  assert.match(result.error, /本次未执行/);
  assert.equal(accessed, false);
});

test('undo restores text identifiers, dates, booleans and apostrophes without changing cell types', async t => {
  const previous = global.Excel;
  t.after(() => { global.Excel = previous; });
  const original = [['00123', '12345678901234567890', '2024-03', 'TRUE', "'quoted", '=literal', 7, true, 2]];
  const originalFormulas = clone(original);
  originalFormulas[0][8] = '=1+1';
  let values = clone(original);
  let formulas = clone(originalFormulas);
  // Simulate Excel's typed-input conversion; quote prefixes must preserve literal strings.
  function assign(grid) {
    values = grid.map(row => row.map(value => {
      if (typeof value !== 'string') return value;
      if (value.startsWith("'")) return value.slice(1);
      if (/^\d+$/.test(value)) return Number(value);
      if (value === 'TRUE') return true;
      if (value === '2024-03') return 45352;
      if (value === '=1+1') return 2;
      return value;
    }));
    formulas = grid.map((row, r) => row.map((value, c) =>
      typeof value === 'string' && value.startsWith('=') ? value : values[r][c]));
  }
  const range = {
    address: 'S!A1:I1', load() {}, getAbsoluteResizedRange() { return this; },
    get values() { return clone(values); }, set values(grid) { assign(grid); },
    get formulas() { return clone(formulas); }, set formulas(grid) { assign(grid); },
  };
  const sheet = { isNullObject: false, load() {}, getRange: () => range };
  global.Excel = { run: fn => fn({ workbook: { worksheets: { getItem: () => sheet, getItemOrNullObject: () => sheet } }, sync: async () => {} }) };
  const rt = runtime();
  const result = await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values: [original[0].map(() => 0)] });
  assert.equal(result.ok, true, result.error);
  await rt.applyUndo(rt.getUndoStack()[0].id);
  assert.deepEqual(values, original);
  assert.deepEqual(formulas, originalFormulas);
  assert.equal(rt.getUndoStack().length, 0);
});

test('copy_sheet renames the returned copy and leaves both source and unrelated active sheet intact', async t => {
  const { state, worksheets } = excelMock(t);
  const result = await execute(runtime(), 'copy_sheet', { name: 'S', new_name: 'Copy' });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(worksheets.map(s => s.name), ['S', 'Other', 'Copy']);
  assert.equal(state.active.name, 'Copy');
});

test('freeze_panes uses exact row/column counts and B2 freezes A1 only', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  for (const [args, expected] of [
    [{}, ['rows', 1]], [{ columns: 2 }, ['columns', 2]], [{ rows: 2, columns: 3 }, ['range', 0, 0, 2, 3]],
    [{ cell: 'B2' }, ['range', 0, 0, 1, 1]], [{ rows: 0, columns: 0 }, ['none']],
  ]) {
    const result = await execute(rt, 'freeze_panes', { sheet: 'S', ...args });
    assert.equal(result.ok, true, result.error);
    assert.deepEqual(state.frozen.at(-1), expected);
  }
  const before = state.frozen.length;
  assert.equal((await execute(rt, 'freeze_panes', { sheet: 'S', rows: 1.5 })).ok, false);
  assert.equal(state.frozen.length, before);
});

test('polish_range freezes the header only and respects header=false', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  assert.equal((await execute(rt, 'polish_range', { sheet: 'S', number_formats: false })).ok, true);
  assert.deepEqual(state.frozen, [['rows', 1]]);
  const result = await execute(rt, 'polish_range', { sheet: 'S', header: false, number_formats: false });
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.freeze_header, false);
  assert.equal(state.frozen.length, 1);
});

test('nested sort ascending="false" reaches Excel as descending', async t => {
  const { state } = excelMock(t);
  const result = await execute(runtime(), 'sort_range', { sheet: 'S', keys: [{ column: '金额', ascending: 'false' }] });
  assert.equal(result.ok, true, result.error);
  assert.equal(state.sort.fields[0].ascending, false);
});

test('invalid arguments fail before checkpoint or workbook access', async t => {
  excelMock(t);
  let checkpoints = 0;
  const rt = runtime({ 'src/core/tools/checkpoint.ts': { ensureCheckpoint: async () => { checkpoints++; } } });
  const result = await execute(rt, 'write_range', { values: [[1]] }, ctx({ settings: { ...settings, autoCheckpoint: true } }));
  assert.equal(result.ok, false);
  assert.match(result.error, /sheet/);
  assert.equal(checkpoints, 0);
});

test('destructive tools check the sheet exists before asking the user to confirm', async t => {
  const { state } = excelMock(t);
  let asked = 0;
  const context = ctx({ confirm: async () => { asked++; return true; } });
  for (const [name, args] of [
    ['delete_rows', { sheet: '不存在', start_row: 2, count: 1 }],
    ['delete_columns', { sheet: '不存在', start_column: 'B', count: 1 }],
    ['delete_sheet', { name: '不存在' }],
    ['insert_rows', { sheet: '不存在', before_row: 2, count: 50 }],
  ]) {
    const result = await execute(runtime(), name, args, context);
    assert.equal(result.ok, false, name);
    assert.match(result.error, /工作表「不存在」不存在/, name);
    assert.match(result.error, /现有工作表：「S」、「Other」/, name + ' lists existing sheets once');
    assert.equal(result.error.split('现有工作表').length, 2, name);
  }
  assert.equal(asked, 0, 'no confirmation card for an operation that cannot succeed');
  assert.deepEqual(state.deleted, []);
  // 表存在时照常询问
  await execute(runtime(), 'delete_sheet', { name: 'Other' }, ctx({ confirm: async () => { asked++; return false; } }));
  assert.equal(asked, 1);
});

test('abort during checkpoint prevents the queued write', async t => {
  const { state } = excelMock(t);
  const controller = new AbortController();
  const rt = runtime({ 'src/core/tools/checkpoint.ts': { ensureCheckpoint: async () => { controller.abort(); } } });
  const result = await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values: [[1]] }, ctx({ signal: controller.signal, settings: { ...settings, autoCheckpoint: true } }));
  assert.equal(result.ok, false);
  assert.match(result.error, /停止/);
  assert.equal(state.writes, 0);
});

test('abort after confirmation prevents writing even if confirmation resolves true', async t => {
  const { state } = excelMock(t);
  const controller = new AbortController();
  const result = await execute(runtime(), 'write_range', { sheet: 'S', start_cell: 'A1', values: [[1]] }, ctx({ signal: controller.signal, confirm: async () => { controller.abort(); return true; } }));
  assert.equal(result.ok, false);
  assert.equal(state.writes, 0);
});

test('write rejects zero-width or object grids and preserves edits made during confirmation', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  for (const values of [[[]], [[{ amount: 10 }]]]) {
    assert.equal((await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values })).ok, false);
  }
  const result = await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values: [[1]] }, ctx({ confirm: async () => { state.rows[1][1] = 99; return true; } }));
  assert.equal(result.ok, false);
  assert.equal(state.rows[1][1], 99);
  assert.equal(state.writes, 0);
});

test('confirmDetailed: rejection note replaces the generic "ask the user" error and nothing is written', async t => {
  const { state } = excelMock(t);
  const asked = [];
  let legacy = 0;
  const result = await execute(runtime(), 'write_range', { sheet: 'S', start_cell: 'A1', values: [[1]] }, ctx({
    confirm: async () => { legacy++; return true; },
    confirmDetailed: async (title, detail, meta) => { asked.push(meta.tool); return { accepted: false, note: '写到 F 列，不要覆盖' }; },
  }));
  assert.equal(result.ok, false);
  assert.match(result.error, /写到 F 列，不要覆盖/);
  assert.match(result.error, /不要原样重试/);
  assert.doesNotMatch(result.error, /向用户询问意图/);
  assert.deepEqual(asked, ['write_range']);
  assert.equal(legacy, 0);
  assert.equal(state.writes, 0);
});

test('confirmDetailed: acceptance writes; plain rejection keeps the original error', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  const accepted = await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values: [[1]] }, ctx({ confirmDetailed: async () => ({ accepted: true }) }));
  assert.equal(accepted.ok, true, accepted.error);
  assert.equal(state.writes, 1);
  const rejected = await execute(rt, 'write_range', { sheet: 'S', start_cell: 'A1', values: [[2]] }, ctx({ confirmDetailed: async () => ({ accepted: false }) }));
  assert.equal(rejected.ok, false);
  assert.match(rejected.error, /向用户询问意图/);
  assert.equal(state.writes, 1);
});

test('ask_user_choice: a custom answer is passed to the model verbatim', async () => {
  const result = await execute(runtime(), 'ask_user_choice', { title: '方向', options: [{ label: '汇总' }, { label: '清洗' }] }, ctx({ choose: async () => ({ kind: 'custom', text: '按客户汇总，只看 2024 年' }) }));
  assert.equal(result.ok, true);
  assert.equal(result.data.custom, '按客户汇总，只看 2024 年');
  assert.match(result.data.hint, /不要再给选项/);
});

test('remove_duplicates rechecks confirmation data and validates selected columns before asking', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  let confirms = 0;
  for (const columns of [[], [10], [-1], [0.5]]) {
    const result = await execute(rt, 'remove_duplicates', { sheet: 'S', columns }, ctx({ confirm: async () => { confirms++; return true; } }));
    assert.equal(result.ok, false);
  }
  assert.equal(confirms, 0);
  const result = await execute(rt, 'remove_duplicates', { sheet: 'S', columns: ['名称'] }, ctx({ confirm: async () => { state.rows[1][1] = 'edited'; return true; } }));
  assert.equal(result.ok, false);
  assert.match(result.error, /发生了变化/);
  assert.equal(state.deduplicated, 0);
});

test('remove_duplicates works when the confirmed range is unchanged', async t => {
  const { state } = excelMock(t);
  const result = await execute(runtime(), 'remove_duplicates', { sheet: 'S', columns: ['名称'] });
  assert.equal(result.ok, true, result.error);
  assert.equal(state.deduplicated, 1);
});

test('delete_blank_rows preserves a previously empty row edited while confirmation is open', async t => {
  const { state } = excelMock(t, [['名称', '金额'], ['', ''], ['乙', 22]]);
  const result = await execute(runtime(), 'delete_blank_rows', { sheet: 'S' }, ctx({ confirm: async () => { state.rows[1][1] = 'new work'; return true; } }));
  assert.equal(result.ok, false);
  assert.match(result.error, /发生了变化/);
  assert.deepEqual(state.deleted, []);
});

test('delete_blank_rows removes unchanged empty rows from bottom to top', async t => {
  const { state } = excelMock(t, [['名称', '金额'], ['', ''], ['乙', 22], ['', '']]);
  const result = await execute(runtime(), 'delete_blank_rows', { sheet: 'S' });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(state.deleted, ['4:4', '2:2']);
});

test('disabled memory tools are rejected by the dispatcher itself', async () => {
  const result = await execute(runtime(), 'list_memories', {});
  assert.equal(result.ok, false);
  assert.match(result.error, /长期记忆已关闭/);
});

test('registry: every declared tool has exactly one implementation and vice versa', () => {
  const { TOOL_REGISTRY } = runtime();
  const { TOOL_DEFINITIONS, TOOL_LABELS } = sourceLoader()('src/core/tools/definitions.ts');
  const declared = TOOL_DEFINITIONS.map(t => t.function.name).sort();
  assert.equal(new Set(declared).size, declared.length, 'duplicate tool definitions');
  assert.deepEqual(Object.keys(TOOL_REGISTRY).sort(), declared);
  for (const name of declared) assert.ok(TOOL_LABELS[name], `missing label for ${name}`);
});

test('registry: tools that write the workbook are flagged so a checkpoint is taken first', () => {
  const { TOOL_REGISTRY } = runtime();
  const writes = /^(write|set|create|delete|clear|insert|format|sort|import|apply|copy|fill|find_replace|remove|rename|add|split|combine|clean|convert|polish|freeze|group|reconcile|stack)_/;
  for (const [name, spec] of Object.entries(TOOL_REGISTRY)) {
    if (writes.test(name)) assert.equal(spec.mutating, true, `${name} writes the workbook but is not mutating`);
  }
  assert.equal(TOOL_REGISTRY.read_range.mutating, undefined);
  assert.equal(TOOL_REGISTRY.remember.memory, true);
});

// ---------------- run_script / set_data_validation / set_visibility ----------------

/** Node 里没有 Web Worker：用同一个加载器拿到沙箱模块，换成直接执行的假执行器，只测工具管线。 */
function scriptRuntime() {
  const load = sourceLoader();
  const rt = load('src/core/tools/excelTools.ts');
  const sandbox = load('src/core/tools/scriptSandbox.ts');
  const jobs = [];
  sandbox.setScriptRunner(async job => {
    jobs.push(job);
    try {
      const value = await new Function('inputs', 'helpers', 'log', job.code)(job.inputs, {}, () => {});
      return { ok: true, value, logs: [], ms: 1 };
    } catch (e) {
      return { ok: false, error: `脚本运行出错: ${e.message}`, logs: [], ms: 1 };
    }
  });
  return { rt, jobs };
}

test('run_script reads full inputs and returns the value without touching the workbook', async t => {
  const { state } = excelMock(t);
  const { rt, jobs } = scriptRuntime();
  const result = await execute(rt, 'run_script', {
    code: 'return inputs.t.slice(1).reduce((s, r) => s + r[1], 0);',
    inputs: [{ name: 't', sheet: 'S' }],
  });
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.result, 33);
  assert.deepEqual(jobs[0].inputs.t, [['名称', '金额'], ['甲', 11], ['乙', 22]]);
  assert.equal(state.writes, 0);
  assert.equal(rt.isMutatingCall('run_script', { code: 'return 1' }), false);
  assert.equal(rt.isMutatingCall('run_script', { code: 'return 1', output: { new_sheet: 'X' } }), true);
});

test('run_script output to an existing range goes through the confirmed, undoable write path', async t => {
  const { state } = excelMock(t);
  const { rt } = scriptRuntime();
  const asked = [];
  const result = await execute(rt, 'run_script', {
    code: 'return [{ 名称: "甲", 翻倍: 22 }, { 名称: "乙", 翻倍: 44 }];',
    output: { sheet: 'S', start_cell: 'A1' },
  }, ctx({ confirm: async title => { asked.push(title); return true; } }));
  assert.equal(result.ok, true, result.error);
  assert.equal(asked.length, 1, 'overwriting existing cells must be confirmed');
  assert.deepEqual(state.rows, [['名称', '翻倍'], ['甲', 44 / 2], ['乙', 44]]);
  assert.equal(result.undoable, true);
});

test('run_script rejects unsafe source, bad inputs and non-tabular output before writing', async t => {
  const { state } = excelMock(t);
  const { rt, jobs } = scriptRuntime();
  for (const [args, pattern] of [
    [{ code: 'return import("https://x")' }, /import/],
    [{ code: 'return 1', inputs: [{ name: 'bad name', sheet: 'S' }] }, /变量名/],
    [{ code: 'return 1', output: {} }, /new_sheet/],
    [{ code: 'return 1', output: { sheet: 'S' } }, /非空数组/],
    [{ code: 'return [[{}]]', output: { sheet: 'S' } }, /第 1 行第 1 列/],
    [{ code: 'throw new Error("boom"); return 1' }, /boom/],
  ]) {
    const result = await execute(rt, 'run_script', args);
    assert.equal(result.ok, false, JSON.stringify(args));
    assert.match(result.error, pattern);
  }
  assert.equal(jobs.length, 3, 'source/argument errors never reach the sandbox');
  assert.equal(state.writes, 0);
});

test('set_data_validation builds list and range rules and refuses lists Excel cannot store', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  let result = await execute(rt, 'set_data_validation', { sheet: 'S', range: 'B2:B3', type: 'list', items: ['是', '否', '是'], error_message: '只能选是或否' });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(state.validation.rule, { list: { inCellDropDown: true, source: '是,否' } });
  assert.equal(state.validation.errorAlert.style, 'Stop');

  result = await execute(rt, 'set_data_validation', { sheet: 'S', range: 'B2:B3', type: 'whole_number', min: '1', max: 100 });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(state.validation.rule, { wholeNumber: { operator: 'Between', formula1: 1, formula2: 100 } });

  result = await execute(rt, 'set_data_validation', { sheet: 'S', range: 'A2:A3', type: 'date', operator: 'greater_or_equal', min: '2026-01-01' });
  assert.deepEqual(state.validation.rule, { date: { operator: 'GreaterThanOrEqualTo', formula1: '2026-01-01' } });

  const before = state.validation;
  for (const args of [
    { type: 'list', items: ['a,b'] },
    { type: 'list', items: Array.from({ length: 80 }, (_, i) => `选项${i}`) },
    { type: 'whole_number', min: 'abc', max: 3 },
    { type: 'date', min: '2026/1/1', operator: 'equal' },
    { type: 'custom', formula: 'COUNTIF(A:A,A2)=1' },
  ]) {
    const bad = await execute(rt, 'set_data_validation', { sheet: 'S', range: 'B2', ...args });
    assert.equal(bad.ok, false, JSON.stringify(args));
  }
  assert.equal(state.validation, before, 'rejected rules never clear the existing validation');
});

test('set_visibility hides whole rows/columns and format_range refuses merges that would drop data', async t => {
  const { state } = excelMock(t);
  const rt = runtime();
  assert.equal((await execute(rt, 'set_visibility', { target: 'rows', sheet: 'S', range: '3' })).ok, true);
  assert.equal((await execute(rt, 'set_visibility', { target: 'columns', sheet: 'S', range: 'c:d', hidden: false })).ok, true);
  assert.deepEqual(state.hidden, [['rows', '3:3', true], ['columns', 'C:D', false]]);
  assert.equal((await execute(rt, 'set_visibility', { target: 'rows', sheet: 'S', range: 'A1' })).ok, false);

  const merge = await execute(rt, 'format_range', { sheet: 'S', range: 'A1:B3', merge: true });
  assert.equal(merge.ok, false);
  assert.match(merge.error, /非空单元格/);
  assert.deepEqual(state.merged, []);
});

// ---------------- 让工具调用一次就对：公式自检、表名纠错、地址拆分 ----------------

test('format_range only merges when every cell after the anchor is empty, including formulas', async t => {
  const { state } = excelMock(t, [['标题', '']]);
  const rt = runtime();
  const merge = () => execute(rt, 'format_range', { sheet: 'S', range: 'A1:B1', merge: true });
  state.formulaRows = [['标题', '=""']];
  assert.equal((await merge()).ok, false);
  state.rows = state.formulaRows = [['', 0]];
  assert.equal((await merge()).ok, false);
  assert.deepEqual(state.merged, []);
  state.rows = [['', '']];
  state.formulaRows = [['=""', '']];
  assert.equal((await merge()).ok, true);
  assert.deepEqual(state.merged, ['A1:B1']);
});

test('set_formulas reports computed errors with addresses and hints so no read-back round is needed', async t => {
  const { state } = excelMock(t, [['', ''], ['', '']]);
  // 模拟 Excel 计算：查找列没匹配上，乘法列正常
  state.calc = grid => grid.map((row, r) => row.map((f, c) => (c === 0 ? '#N/A' : (r + 1) * 4)));
  const result = await execute(runtime(), 'set_formulas', { sheet: 'S', start_cell: 'C5', formulas: [['=XLOOKUP(A5,X:X,Y:Y)', '=B5*2'], ['=XLOOKUP(A6,X:X,Y:Y)', '=B6*2']] });
  assert.equal(result.ok, true, result.error);
  const check = result.data.formula_check;
  assert.equal(check.formula_cells, 4);
  assert.deepEqual(check.errors.map(e => [e.type, e.count, e.cells]), [['#N/A', 2, ['C5', 'C6']]]);
  assert.deepEqual(check.sample, [['#N/A', 4], ['#N/A', 8]]);
  assert.match(result.data.note, /2 个错误值/);

  const plain = await execute(runtime(), 'write_range', { sheet: 'S', start_cell: 'A1', values: [['x']] });
  assert.equal(plain.data.formula_check, undefined, 'value writes carry no formula check');
});

test('formula check counts error types with real addresses, blanks and a value sample', () => {
  const { checkFormulaResults, formulaCheckNote } = sourceLoader()('src/core/tools/formulaCheck.ts');
  const check = checkFormulaResults('Sheet1!$C$5:D7',
    [['#N/A', 4], ['', '#DIV/0!'], ['#N/A', 'text']],
    [['=XLOOKUP(A5,X:X,Y:Y)', '=B5*2'], ['=IF(A6,"",1)', '=1/0'], ['=XLOOKUP(A7,X:X,Y:Y)', 'text']]);
  assert.equal(check.formula_cells, 5);
  assert.equal(check.error_cells, 3);
  assert.equal(check.blank_results, 1);
  assert.deepEqual(check.errors.map(e => [e.type, e.count, e.cells]), [['#N/A', 2, ['C5', 'C7']], ['#DIV/0!', 1, ['D6']]]);
  assert.match(check.errors[0].hint, /类型|空格/);
  assert.match(formulaCheckNote(check), /3 个错误值/);
  const clean = checkFormulaResults('AA10', [[1, 2]], [['=1', '=2']]);
  assert.equal(clean.error_cells, 0);
  assert.match(formulaCheckNote(clean), /无需再读回/);
  assert.equal(formulaCheckNote(checkFormulaResults('A1', [[1]], [[1]])), '', 'plain values are not formula writes');
});

test('wrong sheet names come back with the existing sheets and the closest match', async t => {
  excelMock(t);
  const rt = runtime();
  const result = await execute(rt, 'read_range', { sheet: 's ', range: 'A1:B2' });
  assert.equal(result.ok, false);
  assert.match(result.error, /工作表「s 」不存在，你可能想用「S」/);
  assert.match(result.error, /现有工作表：「S」、「Other」/);

  const { closestSheetName, referencedSheets } = sourceLoader()('src/core/tools/argHints.ts');
  assert.equal(closestSheetName('１月销售', ['1月 销售', '2月 销售']), '1月 销售');
  assert.equal(closestSheetName('订单明细', ['订单明细表', '客户']), '订单明细表');
  assert.equal(closestSheetName('库存', ['订单', '客户']), null);
  assert.deepEqual(referencedSheets({ inputs: [{ sheet: 'A' }], output: { new_sheet: 'New', sheet: 'B' }, sheets: ['C'] }).sort(), ['A', 'B', 'C']);
});

test('"Sheet!A1" addresses are split into sheet + address only when the intent is unambiguous', async t => {
  const { splitSheetQualifiedAddresses } = sourceLoader()('src/core/tools/argHints.ts');
  const props = { sheet: {}, range: {}, source_sheet: {}, source_range: {} };
  let args = { range: "'1月 销售'!A1:C9" };
  assert.equal(splitSheetQualifiedAddresses(args, props), true);
  assert.deepEqual(args, { range: 'A1:C9', sheet: '1月 销售' });
  args = { sheet: 'S', range: 'S!B2' };
  splitSheetQualifiedAddresses(args, props);
  assert.deepEqual(args, { sheet: 'S', range: 'B2' });
  args = { sheet: 'S', range: 'Other!B2' };
  assert.equal(splitSheetQualifiedAddresses(args, props), false, 'conflicting sheet names are left for the tool to reject');
  args = { range: "O'Brien!A1" };
  assert.equal(splitSheetQualifiedAddresses(args, { range: {} }), false, 'tools without a sheet field are untouched');

  const { state } = excelMock(t);
  const result = await execute(runtime(), 'sort_range', { range: 'S!A1:B3', key_column: '金额' });
  assert.equal(result.ok, true, result.error);
  assert.ok(state.sort);
});
