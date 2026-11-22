const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const load = sourceLoader();
const reshape = load('src/core/tools/reshape.ts');
const { keepTextAsText } = load('src/core/tools/safeValues.ts');
const { DEFAULT_SETTINGS } = load('src/core/settings.ts');
const tools = load('src/core/tools/excelTools.ts');
const settings = { ...DEFAULT_SETTINGS, autoCheckpoint: false, useMemory: false };

const letter = i => { let s = ''; let n = i + 1; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const colIndex = s => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

// 按索引切片的工作簿模拟：足以核对读全表、写新表、数字格式和清理；不替代桌面 Excel 验收。
function workbook(t, data, formats = {}) {
  const previous = global.Excel;
  const state = { sheets: [], active: null, syncs: 0, onSync: null };
  const nullObject = () => ({ isNullObject: true, load() {} });
  function makeSheet(name, grid = [], fmt = []) {
    const sheet = { name, grid: grid.map(r => [...r]), formats: fmt.map(r => [...(r || [])]), frozen: 0, isNullObject: false, load() {} };
    const cell = (r, c) => sheet.grid[r]?.[c] ?? '';
    const rangeAt = (r, c, nr, nc) => ({
      isNullObject: false, rowIndex: r, columnIndex: c, rowCount: nr, columnCount: nc,
      address: `${sheet.name}!${letter(c)}${r + 1}:${letter(c + nc - 1)}${r + nr}`,
      load() {},
      get values() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => cell(r + i, c + j))); },
      set values(g) { g.forEach((row, i) => row.forEach((v, j) => { (sheet.grid[r + i] ||= [])[c + j] = v; })); },
      get numberFormat() { return Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => sheet.formats[r + i]?.[c + j] ?? 'General')); },
      set numberFormat(g) { g.forEach((row, i) => row.forEach((v, j) => { (sheet.formats[r + i] ||= [])[c + j] = v; })); },
      getIntersectionOrNullObject(o) {
        const r0 = Math.max(r, o.rowIndex); const c0 = Math.max(c, o.columnIndex);
        const r1 = Math.min(r + nr, o.rowIndex + o.rowCount); const c1 = Math.min(c + nc, o.columnIndex + o.columnCount);
        return r1 > r0 && c1 > c0 ? rangeAt(r0, c0, r1 - r0, c1 - c0) : nullObject();
      },
      format: { font: {}, autofitColumns() {} },
    });
    const used = () => {
      const rows = sheet.grid.length;
      const cols = Math.max(0, ...sheet.grid.map(row => row.length));
      return rows && cols ? rangeAt(0, 0, rows, cols) : null;
    };
    sheet.getRangeByIndexes = rangeAt;
    sheet.getUsedRangeOrNullObject = () => used() || nullObject();
    sheet.getUsedRange = () => used() || rangeAt(0, 0, 1, 1);
    sheet.getRange = addr => {
      const m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(addr);
      const c0 = colIndex(m[1]); const c1 = colIndex(m[3] || m[1]);
      const r0 = m[2] ? Number(m[2]) - 1 : 0;
      const r1 = m[4] ? Number(m[4]) - 1 : m[2] ? r0 : 1048575;
      return rangeAt(r0, c0, r1 - r0 + 1, c1 - c0 + 1);
    };
    sheet.freezePanes = { freezeRows(n) { sheet.frozen = n; } };
    sheet.activate = () => { state.active = sheet.name; };
    sheet.delete = () => { state.sheets.splice(state.sheets.indexOf(sheet), 1); };
    return sheet;
  }
  for (const [name, grid] of Object.entries(data)) state.sheets.push(makeSheet(name, grid, formats[name] || []));
  const worksheets = {
    get items() { return state.sheets; },
    load() {},
    getItem(name) { const s = state.sheets.find(x => x.name === name); if (!s) throw new Error(`ItemNotFound ${name}`); return s; },
    getItemOrNullObject(name) { return state.sheets.find(x => x.name === name) || nullObject(); },
    add(name) { if (state.sheets.some(x => x.name === name)) throw new Error(`duplicate ${name}`); const s = makeSheet(name); state.sheets.push(s); return s; },
    getActiveWorksheet() { return state.sheets.find(s => s.name === state.active) || state.sheets[0]; },
  };
  const context = { workbook: { worksheets }, sync: async () => { state.syncs++; await state.onSync?.(state); } };
  global.Excel = { run: fn => fn(context) };
  t.after(() => { global.Excel = previous; });
  return { state, sheet: name => state.sheets.find(s => s.name === name) };
}

function ctx(overrides = {}) { return { settings, confirm: async () => true, choose: async () => ({ kind: 'skipped' }), ...overrides }; }
const execute = (name, args, context = ctx()) => tools.executeTool(name, JSON.stringify(args), context);

test('keepTextAsText quote-prefixes only strings Excel would coerce', () => {
  for (const s of ['00123', '2024-03', '1/2', '50%', '1,200', '3月12日', 'TRUE', '=1+1', "'x"]) assert.equal(keepTextAsText(s), `'${s}`);
  for (const v of ['华东', 'A-001', '2024-Q1', '2024-W09', '', 123, null]) assert.equal(keepTextAsText(v), v);
});

test('dateBucket maps Excel serials and date text to year / quarter / month / ISO week / day', () => {
  const march8 = 45359;
  assert.deepEqual(reshape.dateBucket(march8, 'month'), { key: '2024-03', value: 45352 });
  assert.deepEqual(reshape.dateBucket('2024/3/8', 'month'), { key: '2024-03', value: 45352 });
  assert.deepEqual(reshape.dateBucket(march8, 'quarter'), { key: '2024-Q1', value: '2024-Q1' });
  assert.deepEqual(reshape.dateBucket(march8, 'year'), { key: '2024', value: 2024 });
  assert.deepEqual(reshape.dateBucket(march8, 'day'), { key: '2024-03-08', value: march8 });
  assert.equal(reshape.dateBucket('2024-03-08', 'week').key, '2024-W10');
  assert.equal(reshape.dateBucket('2021-01-03', 'week').key, '2020-W53');
  assert.equal(reshape.dateBucket('2024-12-30', 'week').key, '2025-W01');
  assert.equal(reshape.dateBucket('不是日期', 'month'), null);
});

test('groupSummary covers every row, labels blanks and bad dates, and reports skipped text numbers', () => {
  const rows = [
    ['East', '2024-01-05', 100, 'A'],
    ['East ', 45300, 50, 'B'],
    ['North', '2024-02-01', '¥30', 'A'],
    ['', '2024-02-10', 20, 'C'],
    ['North', 'bad', 10, 'A'],
    ['', '', '', ''],
  ];
  const groups = [{ index: 1, header: '日期', datePart: 'month' }, { index: 0, header: '地区' }];
  const metrics = [{ index: 2, agg: 'sum', header: '金额_合计' }, { index: null, agg: 'count_rows', header: '行数' }, { index: 3, agg: 'count_distinct', header: '客户_去重计数' }];
  const r = reshape.groupSummary(rows, groups, metrics);
  assert.deepEqual(r.grid, [
    ['日期(月)', '地区', '金额_合计', '行数', '客户_去重计数'],
    [45292, 'East', 150, 2, 2],
    [45323, 'North', 0, 1, 1],
    [45323, '(空白)', 20, 1, 1],
    ['(不是日期)', 'North', 10, 1, 1],
    ['合计', '', 180, 5, 3],
  ]);
  assert.deepEqual([r.rowsUsed, r.blankRowsSkipped, r.blankGroupRows, r.unparsedDates, r.nonNumeric], [5, 1, 1, 1, [1, 0, 0]]);

  const parsed = reshape.groupSummary(rows, groups.slice(1), metrics.slice(0, 1), { parseTextNumbers: true, sort: 'desc' });
  assert.deepEqual(parsed.grid.slice(1), [['East', 150], ['North', 40], ['(空白)', 20], ['合计', 210]]);
  assert.deepEqual(parsed.nonNumeric, [0]);
});

test('groupSummary average / min / max ignore blanks and tidy float noise', () => {
  const rows = [['a', 0.1], ['a', 0.2], ['a', ''], ['b', 5]];
  const r = reshape.groupSummary(rows, [{ index: 0, header: 'k' }], [
    { index: 1, agg: 'sum', header: 's' }, { index: 1, agg: 'average', header: 'avg' },
    { index: 1, agg: 'min', header: 'min' }, { index: 1, agg: 'max', header: 'max' }, { index: 1, agg: 'count', header: 'n' },
  ], { totalRow: false });
  assert.deepEqual(r.grid.slice(1), [['a', 0.3, 0.15, 0.1, 0.2, 2], ['b', 5, 5, 5, 5, 1]]);
  assert.deepEqual(r.blanks, [1, 1, 1, 1, 1]);
});

test('reconcileTables matches composite keys across spacing / full-width, sums duplicates and honours tolerance', () => {
  const left = { label: '订单', keys: [0, 1], compare: [2, 3], rows: [
    ['A001', 1, 100, '已付'], ['A001', 2, 50, '已付'], ['A002', 1, 200, '未付'], ['Ａ003', 1, 80, '已付'],
    ['A004', 1, 10, '已付'], ['A004', 1, 5, '已付'], ['', '', 99, 'x'],
  ] };
  const right = { label: '流水', keys: [0, 1], compare: [2, 3], rows: [
    ['A001', 1, 100.004, '已付'], ['A001 ', 2, 49, '已付'], ['A003', 1, 80, '已付'], ['A004', 1, '15', '已付'], ['A005', 1, 70, '已付'],
  ] };
  const r = reshape.reconcileTables(left, right, { keyHeaders: ['订单号', '行号'], compareHeaders: ['金额', '状态'], tolerance: 0.01 });
  assert.deepEqual(r.grid, [
    ['状态', '差异列', '订单号', '行号', '金额_订单', '金额_流水', '金额_差额', '状态_订单', '状态_流水', '订单行数', '流水行数'],
    ['不一致', '金额', 'A001', 2, 50, 49, 1, '已付', '已付', 1, 1],
    ['仅左表', '', 'A002', 1, 200, '', 200, '未付', '', 1, 0],
    ['仅右表', '', 'A005', 1, '', 70, -70, '', '已付', 0, 1],
  ]);
  assert.deepEqual([r.matched, r.mismatched, r.onlyLeft, r.onlyRight], [3, 1, 1, 1]);
  assert.deepEqual([r.leftRows, r.leftBlankKeyRows, r.leftDuplicateKeys, r.rightDuplicateKeys], [7, 1, 1, 0]);
  assert.deepEqual(r.columns[0], { header: '金额', numeric: true, mismatched: 1, left_total: 445, right_total: 314.004, difference: 130.996 });
  assert.equal(r.columns[1].numeric, false);

  const strict = reshape.reconcileTables(left, right, { keyHeaders: ['订单号', '行号'], compareHeaders: ['金额', '状态'], includeMatched: true });
  assert.equal(strict.mismatched, 2, 'without tolerance 100 vs 100.004 differs');
  assert.equal(strict.grid.filter(row => row[0] === '一致').length, 2);
});

test('stackTables aligns headers by name, keeps first formats, suffixes duplicate headers and reports gaps', () => {
  const sources = [
    { label: '1月', headers: ['日期', '门店', '金额'], rows: [[45292, 'A', 10], [45293, 'B', 20]], formats: ['yyyy-mm-dd', '', '#,##0.00'] },
    { label: '2月', headers: ['门店', '金额 ', '备注'], rows: [['A', 30, 'x'], ['', '', '']], formats: ['', '0.00', ''] },
    { label: '3月', headers: ['金额', '金额', '门店'], rows: [[1, 2, 'C']] },
  ];
  const r = reshape.stackTables(sources);
  assert.deepEqual(r.grid, [
    ['来源', '日期', '门店', '金额', '备注', '金额_2'],
    ['1月', 45292, 'A', 10, '', ''],
    ['1月', 45293, 'B', 20, '', ''],
    ['2月', '', 'A', 30, 'x', ''],
    ['3月', '', 'C', 1, '', 2],
  ]);
  assert.deepEqual(r.formats, [undefined, 'yyyy-mm-dd', undefined, '#,##0.00', undefined, undefined]);
  assert.deepEqual(r.sources.map(s => [s.label, s.rows, s.missing]), [['1月', 2, ['备注', '金额_2']], ['2月', 1, ['日期', '金额_2']], ['3月', 1, ['日期', '备注']]]);

  const picked = reshape.stackTables(sources, { sourceHeader: null, columns: ['门店', '金额', '不存在'] });
  assert.deepEqual(picked.grid[0], ['门店', '金额', '不存在']);
  assert.deepEqual(picked.unknownColumns, ['不存在']);
  assert.deepEqual(picked.sources[0].dropped, ['日期']);

  const clash = reshape.stackTables([{ label: 'x', headers: ['来源', 'v'], rows: [['a', 1]] }, { label: 'y', headers: ['v'], rows: [[2]] }]);
  assert.deepEqual(clash.grid[0], ['来源_Sheet', '来源', 'v']);
});

test('splitByKey groups trimmed values, sorts numbers first and puts blanks last', () => {
  const groups = reshape.splitByKey([['East', 1], ['North', 2], ['', 3], ['East ', 4], [10, 5], [2, 6], ['', '']], 0);
  assert.deepEqual(groups.map(g => [g.label, g.rows.length]), [['2', 1], ['10', 1], ['East', 2], ['North', 1], ['(空白)', 1]]);
});

const SALES = [['地区', '下单日期', '金额', '客户编号'], ['East', 45292, 100, '00123'], ['North', 45323, 50, '00456'], ['East', 45300, 'n/a', '00789']];
const SALES_FORMATS = [[], ['General', 'yyyy-mm-dd', '#,##0.00', '@'], ['General', 'yyyy-mm-dd', '#,##0.00', '@']];

test('group_summary writes a new sheet with number formats, header freeze and warnings; source untouched', async t => {
  const { state, sheet } = workbook(t, { 销售: SALES }, { 销售: SALES_FORMATS });
  const before = JSON.stringify(sheet('销售').grid);
  const result = await execute('group_summary', { sheet: '销售', range: 'A:D', group_by: ['地区'], metrics: [{ column: '金额', agg: 'sum' }, { agg: 'count_rows' }] });
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.dest_sheet, '销售_汇总');
  const out = sheet('销售_汇总');
  assert.deepEqual(out.grid, [['地区', '金额_合计', '行数'], ['East', 100, 2], ['North', 50, 1], ['合计', 150, 3]]);
  assert.equal(out.formats[0]?.[1], undefined, 'header row keeps its format');
  assert.deepEqual(out.formats.slice(1).map(r => r[1]), ['#,##0.00', '#,##0.00', '#,##0.00']);
  assert.equal(out.frozen, 1);
  assert.equal(state.active, '销售_汇总');
  assert.match(result.data.warnings.join(), /1 个非空单元格不是数字/);
  assert.equal(JSON.stringify(sheet('销售').grid), before);
});

test('group_summary by month writes date serials formatted yyyy-mm and never overwrites an existing sheet', async t => {
  const { sheet } = workbook(t, { 销售: SALES, 销售_汇总: [['old']] }, { 销售: SALES_FORMATS });
  const result = await execute('group_summary', { sheet: '销售', date_column: '下单日期', metrics: [{ column: 'C', agg: 'sum', header: '销售额' }] });
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.dest_sheet, '销售_汇总(2)');
  assert.deepEqual(sheet('销售_汇总').grid, [['old']]);
  const out = sheet('销售_汇总(2)');
  assert.deepEqual(out.grid, [['下单日期(月)', '销售额'], [45292, 100], [45323, 50], ['合计', 150]]);
  assert.deepEqual(out.formats.slice(1).map(r => r[0]), ['yyyy-mm', 'yyyy-mm', 'yyyy-mm']);
});

test('reshape tools reject missing sheets, bad columns and empty specs before writing', async t => {
  const { state } = workbook(t, { 销售: SALES });
  const count = state.sheets.length;
  assert.match((await execute('group_summary', { sheet: '不存在', group_by: ['地区'] })).error, /找不到工作表「不存在」/);
  assert.match((await execute('group_summary', { sheet: '销售', group_by: ['区域'] })).error, /找不到group_by\[0\]「区域」/);
  assert.match((await execute('group_summary', { sheet: '销售' })).error, /group_by/);
  assert.match((await execute('reconcile_tables', { left_sheet: '销售', right_sheet: '销售', keys: [] })).error, /至少需要 1 项/);
  assert.equal(state.sheets.length, count);
});

test('split_by_column keeps text codes as text, applies formats and needs confirmation above ten sheets', async t => {
  const { sheet } = workbook(t, { 销售: SALES, East: [['taken']] }, { 销售: SALES_FORMATS });
  const result = await execute('split_by_column', { sheet: '销售', column: '地区' });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(result.data.sheets, [{ sheet: 'East(2)', value: 'East', rows: 2 }, { sheet: 'North', value: 'North', rows: 1 }]);
  assert.deepEqual(sheet('East(2)').grid, [['地区', '下单日期', '金额', '客户编号'], ['East', 45292, 100, "'00123"], ['East', 45300, 'n/a', "'00789"]]);
  assert.equal(sheet('North').formats[1][1], 'yyyy-mm-dd');
  assert.deepEqual(sheet('East').grid, [['taken']]);

  const many = [['k'], ...Array.from({ length: 12 }, (_, i) => [`v${i}`])];
  workbook(t, { 多: many });
  let asked = 0;
  const rejected = await execute('split_by_column', { sheet: '多', column: 'k' }, ctx({ confirm: async () => { asked++; return false; } }));
  assert.equal(asked, 1);
  assert.equal(rejected.ok, false);
  const limited = await execute('split_by_column', { sheet: '多', column: 'k', max_sheets: 5 });
  assert.match(limited.error, /12 个不同取值，超过 max_sheets=5/);
});

test('reconcile_tables writes one filterable sheet, or none when the tables agree', async t => {
  const { state, sheet } = workbook(t, {
    订单: [['订单号', '金额'], ['A1', 10], ['A2', 20]],
    流水: [['单号', '实收'], ['A1', 10], ['A3', 5]],
    副本: [['订单号', '金额'], ['A2', 20], ['A1', 10]],
  }, { 订单: [[], ['General', '0.00']] });
  const result = await execute('reconcile_tables', { left_sheet: '订单', right_sheet: '流水', keys: ['订单号'], right_keys: ['单号'], compare_columns: ['金额'], right_compare_columns: ['实收'] });
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(sheet('对账结果').grid, [
    ['状态', '差异列', '订单号', '金额_订单', '金额_流水', '金额_差额'],
    ['仅左表', '', 'A2', 20, '', 20],
    ['仅右表', '', 'A3', '', 5, -5],
  ]);
  assert.equal(sheet('对账结果').formats[1][3], '0.00');
  assert.deepEqual([result.data.matched, result.data.only_left, result.data.only_right], [1, 1, 1]);

  const count = state.sheets.length;
  const same = await execute('reconcile_tables', { left_sheet: '订单', right_sheet: '副本', keys: ['订单号'], compare_columns: ['金额'] });
  assert.equal(same.ok, true, same.error);
  assert.equal(same.data.dest_sheet, null);
  assert.equal(state.sheets.length, count);
});

test('stack_sheets reads every row in blocks and removes the half-written sheet when stopped', async t => {
  const big = n => [['id', 'v'], ...Array.from({ length: n }, (_, i) => [i, i * 2])];
  const { state, sheet } = workbook(t, { A: big(12000), B: big(9000) });
  const ok = await execute('stack_sheets', { sheets: ['A', 'B'] });
  assert.equal(ok.ok, true, ok.error);
  assert.equal(ok.data.rows, 21000);
  assert.deepEqual(sheet('合并结果').grid[12001], ['B', 0, 0]);

  const controller = new AbortController();
  state.onSync = s => { const out = s.sheets.find(x => x.name === '合并结果(2)'); if (out && out.grid.length) controller.abort(); };
  const stopped = await execute('stack_sheets', { sheets: ['A', 'B'] }, ctx({ signal: controller.signal }));
  assert.equal(stopped.ok, false);
  assert.match(stopped.error, /停止/);
  assert.equal(sheet('合并结果(2)'), undefined);
});
