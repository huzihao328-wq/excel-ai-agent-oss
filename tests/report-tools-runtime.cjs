const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const load = sourceLoader();
const tools = load('src/core/tools/excelTools.ts');
const clean = load('src/core/tools/cleanValues.ts');
const extract = load('src/core/tools/textExtract.ts');
const fill = load('src/core/tools/fillBlanks.ts');
const reshape = load('src/core/tools/reshape.ts');
const layout = load('src/core/tools/layoutTools.ts');
const { DEFAULT_SETTINGS } = load('src/core/settings.ts');
const settings = { ...DEFAULT_SETTINGS, autoCheckpoint: false, useMemory: false };

const letter = i => { let s = ''; let n = i + 1; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const colIndex = s => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const grid2 = (nr, nc, at) => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => at(i, j)));

/**
 * 按索引寻址的工作簿模拟：区分值与公式、数字格式、合并区域、超链接和页面布局，
 * 写入时像 Excel 一样把开头的单引号当作文本前缀去掉。不替代桌面 Excel 验收。
 */
function book(t, data, options = {}) {
  const previous = { Excel: global.Excel, Office: global.Office };
  const state = { sheets: [], active: null, unmerged: [], pendingError: null, confirms: [] };
  const nullObject = () => ({ isNullObject: true, load() {} });
  const overlaps = (m, r, c, nr, nc) => m.row < r + nr && r < m.row + m.rows && m.column < c + nc && c < m.column + m.columns;
  function makeLayout() {
    const pl = {
      orientation: 'Portrait', paperSize: 'Letter', centerHorizontally: false, printGridlines: false, titleRows: null, titleColumns: null, printArea: null,
      topMargin: 0, bottomMargin: 0, leftMargin: 0, rightMargin: 0, headerMargin: 0, footerMargin: 0, _zoom: { scale: 100 },
      get zoom() { return pl._zoom; },
      set zoom(v) { if (options.rejectZeroHeight && v.verticalFitToPages === 0) state.pendingError = new Error('InvalidArgument'); else pl._zoom = v; },
      load() {},
      setPrintTitleRows(v) { pl.titleRows = v; },
      setPrintTitleColumns(v) { pl.titleColumns = v; },
      setPrintArea(v) { pl.printArea = v; },
      headersFooters: { defaultForAllPages: { centerHeader: '', centerFooter: '', leftFooter: '' } },
    };
    return pl;
  }
  function makeSheet(name, grid = []) {
    const sheet = {
      name, visibility: 'Visible', position: null, isNullObject: false, frozen: 0,
      grid: grid.map(r => [...r]), f: [], fmt: (options.formats?.[name] || []).map(r => [...(r || [])]),
      merges: (options.merges?.[name] || []).map(m => ({ ...m })), fills: {}, links: {}, pageLayout: makeLayout(),
      load() {},
    };
    const get = (r, c) => sheet.grid[r]?.[c] ?? '';
    const write = (r, c, v, asFormula) => {
      (sheet.f[r] ||= []);
      (sheet.grid[r] ||= []);
      if (typeof v === 'string' && v.startsWith("'")) { sheet.grid[r][c] = v.slice(1); sheet.f[r][c] = undefined; }
      else if (asFormula && typeof v === 'string' && v.startsWith('=')) { sheet.grid[r][c] = v; sheet.f[r][c] = v; }
      else { sheet.grid[r][c] = v ?? ''; sheet.f[r][c] = undefined; }
    };
    const rangeAt = (r, c, nr, nc) => {
      const range = {
        worksheet: sheet,
        isNullObject: false, rowIndex: r, columnIndex: c, rowCount: nr, columnCount: nc,
        address: `${sheet.name}!${letter(c)}${r + 1}:${letter(c + nc - 1)}${r + nr}`,
        load() {},
        get values() { return grid2(nr, nc, (i, j) => get(r + i, c + j)); },
        set values(g) { g.forEach((row, i) => row.forEach((v, j) => write(r + i, c + j, v, false))); },
        get formulas() { return grid2(nr, nc, (i, j) => sheet.f[r + i]?.[c + j] ?? get(r + i, c + j)); },
        set formulas(g) { g.forEach((row, i) => row.forEach((v, j) => write(r + i, c + j, v, true))); },
        get numberFormat() { return grid2(nr, nc, (i, j) => sheet.fmt[r + i]?.[c + j] ?? 'General'); },
        set numberFormat(g) { g.forEach((row, i) => row.forEach((v, j) => { (sheet.fmt[r + i] ||= [])[c + j] = v; })); },
        getCell(i, j) { return rangeAt(r + i, c + j, 1, 1); },
        getAbsoluteResizedRange(rows, cols) { return rangeAt(r, c, rows, cols); },
        getIntersectionOrNullObject(o) {
          const r0 = Math.max(r, o.rowIndex); const c0 = Math.max(c, o.columnIndex);
          const r1 = Math.min(r + nr, o.rowIndex + o.rowCount); const c1 = Math.min(c + nc, o.columnIndex + o.columnCount);
          return r1 > r0 && c1 > c0 ? rangeAt(r0, c0, r1 - r0, c1 - c0) : nullObject();
        },
        getMergedAreasOrNullObject() {
          const items = sheet.merges.filter(m => overlaps(m, r, c, nr, nc))
            .map(m => ({ rowIndex: m.row, columnIndex: m.column, rowCount: m.rows, columnCount: m.columns }));
          return { isNullObject: items.length === 0, load() {}, areas: { items, load() {} } };
        },
        unmerge() {
          state.unmerged.push(range.address.split('!')[1]);
          sheet.merges = sheet.merges.filter(m => !overlaps(m, r, c, nr, nc));
        },
        set hyperlink(h) { sheet.links[`${letter(c)}${r + 1}`] = h; write(r, c, h.textToDisplay, false); },
        clear() {
          for (let i = r; i < Math.min(r + nr, sheet.grid.length); i++) {
            for (let j = c; j < Math.min(c + nc, sheet.grid[i].length); j++) {
              write(i, j, '', false);
              const address = letter(j) + (i + 1);
              delete sheet.links[address];
              delete sheet.fills[address];
              if (sheet.fmt[i]) delete sheet.fmt[i][j];
            }
          }
        },
        format: {
          font: {}, autofitColumns() {},
          fill: { set color(v) { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) sheet.fills[`${letter(c + j)}${r + i + 1}`] = v; } },
        },
      };
      return range;
    };
    const used = () => {
      const rows = sheet.grid.length;
      const cols = Math.max(0, ...sheet.grid.map(row => row?.length || 0));
      return rows && cols ? rangeAt(0, 0, rows, cols) : null;
    };
    sheet.getRangeByIndexes = rangeAt;
    sheet.getUsedRangeOrNullObject = () => used() || nullObject();
    sheet.getUsedRange = () => used() || rangeAt(0, 0, 1, 1);
    sheet.getRange = addr => {
      if (addr == null) return rangeAt(0, 0, 1048576, 16384);
      const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(addr.replace(/^.*!/, ''));
      const r0 = Number(m[2]) - 1; const c0 = colIndex(m[1]);
      return rangeAt(r0, c0, (m[4] ? Number(m[4]) - 1 : r0) - r0 + 1, (m[3] ? colIndex(m[3]) : c0) - c0 + 1);
    };
    sheet.freezePanes = { freezeRows(n) { sheet.frozen = n; } };
    const names = new Map();
    sheet.names = {
      getItemOrNullObject(key) { return names.get(key) ?? nullObject(); },
      add(key, range, comment) {
        assert.ok(!names.has(key));
        const item = {
          comment, type: 'Range', isNullObject: false, load() {},
          getRange() { return range; },
          set formula(reference) { range = sheet.getRange(reference); },
        };
        names.set(key, item);
        return item;
      },
    };
    sheet.activate = () => { state.active = sheet.name; };
    sheet.delete = () => { state.sheets.splice(state.sheets.indexOf(sheet), 1); };
    return sheet;
  }
  for (const [name, grid] of Object.entries(data)) state.sheets.push(makeSheet(name, grid));
  for (const [name, visibility] of Object.entries(options.visibility || {})) state.sheets.find(s => s.name === name).visibility = visibility;
  const worksheets = {
    get items() { return state.sheets; },
    load() {},
    getItem(name) { const s = state.sheets.find(x => x.name === name); if (!s) throw new Error(`ItemNotFound ${name}`); return s; },
    getItemOrNullObject(name) { return state.sheets.find(x => x.name === name) || nullObject(); },
    add(name) { if (state.sheets.some(x => x.name === name)) throw new Error(`duplicate ${name}`); const s = makeSheet(name); state.sheets.push(s); return s; },
    getActiveWorksheet() { return state.sheets.find(s => s.name === state.active) || state.sheets[0]; },
  };
  const context = {
    workbook: { worksheets },
    sync: async () => { if (state.pendingError) { const e = state.pendingError; state.pendingError = null; throw e; } },
  };
  global.Excel = { run: fn => fn(context) };
  global.Office = options.mergedApi === false ? undefined : { context: { requirements: { isSetSupported: () => true } } };
  t.after(() => { global.Excel = previous.Excel; global.Office = previous.Office; });
  return { state, sheet: name => state.sheets.find(s => s.name === name) };
}

function ctx(state, overrides = {}) {
  return { settings, confirm: async (title, detail) => { state?.confirms.push({ title, detail }); return true; }, choose: async () => ({ kind: 'skipped' }), ...overrides };
}
const execute = (name, args, context) => tools.executeTool(name, JSON.stringify(args), context);
const column = (sheet, c) => sheet.grid.map(row => row?.[c] ?? '');

// ---------------- 纯函数 ----------------

test('toRmbUpper writes cheque-style amounts, including zeros across 万/亿 sections', () => {
  const cases = {
    1234.56: '壹仟贰佰叁拾肆元伍角陆分',
    1001: '壹仟零壹元整',
    10: '壹拾元整',
    100000: '壹拾万元整',
    1000100: '壹佰万零壹佰元整',
    100010000: '壹亿零壹万元整',
    20000000.07: '贰仟万元零柒分',
    10.05: '壹拾元零伍分',
    0.5: '伍角',
    0: '零元整',
    '-20.3': '负贰拾元叁角',
  };
  for (const [n, text] of Object.entries(cases)) assert.equal(clean.toRmbUpper(Number(n)), text, n);
  assert.equal(clean.toRmbUpper(Infinity), null);
});

test('new clean ops: full-width, all spaces, case and RMB upper, with aliases', () => {
  assert.deepEqual(clean.normalizeCleanOps(['halfwidth', 'no_spaces', 'RMB', 'to-upper']), ['to_halfwidth', 'remove_spaces', 'rmb_upper', 'upper']);
  assert.throws(() => clean.normalizeCleanOps(['pinyin']), /to_halfwidth \/ remove_spaces/);
  assert.equal(clean.applyCleanOps('１２３', ['to_halfwidth', 'to_number']).value, 123);
  assert.equal(clean.applyCleanOps('Ａｂｃ－１', ['to_halfwidth', 'upper']).value, 'ABC-1');
  assert.equal(clean.applyCleanOps(' 张 三　', ['remove_spaces']).value, '张三');
  assert.equal(clean.applyCleanOps('¥1,234.56', ['rmb_upper']).value, '壹仟贰佰叁拾肆元伍角陆分');
  assert.deepEqual(clean.applyCleanOps('待定', ['rmb_upper']), { value: '待定', changed: false, unparsed: true });
  assert.equal(clean.applyCleanOps(42, ['upper']).value, 42);
});

test('Chinese ID cards: checksum, 15-digit legacy numbers, gender, birthday and age', () => {
  const today = new Date(2026, 9, 23);
  const male = '44030619900307123' + extract.idCheckDigit('44030619900307123');
  assert.equal(extract.idCheckDigit('11010519491231002'), 'X');
  assert.deepEqual(extract.parseChineseId('11010519491231002x', today), { id: '11010519491231002X', birth: '1949-12-31', gender: '女', check: '正确', valid: true });
  assert.equal(extract.parseChineseId(male, today).gender, '男');
  const wrong = male.slice(0, 17) + (male[17] === '0' ? '1' : '0');
  assert.equal(extract.parseChineseId(wrong, today).check, '校验位错误');
  assert.equal(extract.parseChineseId('110105199013451234', today).check, '出生日期无效');
  assert.deepEqual(extract.parseChineseId('110105491231002', today), { id: '110105491231002', birth: '1949-12-31', gender: '女', check: '正确', valid: true });
  assert.equal(extract.parseChineseId('12345', today), null);
  assert.equal(extract.ageOn('1949-12-31', today), 76);
  assert.equal(extract.ageOn('1990-10-23', today), 36);
  assert.equal(extract.ageOn('1990-10-24', today), 35);
});

test('extractValue pulls phones, emails, ids, numbers, dates, urls and text runs out of free text', () => {
  const today = new Date(2026, 9, 23);
  const x = (v, target, opts = {}) => extract.extractValue(v, target, { today, ...opts }).value;
  const note = '电话：+86 138-1234-5678，备用13912345678';
  assert.equal(x(note, 'mobile'), '13812345678');
  assert.equal(x(note, 'mobile', { all: true }), '13812345678、13912345678');
  assert.equal(x('11010519491231002X', 'mobile'), '');
  assert.equal(extract.extractValue('无', 'mobile', { today }).missing, true);
  assert.equal(extract.extractValue('', 'mobile', { today }).missing, false);
  assert.equal(x('联系 Zhang.San@example.com.cn 谢谢', 'email'), 'Zhang.San@example.com.cn');
  assert.equal(x('身份证：11010519491231002X，张三', 'id_card'), '11010519491231002X');
  assert.equal(x('身份证：１１０１０５１９４９１２３１００２Ｘ', 'gender'), '女');
  assert.equal(x('11010519491231002X', 'age'), 76);
  assert.equal(x('合计¥1,234.50元', 'number'), 1234.5);
  assert.equal(x('型号X-200', 'number'), 200);
  assert.equal(x('温度 -5 度', 'number'), -5);
  assert.equal(x('增长12%', 'number'), 0.12);
  assert.equal(x(42, 'number'), 42);
  assert.equal(x('2024年3月5日下单', 'date'), '2024-03-05');
  assert.equal(x('发货 2024/03/09 18:00', 'date'), '2024-03-09');
  assert.equal(x(45359, 'date'), '2024-03-08');
  assert.equal(x('张三Tom 138', 'chinese'), '张三');
  assert.equal(x('张三Tom 138', 'english'), 'Tom');
  assert.equal(x('见 https://example.com/a?b=1，谢谢', 'url'), 'https://example.com/a?b=1');
  assert.deepEqual(extract.normalizeExtractTargets(['手机号', 'birthday', 'sex', 'phone']), ['mobile', 'birth_date', 'gender']);
  assert.throws(() => extract.normalizeExtractTargets(['regex']), /run_script/);
});

test('an 18-digit ID stored as a number keeps its birthday but not gender or checksum', () => {
  const today = new Date(2026, 9, 23);
  const stored = 440306199003071230;
  assert.equal(extract.extractValue(stored, 'birth_date', { today }).value, '1990-03-07');
  assert.equal(extract.extractValue(stored, 'age', { today }).value, 36);
  assert.deepEqual(extract.extractValue(stored, 'gender', { today }), { value: '', missing: false, precisionLost: true });
  assert.equal(extract.extractValue(stored, 'id_check', { today }).value, '存成数字，后几位已丢失');
});

test('planBlankFill fills vertical merges and chosen columns, skipping header merges and blank rows', () => {
  const values = [
    ['部门', '姓名', '金额'],
    ['销售', '甲', 10],
    ['', '乙', ''],
    ['', '丙', 5],
    ['', '', ''],
    ['财务', '丁', ''],
    ['', '戊', 3],
  ];
  const merged = [{ row: 1, column: 0, rows: 3, columns: 1 }, { row: 0, column: 0, rows: 1, columns: 3 }];
  const onlyMerges = fill.planBlankFill(values, { columns: [], mode: 'above', startRow: 1, merged });
  assert.deepEqual(onlyMerges.cells.map(c => [c.row, c.column, c.value]), [[2, 0, '销售'], [3, 0, '销售']]);
  assert.equal(onlyMerges.mergedCells, 2);

  const above = fill.planBlankFill(values, { columns: [0], mode: 'above', startRow: 1, merged });
  assert.deepEqual(above.cells.map(c => [c.row, c.column, c.value]), [[2, 0, '销售'], [3, 0, '销售'], [6, 0, '财务']]);
  assert.deepEqual(above.perColumn, { 0: 1 });
  assert.deepEqual(above.cells[2].from, { row: 5, column: 0 });

  const zero = fill.planBlankFill(values, { columns: [2], mode: 'value', value: 0, startRow: 1 });
  assert.deepEqual(zero.cells.map(c => [c.row, c.column, c.value, c.from]), [[2, 2, 0, null], [5, 2, 0, null]]);

  const left = fill.planBlankFill([['a', '', ''], ['', 'x', '']], { columns: [0, 1, 2], mode: 'left', startRow: 0 });
  assert.deepEqual(left.cells.map(c => [c.row, c.column, c.value]), [[0, 1, 'a'], [0, 2, 'a'], [1, 2, 'x']]);
  assert.deepEqual(left.cells[1].from, { row: 0, column: 0 });
  assert.equal(left.unfilled, 1);

  const leading = fill.planBlankFill([['h', 'x'], ['', 'y'], ['v', 'z']], { columns: [0], mode: 'above', startRow: 1 });
  assert.deepEqual([leading.cells.length, leading.unfilled], [0, 1]);
  const blankRow = fill.planBlankFill([['h'], ['a'], [''], ['b']], { columns: [0], mode: 'above', startRow: 1 });
  assert.deepEqual([blankRow.cells.length, blankRow.unfilled], [0, 0]);
});

test('fillRuns groups consecutive cells per column; fillBounds spans them', () => {
  const cells = [{ row: 6, column: 0 }, { row: 2, column: 0 }, { row: 3, column: 0 }, { row: 2, column: 2 }];
  assert.deepEqual(fill.fillRuns(cells).map(r => [r.row, r.column, r.cells.length]), [[2, 0, 2], [6, 0, 1], [2, 2, 1]]);
  assert.deepEqual(fill.fillBounds(cells), { top: 2, left: 0, bottom: 6, right: 2 });
  assert.equal(fill.fillBounds([]), null);
});

test('unpivotRows expands each value column into a row and skips blank values', () => {
  const rows = [['A', '华东', 10, '', 30], ['', '', '', '', ''], ['B', '华北', 5, 6, 7]];
  const r = reshape.unpivotRows(rows, [0, 1], [{ index: 2, label: '1月' }, { index: 3, label: '2月' }, { index: 4, label: '3月' }]);
  assert.deepEqual(r.rows, [['A', '华东', '1月', 10], ['A', '华东', '3月', 30], ['B', '华北', '1月', 5], ['B', '华北', '2月', 6], ['B', '华北', '3月', 7]]);
  assert.deepEqual([r.sourceRows, r.skippedBlanks], [2, 1]);
  assert.equal(reshape.unpivotRows(rows, [0], [{ index: 3, label: '2月' }], { skipBlanks: false }).rows.length, 2);
});

const OLD = [['工号', '姓名', '部门', '工资'], ['001', '张三', '销售', 5000], ['002', '李四', '财务', 6000], ['003', '王五', '销售', 5500]];
const NEW = [['工号', '部门', '姓名', '工资', '备注'], ['002', '财务', '李四', 6500, ''], ['001', '销售', '张三', '5000', ''], ['004', '行政', '赵六', 4000, '新人']];
const side = (grid, keys) => ({ headers: grid[0], rows: grid.slice(1), firstRow: 2, firstColumn: 0, keys });

test('diffTables by key aligns reordered rows and columns and reports edits, additions and removals', () => {
  const d = reshape.diffTables(side(OLD, [0]), side(NEW, [0]), { matchByHeader: true });
  assert.deepEqual(d.changes.map(c => [c.kind, c.rowLabel, c.column, c.before, c.after, c.oldAddress, c.newAddress]), [
    ['修改', '002', '工资', 6000, 6500, 'D3', 'D2'],
    ['删除行', '003', '(整行)', '003 | 王五 | 销售 | 5500', '', '4:4', ''],
    ['新增行', '004', '(整行)', '', '004 | 行政 | 赵六 | 4000 | 新人', '', '4:4'],
  ]);
  assert.deepEqual(d.changes[0].mark, { row: 1, column: 3, columns: 1 });
  assert.deepEqual(d.changes[2].mark, { row: 3, column: 0, columns: 5 });
  assert.deepEqual([d.changedCells, d.changedRows, d.unchangedRows, d.addedRows, d.removedRows], [1, 1, 1, 1, 1]);
  assert.deepEqual([d.onlyOldColumns, d.onlyNewColumns], [[], ['备注']]);
});

test('diffTables by position, with tolerance, and pairing duplicate keys in order', () => {
  const pos = reshape.diffTables(side([['k', 'v'], ['a', 1], ['b', 2]], []), side([['k', 'v'], ['a', 1.004], ['b', 3], ['c', 4]], []), { matchByHeader: true, tolerance: 0.01 });
  assert.deepEqual(pos.changes.map(c => [c.kind, c.rowLabel, c.column, c.before, c.after]), [['修改', '第 3 行', 'v', 2, 3], ['新增行', '第 4 行', '(整行)', '', 'c | 4']]);
  const dup = reshape.diffTables(side([['id', 'x'], ['1', 'a'], ['1', 'b']], [0]), side([['id', 'x'], ['1', 'a'], ['1', 'c']], [0]), { matchByHeader: true });
  assert.deepEqual(dup.changes.map(c => [c.before, c.after]), [['b', 'c']]);
  assert.equal(dup.duplicateKeys, 1);
});

test('page layout helpers normalise print titles and escape header/footer ampersands', () => {
  assert.equal(layout.normalizeTitleRows('1'), '$1:$1');
  assert.equal(layout.normalizeTitleRows(' $2:$3 '), '$2:$3');
  assert.throws(() => layout.normalizeTitleRows('A1'), /print_title_rows/);
  assert.equal(layout.normalizeTitleColumns('a:c'), '$A:$C');
  assert.equal(layout.headerFooterText('R&D'), 'R&&D');
  assert.deepEqual(layout.buildIndexGrid([{ name: 'S', address: 'A1:B3', rows: 3, hidden: false }, { name: 'E', address: '', rows: 0, hidden: true }]), [
    ['序号', '工作表', '数据区域', '行数', '状态'], [1, 'S', 'A1:B3', 3, ''], [2, 'E', '（空表）', 0, '隐藏'],
  ]);
});

// ---------------- 工具管线 ----------------

const REPORT = [['部门', '姓名', '金额'], ['销售', '甲', 10], ['', '乙', 20], ['', '丙', 30], ['财务', '丁', 40], ['', '戊', 50]];
const REPORT_MERGES = { 报表: [{ row: 1, column: 0, rows: 3, columns: 1 }, { row: 4, column: 0, rows: 2, columns: 1 }] };

test('fill_blank_cells unmerges vertical merges, fills every cell, and can be undone', async t => {
  const { state, sheet } = book(t, { 报表: REPORT }, { merges: REPORT_MERGES });
  const result = await execute('fill_blank_cells', { sheet: '报表' }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(column(sheet('报表'), 0), ['部门', '销售', '销售', '销售', '财务', '财务']);
  assert.deepEqual(state.unmerged, ['A2:A4', 'A5:A6']);
  assert.equal(sheet('报表').merges.length, 0);
  assert.deepEqual([result.data.filled, result.data.merged_areas, result.data.range], [3, 2, '报表!A3:A6']);
  assert.match(state.confirms[0].title, /3 个空白单元格/);
  assert.match(state.confirms[0].detail, /取消 2 个合并区域/);

  const entry = tools.getUndoStack().at(-1);
  assert.equal(entry.address, 'A3:A6');
  await tools.applyUndo(entry.id);
  assert.deepEqual(column(sheet('报表'), 0), ['部门', '销售', '', '', '财务', '']);
});

test('fill_blank_cells on Excel without the merged-area API needs columns, then unmerges those columns', async t => {
  const { state, sheet } = book(t, { 报表: REPORT }, { merges: REPORT_MERGES, mergedApi: false });
  const missing = await execute('fill_blank_cells', { sheet: '报表' }, ctx(state));
  assert.equal(missing.ok, false);
  assert.match(missing.error, /columns/);
  const result = await execute('fill_blank_cells', { sheet: '报表', columns: ['部门'] }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(state.unmerged, ['A2:A6']);
  assert.deepEqual(column(sheet('报表'), 0), ['部门', '销售', '销售', '销售', '财务', '财务']);
  assert.match(result.data.note, /按上方的值补齐/);
});

test('fill_blank_cells copies the source number format and supports a fixed value', async t => {
  const { state, sheet } = book(t, { 明细: [['日期', '数量', '备注'], [45359, 3, 'a'], ['', '', 'b'], ['', 5, 'c']] }, {
    formats: { 明细: [[], ['yyyy-mm-dd', 'General'], ['General', 'General'], ['General', 'General']] },
  });
  const dates = await execute('fill_blank_cells', { sheet: '明细', columns: ['日期'] }, ctx(state));
  assert.equal(dates.ok, true, dates.error);
  assert.deepEqual(column(sheet('明细'), 0), ['日期', 45359, 45359, 45359]);
  assert.equal(sheet('明细').fmt[2][0], 'yyyy-mm-dd');
  const zeros = await execute('fill_blank_cells', { sheet: '明细', columns: ['数量'], mode: 'value', value: 0 }, ctx(state));
  assert.equal(zeros.ok, true, zeros.error);
  assert.deepEqual(column(sheet('明细'), 1), ['数量', 3, 0, 5]);
  const bad = await execute('fill_blank_cells', { sheet: '明细', columns: ['数量'], mode: 'value' }, ctx(state));
  assert.match(bad.error, /mode=value 需要 value/);
});

test('fill_blank_cells stops when the user declines', async t => {
  const { state, sheet } = book(t, { 报表: REPORT }, { merges: REPORT_MERGES });
  const result = await execute('fill_blank_cells', { sheet: '报表' }, ctx(state, { confirm: async () => false }));
  assert.equal(result.ok, false);
  assert.deepEqual(column(sheet('报表'), 0), ['部门', '销售', '', '', '财务', '']);
  assert.deepEqual(state.unmerged, []);
});

test('fill_blank_cells preserves an empty-result formula while filling neighbouring blanks', async t => {
  const { state, sheet } = book(t, { S: [['部门', '编号'], ['销售', 1], ['', 2], ['', 3]] });
  const ws = sheet('S');
  const formula = '=IF(B3>10,"研发","")';
  ws.f[2] = [formula];
  const result = await execute('fill_blank_cells', { sheet: 'S', columns: ['部门'], mode: 'value', value: '其他' }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.filled, 1);
  assert.equal(ws.f[2][0], formula);
  assert.deepEqual(column(ws, 0), ['部门', '销售', '', '其他']);
});

test('extract_text writes ID-derived columns, keeps digit strings as text and flags numeric IDs', async t => {
  const male = '44030619900307123' + extract.idCheckDigit('44030619900307123');
  const { state, sheet } = book(t, { 名单: [
    ['姓名', '身份证号', '备注'],
    ['张三', male, '手机 138 1234 5678'],
    ['李四', 440306199003071230, '无'],
    ['王五', '11010519491231002X', 'wang@x.cn'],
  ] });
  const ids = await execute('extract_text', { sheet: '名单', column: '身份证号', targets: ['birth_date', 'gender', 'age', 'id_check'] }, ctx(state));
  assert.equal(ids.ok, true, ids.error);
  const s = sheet('名单');
  const age = extract.ageOn('1990-03-07', new Date());
  assert.deepEqual(s.grid.map(r => r.slice(3)), [
    ['出生日期', '性别', '年龄', '身份证校验'],
    ['1990-03-07', '男', age, '正确'],
    ['1990-03-07', '', age, '存成数字，后几位已丢失'],
    ['1949-12-31', '女', extract.ageOn('1949-12-31', new Date()), '正确'],
  ]);
  assert.deepEqual([s.fmt[1][3], s.fmt[3][3]], ['yyyy-mm-dd', 'yyyy-mm-dd']);
  assert.ok(ids.data.warnings.some(w => /存成了数字/.test(w)));
  assert.match(ids.data.note, /周岁/);

  const contacts = await execute('extract_text', { sheet: '名单', column: '备注', targets: ['mobile', 'email'] }, ctx(state));
  assert.equal(contacts.ok, true, contacts.error);
  assert.deepEqual(s.grid.map(r => r.slice(7)), [['手机号', '邮箱'], ['13812345678', ''], ['', ''], ['', 'wang@x.cn']]);
  assert.equal(typeof s.grid[1][7], 'string');
  assert.ok(contacts.data.warnings.some(w => /「手机号」有 2 行没找到/.test(w)));

  const none = await execute('extract_text', { sheet: '名单', column: '姓名', targets: ['email'] }, ctx(state));
  assert.equal(none.ok, false);
  assert.match(none.error, /没有找到邮箱/);
});

test('clean_column rmb_upper writes a new column by default and leaves the amounts intact', async t => {
  const { state, sheet } = book(t, { 报销: [['项目', '金额'], ['差旅', 1234.56], ['餐费', '¥100']] });
  const result = await execute('clean_column', { sheet: '报销', column: '金额', operations: ['rmb_upper'] }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(sheet('报销').grid, [['项目', '金额', '金额_大写'], ['差旅', 1234.56, '壹仟贰佰叁拾肆元伍角陆分'], ['餐费', '¥100', '壹佰元整']]);
});

test('unpivot_columns writes a long table to a new sheet', async t => {
  const { state, sheet } = book(t, { 销量: [['产品', '地区', '1月', '2月', '3月'], ['A', '华东', 10, '', 30], ['B', '华北', 5, 6, 7]] });
  const result = await execute('unpivot_columns', { sheet: '销量', id_columns: ['产品', '地区'], attribute_header: '月份', value_header: '销量' }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.equal(result.data.dest_sheet, '销量_长表');
  assert.deepEqual(sheet('销量_长表').grid, [
    ['产品', '地区', '月份', '销量'],
    ['A', '华东', '1月', 10], ['A', '华东', '3月', 30],
    ['B', '华北', '1月', 5], ['B', '华北', '2月', 6], ['B', '华北', '3月', 7],
  ]);
  assert.equal(result.data.skipped_blanks, 1);
  const byValues = await execute('unpivot_columns', { sheet: '销量', value_columns: ['1月', '2月'] }, ctx(state));
  assert.deepEqual(byValues.data.id_columns, ['产品', '地区', '3月']);
  const neither = await execute('unpivot_columns', { sheet: '销量' }, ctx(state));
  assert.match(neither.error, /id_columns/);
});

test('compare_sheets lists differences on a new sheet and highlights the new version', async t => {
  const { state, sheet } = book(t, { 旧: OLD, 新: NEW, 副本: OLD });
  const result = await execute('compare_sheets', { old_sheet: '旧', new_sheet: '新', keys: ['工号'], highlight: true }, ctx(state));
  assert.equal(result.ok, true, result.error);
  assert.deepEqual(sheet('版本差异').grid, [
    ['变化', '工号', '列', '原值', '新值', '原位置', '新位置'],
    ['新增列', '', '备注', '', '', '', ''],
    ['修改', '002', '工资', 6000, 6500, 'D3', 'D2'],
    ['删除行', '003', '(整行)', '003 | 王五 | 销售 | 5500', '', '4:4', ''],
    ['新增行', '004', '(整行)', '', '004 | 行政 | 赵六 | 4000 | 新人', '', '4:4'],
  ]);
  assert.equal(sheet('新').fills.D2, '#FFEB9C');
  assert.deepEqual(['A4', 'B4', 'C4', 'D4', 'E4'].map(a => sheet('新').fills[a]), Array(5).fill('#C6EFCE'));
  assert.equal(result.data.highlighted, 2);
  assert.deepEqual(Object.keys(sheet('旧').fills), []);

  const same = await execute('compare_sheets', { old_sheet: '旧', new_sheet: '副本' }, ctx(state));
  assert.equal(same.ok, true);
  assert.equal(same.data.dest_sheet, null);
});

test('set_page_layout applies print settings and falls back when fit-to-width height 0 is rejected', async t => {
  const { state, sheet } = book(t, { 报表: REPORT }, { rejectZeroHeight: true });
  const result = await execute('set_page_layout', {
    sheet: '报表', orientation: 'landscape', paper_size: 'a4', fit_to: 'width', print_title_rows: '1', margins: 'narrow',
    center_horizontally: true, page_numbers: true, footer_text: 'R&D 部',
  }, ctx(state));
  assert.equal(result.ok, true, result.error);
  const pl = sheet('报表').pageLayout;
  assert.deepEqual([pl.orientation, pl.paperSize, pl.titleRows, pl.leftMargin, pl.centerHorizontally], ['Landscape', 'A4', '$1:$1', 18, true]);
  assert.deepEqual(pl.zoom, { horizontalFitToPages: 1 });
  assert.equal(pl.headersFooters.defaultForAllPages.centerFooter, '第 &P 页，共 &N 页');
  assert.equal(pl.headersFooters.defaultForAllPages.leftFooter, 'R&&D 部');
  assert.equal(result.data.current.orientation, '横向');
  assert.ok(result.data.applied.includes('所有列缩到一页宽'));

  const nothing = await execute('set_page_layout', { sheet: '报表' }, ctx(state));
  assert.match(nothing.error, /没有给出要改的打印设置/);
  const both = await execute('set_page_layout', { sheet: '报表', fit_to: 'page', scale: 80 }, ctx(state));
  assert.match(both.error, /二选一/);
});

test('create_sheet_index builds a linked index first in the workbook and refreshes it in place', async t => {
  const { state, sheet } = book(t, { 报表: REPORT, 名单: [['a'], [1]], 内部: [['x']], __AI_CP_1: [['cp']] }, { visibility: { 内部: 'Hidden' } });
  const first = await execute('create_sheet_index', {}, ctx(state));
  assert.equal(first.ok, true, first.error);
  const index = sheet('目录');
  assert.deepEqual(index.grid, [['序号', '工作表', '数据区域', '行数'], [1, '报表', 'A1:C6', 6], [2, '名单', 'A1:A2', 2]]);
  assert.deepEqual(index.links.B2, { documentReference: "'报表'!A1", textToDisplay: '报表', screenTip: '转到「报表」' });
  assert.equal(index.position, 0);
  assert.equal(state.active, '目录');

  const again = await execute('create_sheet_index', { include_hidden: true }, ctx(state));
  assert.equal(again.data.refreshed, true);
  assert.equal(state.sheets.filter(s => s.name.startsWith('目录')).length, 1);
  assert.deepEqual(index.grid.at(-1), [3, '内部', 'A1:A1', 1, '隐藏']);
  index.getRange('F2').values = [['人工备注']];
  const shrunk = await execute('create_sheet_index', {}, ctx(state));
  assert.equal(shrunk.ok, true, shrunk.error);
  assert.equal(index.getRange('F2').values[0][0], '人工备注');
  assert.deepEqual(index.getRange('A4:E4').values, [['', '', '', '', '']]);
  assert.equal(index.getRange('E1').values[0][0], '');
  assert.equal(index.links.B4, undefined);
  index.getRange('E2').formulas = [['=""']];
  const blocked = await execute('create_sheet_index', { include_hidden: true }, ctx(state));
  assert.equal(blocked.ok, false);
  assert.equal(index.getRange('E2').formulas[0][0], '=""');
  assert.equal(index.links.B2.textToDisplay, '报表');
});

test('create_sheet_index preserves an unmarked same-name sheet and reuses its generated alternative', async t => {
  const original = [layout.INDEX_HEADERS, [1, '销售', 'A1', 1, '', '人工备注']];
  const { state, sheet } = book(t, { Index: original, 销售: [['a']] });
  const create = () => execute('create_sheet_index', { name: 'index' }, ctx(state));
  const first = await create();
  assert.equal(first.ok, true, first.error);
  assert.equal(first.data.sheet, 'index(2)');
  assert.equal(first.data.refreshed, false);
  const again = await create();
  assert.equal(again.data.sheet, first.data.sheet);
  assert.equal(again.data.refreshed, true);
  assert.deepEqual(sheet('Index').grid, original);
});

test('large amounts use the specified absolute tolerance for both comparison tools', () => {
  const old = [['id', '金额'], ['A', 1_000_000_000]];
  const next = [['id', '金额'], ['A', 1_000_000_001]];
  const ledger = (data, label) => ({ rows: data.slice(1), keys: [0], compare: [1], label });
  for (const [tolerance, expected] of [[0, 1], [0.5, 1], [1, 0]]) {
    const diff = reshape.diffTables(side(old, [0]), side(next, [0]), { matchByHeader: true, tolerance });
    assert.equal(diff.changedCells, expected);
    const result = reshape.reconcileTables(ledger(old, '旧'), ledger(next, '新'), {
      keyHeaders: ['id'], compareHeaders: ['金额'], tolerance,
    });
    assert.equal(result.mismatched, expected);
  }
});
