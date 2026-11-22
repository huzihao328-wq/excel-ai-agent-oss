const test = require('node:test');
const assert = require('node:assert/strict');
const { sourceLoader } = require('./load-source.cjs');

const { findRangeRefs } = sourceLoader()('src/taskpane/rangeLinks.ts');
const sheets = ['汇总', '月度汇总', 'Sheet1', 'Q1 销售', "O'Brien"];
const refs = text => findRangeRefs(text, sheets).map(r => [text.slice(r.start, r.end), r.sheet, r.address]);

test('replies link addresses on sheets that exist, including inside Chinese sentences', () => {
  assert.deepEqual(refs('结果写在汇总!A1:D12，共 11 行'), [['汇总!A1:D12', '汇总', 'A1:D12']]);
  assert.deepEqual(refs('见 月度汇总!$B$2'), [['月度汇总!$B$2', '月度汇总', '$B$2']]);
  assert.deepEqual(refs('整列 Sheet1!C:E 和整行 Sheet1!2:5'), [
    ['Sheet1!C:E', 'Sheet1', 'C:E'],
    ['Sheet1!2:5', 'Sheet1', '2:5'],
  ]);
});

test('quoted names, corner brackets and case follow how the model writes them', () => {
  assert.deepEqual(refs("在 'Q1 销售'!B2:B9 里"), [["'Q1 销售'!B2:B9", 'Q1 销售', 'B2:B9']]);
  assert.deepEqual(refs("'O''Brien'!a1"), [["'O''Brien'!a1", "O'Brien", 'A1']]);
  assert.deepEqual(refs('已写入新表「汇总」A1:C4'), [['「汇总」A1:C4', '汇总', 'A1:C4']]);
  assert.deepEqual(refs('切到「汇总」看看'), [['「汇总」', '汇总', undefined]]);
  assert.deepEqual(refs('sheet1!b3'), [['sheet1!b3', 'Sheet1', 'B3']]);
});

test('no link for unknown sheets, partial words or things that only look like addresses', () => {
  assert.deepEqual(refs('删掉的表 旧数据!A1:B2'), []);
  assert.deepEqual(refs('MySheet1!A1'), []);
  assert.deepEqual(refs('汇总!ABC 与 汇总!A1B'), []);
  assert.deepEqual(refs('做完了汇总! 很快'), []);
  assert.deepEqual(refs('「新表」A1'), []);
  assert.deepEqual(findRangeRefs('汇总!A1', []), []);
});
