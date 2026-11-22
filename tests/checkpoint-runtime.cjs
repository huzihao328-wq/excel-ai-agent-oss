const test = require('node:test');
const assert = require('node:assert/strict');
const { sourceLoader } = require('./load-source.cjs');

function workbook() {
  const items = [];
  let failCleanup = false;
  function sheet(name, visibility = 'Visible') {
    const ws = {
      name, visibility, position: items.length,
      tables: { items: [], load() {} }, load() {}, activate() {},
      getRange() { return { getSpecialCellsOrNullObject: () => ({ isNullObject: true, load() {} }) }; },
      copy() { return sheet(this.name + '-copy', this.visibility); },
      delete() {
        if (this.visibility === 'VeryHidden') throw Error('InvalidOperation');
        if (failCleanup) throw Error('Cleanup unavailable');
        items.splice(items.indexOf(this), 1);
      },
    };
    items.push(ws);
    return ws;
  }
  const original = sheet('Sales');
  const sheets = {
    get items() { return [...items]; }, load() {},
    getActiveWorksheet: () => original,
    getItem: name => items.find(s => s.name === name),
  };
  global.Excel = {
    SheetVisibility: { visible: 'Visible', hidden: 'Hidden', veryHidden: 'VeryHidden' },
    WorksheetPositionType: { end: 'End' },
    SpecialCellType: { formulas: 'Formulas' },
    run: fn => fn({ workbook: { worksheets: sheets }, sync: async () => {} }),
  };
  global.Office = { context: { document: { settings: { set() {}, remove() {}, saveAsync() {} } } } };
  return { items, original, addSheet: sheet, failCleanup() { failCleanup = true; } };
}

test('successive checkpoints remove VeryHidden backups without touching the user sheet', async () => {
  const wb = workbook();
  const cp = sourceLoader()('src/core/tools/checkpoint.ts');
  const events = [];
  cp.onCheckpointChange((_, event) => events.push(event));
  cp.beginCheckpointSession('first');
  const first = await cp.ensureCheckpoint();
  cp.beginCheckpointSession('second');
  const second = await cp.ensureCheckpoint();
  assert.notEqual(first.id, second.id);
  assert.deepEqual(events, ['creating', 'created', 'creating', 'created']);
  assert.equal(wb.items.length, 2);
  assert.ok(wb.items.includes(wb.original));
  assert.equal(wb.original.visibility, 'Visible');
  assert.equal(wb.items[1].name, second.sheets[0].backupName);
  assert.equal(wb.items[1].visibility, 'VeryHidden');
});

test('cleanup failure preserves the new checkpoint and is not reported as snapshot failure', async () => {
  const wb = workbook();
  const cp = sourceLoader()('src/core/tools/checkpoint.ts');
  cp.beginCheckpointSession('first');
  await cp.ensureCheckpoint();
  wb.failCleanup();
  const events = [];
  cp.onCheckpointChange((_, event) => events.push(event));
  cp.beginCheckpointSession('second');
  const second = await cp.ensureCheckpoint();
  assert.deepEqual(events, ['creating', 'created', 'cleanup-failed']);
  assert.equal(cp.getCheckpoint(), second);
  assert.equal(await cp.ensureCheckpoint(), second);
  assert.equal(wb.items.length, 3);
});

test('checkpoint freezes original cross-sheet formulas before later edits', async () => {
  const wb = workbook();
  const expressions = [['=SUMIF(Source!A2:A6,A2,Source!E2:E6)']];
  wb.original.getRange = () => ({
    getSpecialCellsOrNullObject: () => ({
      isNullObject: false, load() {},
      areas: { load() {}, items: [{ address: 'Sales!B2', formulas: expressions }] },
    }),
  });
  const cp = sourceLoader()('src/core/tools/checkpoint.ts');
  cp.beginCheckpointSession('formulas');
  const snapshot = await cp.ensureCheckpoint();
  expressions[0][0] = '=SUMIF(#REF!,A2,#REF!)';
  assert.deepEqual(snapshot.formulas, [{
    sheet: 'Sales', range: 'B2', formulas: [['=SUMIF(Source!A2:A6,A2,Source!E2:E6)']],
  }]);
});

// ---- 文件快照模式：备份存进 IndexedDB（此处 mock），工作簿里不留备份表 ----

function fileWorkbook({ insertFails = false } = {}) {
  const wb = workbook();
  const stored = new Map();
  const deleted = [];
  const inserted = [];
  const store = {
    putCheckpointFile: async rec => { stored.set(rec.id, rec); },
    getCheckpointFile: async id => stored.get(id) || null,
    hasCheckpointFile: async id => stored.has(id),
    deleteCheckpointFile: async id => { deleted.push(id); stored.delete(id); },
    pruneCheckpointFiles: async () => {},
  };
  global.indexedDB = {};
  global.btoa = s => Buffer.from(s, 'binary').toString('base64');
  Object.assign(global.Office, {
    FileType: { Compressed: 'compressed' },
    AsyncResultStatus: { Succeeded: 'succeeded' },
  });
  global.Office.context.requirements = { isSetSupported: (set, v) => set === 'ExcelApi' && v === '1.13' };
  global.Office.context.document.getFileAsync = (_type, _opts, cb) => cb({
    status: 'succeeded',
    value: {
      sliceCount: 2,
      getSliceAsync: (i, done) => done({ status: 'succeeded', value: { data: i === 0 ? [80, 75] : [3, 4] } }),
      closeAsync: done => done(),
    },
  });
  const ctxWorkbook = {
    worksheets: null,
    insertWorksheetsFromBase64(base64, options) {
      if (insertFails) throw Error('Unsupported content');
      inserted.push({ base64, options });
      for (const name of options.sheetNamesToInsert) wb.addSheet(name);
    },
  };
  const run = global.Excel.run;
  global.Excel.run = fn => run(ctx => fn({ ...ctx, workbook: { ...ctx.workbook, insertWorksheetsFromBase64: ctxWorkbook.insertWorksheetsFromBase64 } }));
  const cp = sourceLoader({ 'src/core/checkpointStore.ts': store })('src/core/tools/checkpoint.ts');
  return { wb, cp, stored, deleted, inserted };
}

test('file checkpoint stores the workbook bytes outside the workbook and removes legacy backups', async () => {
  const { wb, cp, stored } = fileWorkbook();
  wb.addSheet('__AI_CP_old_00', 'Hidden');
  cp.beginCheckpointSession('file');
  const snap = await cp.ensureCheckpoint();
  assert.equal(snap.storage, 'file');
  assert.deepEqual([...new Uint8Array(stored.get(snap.id).bytes)], [80, 75, 3, 4]);
  assert.deepEqual(wb.items.map(s => s.name), ['Sales'], 'no backup sheets left in the workbook');
});

test('file checkpoint restore swaps in the snapshot sheets and drops the edited ones', async () => {
  const { wb, cp, deleted, inserted } = fileWorkbook();
  cp.beginCheckpointSession('file');
  const snap = await cp.ensureCheckpoint();
  wb.addSheet('Scratch');
  const message = await cp.restoreCheckpoint();
  assert.match(message, /1 个 Sheet/);
  assert.equal(inserted.length, 1);
  assert.deepEqual(inserted[0].options.sheetNamesToInsert, ['Sales']);
  assert.equal(inserted[0].base64, Buffer.from([80, 75, 3, 4]).toString('base64'));
  assert.deepEqual(wb.items.map(s => s.name), ['Sales']);
  assert.notEqual(wb.items[0], wb.original, 'the edited sheet was replaced by the snapshot copy');
  assert.ok(deleted.includes(snap.id), 'snapshot is discarded after restore');
  assert.equal(cp.getCheckpoint(), null);
});

test('file checkpoint restore leaves the workbook untouched when inserting fails', async () => {
  const { wb, cp } = fileWorkbook({ insertFails: true });
  cp.beginCheckpointSession('file');
  const snap = await cp.ensureCheckpoint();
  wb.addSheet('Scratch');
  await assert.rejects(cp.restoreCheckpoint(), /工作簿未改动/);
  assert.deepEqual(wb.items.map(s => s.name), ['Sales', 'Scratch']);
  assert.equal(cp.getCheckpoint(), snap, 'checkpoint remains available for another try');
});
