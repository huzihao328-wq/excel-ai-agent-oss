const test = require('node:test');
const assert = require('node:assert/strict');
const { sourceLoader } = require('./load-source.cjs');

function browser(t, initial = {}) {
  const values = new Map(Object.entries(initial));
  const originals = new Map(['window', 'localStorage', 'OfficeRuntime', 'fetch'].map(key => [key, Object.getOwnPropertyDescriptor(global, key)]));
  const timers = new Set();
  global.window = {
    setTimeout(fn, delay) { const id = setTimeout(fn, Math.min(delay, 30)); timers.add(id); return id; },
    clearTimeout(id) { clearTimeout(id); timers.delete(id); },
    addEventListener() {},
  };
  global.localStorage = {
    getItem: key => values.get(key) ?? null,
    setItem: (key, value) => values.set(key, String(value)),
    removeItem: key => values.delete(key),
  };
  delete global.OfficeRuntime;
  t.after(() => {
    for (const id of timers) clearTimeout(id);
    for (const [key, desc] of originals) {
      if (desc) Object.defineProperty(global, key, desc);
      else delete global[key];
    }
  });
  return values;
}

const profile = { updatedAt: 10, skills: [{ name: 'old skill' }], memories: [{ content: 'old memory' }], templates: [{ name: 'old template' }] };

test('conversation: malformed index entries are ignored without crashing the pane', t => {
  const valid = { id: 'valid', title: '保留的对话', createdAt: 1, updatedAt: 2 };
  browser(t, { 'excel-ai-agent-conv-index': JSON.stringify({ currentId: 'valid', items: [null, {}, 'bad', { id: 'no-title' }, valid] }) });
  const { listConversations } = sourceLoader()('src/core/conversations.ts');
  assert.deepEqual(listConversations(), [valid]);
});

test('conversation: blocked storage cannot strand send in its running state', t => {
  browser(t);
  global.localStorage.setItem = () => { throw new Error('storage full'); };
  const api = sourceLoader()('src/core/conversations.ts');
  assert.equal(api.setCurrentConversationId('new'), false);
  assert.equal(api.saveConversation({ id: 'new', title: 'test', createdAt: 1, updatedAt: 2, items: [], messages: [] }), false);
});

test('startup: a stalled backup write cannot block hydration; concurrent hydration runs once', async t => {
  browser(t, { 'excel-ai-agent-skills': '[]' });
  let gets = 0, puts = 0;
  let release;
  global.fetch = async (_url, init) => {
    if (init.method === 'GET') { gets++; return new Response('{}'); }
    puts++;
    return new Promise(resolve => { release = () => resolve(new Response('{}')); });
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  const first = store.hydrateUserData();
  assert.equal(first, store.hydrateUserData());
  await first;
  assert.equal(gets, 1);
  assert.equal(puts, 1);
  assert.ok(release, 'backup has started but hydration already completed');
  release();
});

test('storage: an intentional empty collection does not resurrect from backup', async t => {
  const values = browser(t, {
    'excel-ai-agent-skills': '[]', 'excel-ai-agent-memory': '[]', 'excel-ai-agent-prompt-templates': '[]',
  });
  const writes = [];
  global.fetch = async (_url, init) => {
    if (init.method === 'PUT') writes.push(JSON.parse(init.body));
    return new Response(JSON.stringify(profile));
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  await store.flushUserData();
  assert.equal(values.get('excel-ai-agent-skills'), '[]');
  assert.equal(values.get('excel-ai-agent-memory'), '[]');
  assert.equal(values.get('excel-ai-agent-prompt-templates'), '[]');
  assert.deepEqual(writes.at(-1).skills, []);
});

test('storage: absent local collections still inherit saved data', async t => {
  const values = browser(t);
  global.fetch = async () => new Response(JSON.stringify(profile));
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  await store.flushUserData();
  assert.deepEqual(JSON.parse(values.get('excel-ai-agent-skills')), profile.skills);
  assert.deepEqual(JSON.parse(values.get('excel-ai-agent-memory')), profile.memories);
});

test('storage: unavailable backups and an empty local store never overwrite backups with blanks', async t => {
  const values = browser(t);
  let writes = 0;
  global.fetch = async (_url, init) => {
    if (init.method === 'PUT') writes++;
    throw new Error('offline');
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  await store.flushUserData();
  assert.equal(values.size, 0);
  assert.equal(writes, 0);
});

test('storage: timed-out writes release the queue for the next profile', async t => {
  browser(t, { 'excel-ai-agent-skills': '[]' });
  let writes = 0;
  global.fetch = async (_url, init) => {
    if (init.method === 'GET') return new Response('{}');
    writes++;
    if (writes > 1) return new Response('{}');
    return new Promise((_, reject) => init.signal.addEventListener('abort', () => reject(new Error('timeout')), { once: true }));
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  await store.flushUserData();
  assert.equal(writes, 2);
});

function attachments(t, hooks = {}) {
  browser(t);
  const stored = new Map();
  const load = sourceLoader({
    'src/core/attachmentStore.ts': {
      listStoredAttachments: async () => [],
      getStoredAttachment: async name => stored.get(name) || null,
      putStoredAttachment: async item => { stored.set(item.name, item); },
      deleteStoredAttachment: async name => { stored.delete(name); },
      ...hooks,
    },
  });
  return { api: load('src/core/tools/externalFiles.ts'), stored };
}

test('storage: a late Office write cannot overwrite a newer profile after timeout', async t => {
  browser(t, { 'excel-ai-agent-skills': '["first"]' });
  const officeWrites = [];
  const diskWrites = [];
  let releaseFirst;
  let secondWritten;
  const second = new Promise(resolve => { secondWritten = resolve; });
  global.OfficeRuntime = { storage: {
    getItem: async () => null,
    setItem: async (_key, value) => {
      officeWrites.push(JSON.parse(value));
      if (officeWrites.length === 1) await new Promise(resolve => { releaseFirst = resolve; });
      else secondWritten();
    },
  } };
  global.fetch = async (_url, init) => {
    if (init.method === 'PUT') diskWrites.push(JSON.parse(init.body));
    return new Response('{}');
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  localStorage.setItem('excel-ai-agent-skills', '["second"]');
  await store.flushUserData();
  assert.equal(officeWrites.length, 1, 'non-cancellable Office writes overlapped');
  assert.deepEqual(diskWrites.at(-1).skills, ['second'], 'disk backup must remain available');
  releaseFirst();
  await second;
  assert.deepEqual(officeWrites.at(-1).skills, ['second']);
});
test('storage: the API key never reaches the disk profile but stays in Office storage', async t => {
  browser(t, { 'excel-ai-agent-settings': JSON.stringify({ model: 'm', apiKey: 'sk-secret' }) });
  const officeWrites = [];
  const diskWrites = [];
  global.OfficeRuntime = { storage: {
    getItem: async () => null,
    setItem: async (_key, value) => { officeWrites.push(JSON.parse(value)); },
  } };
  global.fetch = async (_url, init) => {
    if (init.method === 'PUT') diskWrites.push(JSON.parse(init.body));
    return new Response('{}', { status: 404 });
  };
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  await store.flushUserData();
  assert.ok(diskWrites.length > 0);
  for (const w of diskWrites) assert.equal(w.settings.apiKey, undefined);
  assert.equal(diskWrites.at(-1).settings.model, 'm');
  assert.equal(officeWrites.at(-1).settings.apiKey, 'sk-secret');
});

test('storage: a newer key-less disk profile still inherits the key from Office storage', async t => {
  const values = browser(t);
  global.OfficeRuntime = { storage: {
    getItem: async () => JSON.stringify({ updatedAt: 1, settings: { model: 'old', apiKey: 'sk-office' } }),
    setItem: async () => {},
  } };
  global.fetch = async (_url, init) => init.method === 'PUT'
    ? new Response('{}')
    : new Response(JSON.stringify({ updatedAt: 5, settings: { model: 'new' } }));
  const store = sourceLoader()('src/core/durableStore.ts');
  await store.hydrateUserData();
  assert.deepEqual(JSON.parse(values.get('excel-ai-agent-settings')), { model: 'new', apiKey: 'sk-office' });
});

test('user-data server: strips API keys and rejects cross-site browser callers', async () => {
  const os = require('node:os');
  const fs = require('node:fs');
  const path = require('node:path');
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eaa-'));
  const prev = process.env.LOCALAPPDATA;
  process.env.LOCALAPPDATA = dir;
  try {
    delete require.cache[require.resolve('../scripts/userdata-store.js')];
    const { handleUserDataRequest, profilePath } = require('../scripts/userdata-store.js');
    const { EventEmitter } = require('node:events');
    const call = (method, headers, body) => new Promise(resolve => {
      const req = new EventEmitter();
      Object.assign(req, { method, url: '/api/user-data', headers, socket: { remoteAddress: '127.0.0.1' } });
      req.off = req.removeListener;
      const res = { headersSent: false, writeHead(status) { this.status = status; this.headersSent = true; }, end(b) { resolve({ status: this.status, body: b }); } };
      handleUserDataRequest(req, res);
      if (body !== undefined) { req.emit('data', Buffer.from(body)); req.emit('end'); }
    });
    const put = await call('PUT', { origin: 'https://localhost:3000' }, JSON.stringify({ updatedAt: 1, settings: { model: 'x', apiKey: 'sk' } }));
    assert.equal(put.status, 200);
    assert.deepEqual(JSON.parse(fs.readFileSync(profilePath(), 'utf8')).settings, { model: 'x' });
    assert.equal((await call('GET', { origin: 'https://evil.example' })).status, 403);
    assert.equal((await call('GET', { 'sec-fetch-site': 'cross-site' })).status, 403);
    assert.equal((await call('GET', { 'sec-fetch-site': 'same-origin' })).status, 200);
  } finally {
    if (prev === undefined) delete process.env.LOCALAPPDATA; else process.env.LOCALAPPDATA = prev;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

const csv = text => new TextEncoder().encode(text).buffer;

test('CSV: account identifiers keep leading zeros and all digits', async t => {
  const { api } = attachments(t);
  await api.attachExternalBuffer('ids.csv', csv('id,amount\n00123,12.50\n123456789012345678,0'));
  assert.deepEqual(api.getExternalFiles()[0].sheets[0].rows[1], ['00123', '12.50']);
  assert.equal(api.getExternalFiles()[0].sheets[0].rows[2][0], '123456789012345678');
});

test('attachments: deleting then reattaching a name preserves the newest persisted bytes', async t => {
  const { api, stored } = attachments(t);
  await api.attachExternalBuffer('same.csv', csv('id\nold'));
  await api.attachExternalBuffer('second.csv', csv('id\n2'));
  api.removeExternalFile('same.csv');
  await api.attachExternalBuffer('same.csv', csv('id\nnew'));
  assert.equal(new TextDecoder().decode(stored.get('same.csv').bytes), 'id\nnew');
});

test('attachments: a slow older file read cannot replace the latest file or restore a removed file', async t => {
  const { api } = attachments(t);
  let release;
  const slow = api.loadExternalFile({ name: 'same.csv', size: 10, arrayBuffer: () => new Promise(resolve => { release = resolve; }) });
  await api.loadExternalFile({ name: 'same.csv', size: 10, arrayBuffer: async () => csv('id\nnew') });
  release(csv('id\nold'));
  await slow;
  assert.equal(api.getExternalFiles()[0].sheets[0].rows[1][0], 'new');
  let releaseRemoved;
  const removed = api.loadExternalFile({ name: 'same.csv', size: 10, arrayBuffer: () => new Promise(resolve => { releaseRemoved = resolve; }) });
  api.removeExternalFile('same.csv');
  releaseRemoved(csv('id\nremoved'));
  await removed;
  assert.equal(api.getExternalFiles().length, 0);
});

test('attachments: removing a saved file while restoration is pending cannot resurrect it', async t => {
  let release;
  const { api } = attachments(t, { listStoredAttachments: () => new Promise(resolve => { release = resolve; }) });
  const restoring = api.hydrateExternalFiles();
  api.removeExternalFile('removed.csv');
  release([{ name: 'removed.csv', bytes: csv('id\n1'), loadedAt: 1 }]);
  assert.equal(await restoring, 0);
  assert.equal(api.getExternalFiles().length, 0);
});

test('attachments: replacing a selected file clears the old range reference', async t => {
  const { api } = attachments(t);
  await api.attachExternalBuffer('same.csv', csv('id\nold'));
  api.setExternalSelection({ file: 'same.csv', sheet: 'Sheet1', range: 'A1:A2' });
  await api.attachExternalBuffer('same.csv', csv('id\nnew'));
  assert.equal(api.getExternalSelection(), null);
});
