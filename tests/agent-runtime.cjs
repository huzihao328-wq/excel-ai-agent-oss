const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const load = sourceLoader();
const { readChatResponse } = load('src/core/chatResponse.ts');
const { DEFAULT_SETTINGS } = load('src/core/settings.ts');
const { normalizeToolArguments } = load('src/core/toolArguments.ts');
const { TOOL_DEFINITIONS } = load('src/core/tools/definitions.ts');
const settings = { ...DEFAULT_SETTINGS, baseUrl: 'http://localhost:11434/v1', apiKey: '', autoLearnSkills: false, autoCheckpoint: false, useMemory: false };
const callbacks = { onAssistantDelta() {}, onAssistantDone() {}, onToolStart() {}, onToolEnd() {} };
const execCtx = { settings, confirm: async () => true, choose: async () => ({ kind: 'skipped' }) };
const call = (id, name = 'write_range', args = { sheet: 'S', start_cell: 'A1', values: [[11]] }) => ({ id, type: 'function', function: { name, arguments: JSON.stringify(args) } });
const turn = (toolCalls = [], content = '') => ({ content, toolCalls, finishReason: toolCalls.length ? 'tool_calls' : 'stop', truncated: false });
const event = (data) => `data: ${JSON.stringify(data)}\r\n\r\n`;
const delta = (value, finish_reason = null) => ({ choices: [{ delta: value, finish_reason }] });
function sse(events, end = true, splitBytes = false) {
  const bytes = new TextEncoder().encode(events.map(event).join('') + (end ? 'data: [DONE]\r\n\r\n' : ''));
  return new Response(new ReadableStream({ start(controller) {
    if (splitBytes) for (const byte of bytes) controller.enqueue(Uint8Array.of(byte));
    else controller.enqueue(bytes);
    controller.close();
  } }), { headers: { 'content-type': 'text/event-stream' } });
}
const read = (response) => readChatResponse(response, () => {}, new AbortController().signal);

test('stream: repeated digits, braces, Chinese and JSON strings survive arbitrary fragmentation', async () => {
  const args = JSON.stringify({ sheet: 'S', start_cell: 'A1', values: [[11, 1000, '哈哈', 'repeatrepeat']] });
  const events = [delta({ tool_calls: [{ index: 0, id: 'c1', function: { name: 'write_range', arguments: '' } }] })];
  for (const char of args) events.push(delta({ tool_calls: [{ index: 0, function: { arguments: char } }] }));
  events.push(delta({}, 'tool_calls'));
  const result = await read(sse(events, true, true));
  assert.equal(result.toolCalls[0].function.arguments, args);
  assert.deepEqual(JSON.parse(result.toolCalls[0].function.arguments).values, [[11, 1000, '哈哈', 'repeatrepeat']]);
});

test('stream: repeated text tokens retain numeric meaning', async () => {
  const result = await read(sse(['金额', '1', '1', '1', '元'].map(content => delta({ content }))));
  assert.equal(result.content, '金额111元');
});

test('stream: cumulative argument snapshots in delta remain compatible without altering native JSON', async () => {
  const args = JSON.stringify({ sheet: 'S', start_cell: 'A1', values: [[11]] });
  const pieces = [args.slice(0, 10), args.slice(0, 30), args];
  const events = pieces.map(arguments => delta({ tool_calls: [{ index: 0, id: 'a', function: { name: 'write_range', arguments } }] }));
  const result = await read(sse(events));
  assert.equal(result.toolCalls[0].function.arguments, args);
});

test('stream: message snapshots and non-stream JSON responses are accepted', async () => {
  const result = await read(sse([
    { choices: [{ message: { content: '完成', tool_calls: [call('a')] } }] },
    { choices: [{ message: { content: '完成了', tool_calls: [call('a')] }, finish_reason: 'stop' }] },
  ]));
  assert.equal(result.content, '完成了');
  assert.equal(result.toolCalls.length, 1);
  const json = new Response(JSON.stringify({ choices: [{ message: { content: 'pong' }, finish_reason: 'stop' }] }), { headers: { 'content-type': 'application/json' } });
  assert.equal((await read(json)).content, 'pong');
});

test('stream: partial calls followed by provider error never execute', async () => {
  await assert.rejects(read(sse([delta({ tool_calls: [{ index: 0, ...call('a') }] }), { error: { message: 'quota exhausted' } }])), /quota exhausted/);
});

test('stream: unmarked EOF, empty response, malformed JSON and duplicate IDs fail visibly', async () => {
  await assert.rejects(read(sse([delta({ tool_calls: [{ index: 0, ...call('a') }] })], false)), /完成前中断/);
  await assert.rejects(read(sse([])), /没有返回可用/);
  await assert.rejects(read(new Response('data: {broken}\n\n')), /无效的流式 JSON/);
  await assert.rejects(read(sse([delta({ tool_calls: [{ index: 0, ...call('a') }, { index: 1, ...call('a') }] })])), /重复/);
});

test('stream: DONE closes a connection that remains open', async () => {
  let cancelled = false;
  const response = new Response(new ReadableStream({ start(c) {
    c.enqueue(new TextEncoder().encode(event(delta({ content: 'done' })) + 'data: [DONE]\n\n'));
  }, cancel() { cancelled = true; } }));
  assert.equal((await read(response)).content, 'done');
  assert.equal(cancelled, true);
});

test('stream: cancellation releases a stalled reader', async () => {
  const controller = new AbortController();
  const response = new Response(new ReadableStream({ start(c) { c.enqueue(new TextEncoder().encode(': waiting\n\n')); } }));
  const pending = readChatResponse(response, () => {}, controller.signal);
  controller.abort();
  await assert.rejects(pending, error => error.name === 'AbortError');
});

test('connection: HTTP 200 with an error envelope is not a successful connection', async () => {
  const oldFetch = global.fetch;
  try {
    global.fetch = async () => new Response(JSON.stringify({ error: { message: 'wrong endpoint' } }), { headers: { 'content-type': 'application/json' } });
    await assert.rejects(load('src/core/llm.ts').testConnection(settings), /未返回有效/);
  } finally { global.fetch = oldFetch; }
});

test('pivot schema keeps object value fields available to strict providers', () => {
  const { adaptToolsForEndpoint } = load('src/core/toolSchema.ts');
  const tools = adaptToolsForEndpoint(TOOL_DEFINITIONS, 'https://generativelanguage.googleapis.com/v1beta/openai');
  const schema = tools.find(t => t.function.name === 'create_pivot_table').function.parameters;
  const args = { source_sheet: 'S', source_range: 'A1:B20', values: [{ field: '金额', aggregation: 'average' }] };
  assert.deepEqual(normalizeToolArguments(args, schema), args);
  assert.throws(() => normalizeToolArguments({ ...args, values: [] }, schema), /至少需要/);
  assert.throws(() => normalizeToolArguments({ ...args, values: [{ field: '金额', aggregation: 'typo' }] }, schema), /不符合/);
});

test('arguments: nested booleans normalize; missing fields, wrong types and enums fail', () => {
  const schema = TOOL_DEFINITIONS.find(t => t.function.name === 'sort_range').function.parameters;
  const original = { sheet: 'S', keys: [{ column: '金额', ascending: 'false' }] };
  assert.equal(normalizeToolArguments(original, schema).keys[0].ascending, false);
  assert.equal(original.keys[0].ascending, 'false');
  assert.throws(() => normalizeToolArguments({}, schema), /sheet/);
  // 名字是纯数字的表（如 2024）常被模型写成数字：按文本接受；其他类型仍然报错
  assert.equal(normalizeToolArguments({ sheet: 2024 }, schema).sheet, '2024');
  assert.throws(() => normalizeToolArguments({ sheet: true }, schema), /字符串/);
  assert.throws(() => normalizeToolArguments({ sheet: 'S', keys: [{ ascending: 'no' }] }, schema), /keys\[0\].ascending/);
  const clear = TOOL_DEFINITIONS.find(t => t.function.name === 'clear_range').function.parameters;
  assert.throws(() => normalizeToolArguments({ sheet: 'S', range: 'A1', what: 'typo' }, clear), /contents/);
});

function agentHarness(stream, executeTool = async () => ({ ok: true, data: { written: true } })) {
  return sourceLoader({
    'src/core/llm.ts': { streamChatCompletion: stream, probeContextWindowInfo: async () => null },
    'src/core/tools/excelTools.ts': { TOOL_REGISTRY: { create_sheet: { mutating: true }, write_range: { mutating: true } }, executeTool, getQuickContext: async () => ({ sheetNames: ['S'], activeSheet: 'S', selectionAddress: 'A1' }) },
    'src/core/skills.ts': { matchSkills: () => [] },
  })('src/core/agent.ts');
}

test('tool card summary: plain wording, no raw parameter names or JSON', () => {
  const { summarizeArgs } = agentHarness(async () => ({}));
  const summary = args => summarizeArgs(JSON.stringify(args));
  assert.equal(summary({ sheet: '销售', start_row: 2, count: 3 }), '销售 · 起始行 2 · 数量 3');
  assert.equal(summary({ sheet: '销售', keys: [{ column: '金额', ascending: false }, { column: '日期' }], has_headers: true }), '销售 · 按 金额、日期');
  assert.equal(summary({ sheet: '销售', range: 'A1:D9', columns: ['地区', '金额'], max_parts: 4, ignore_empty: true }), '销售!A1:D9 · 列 地区、金额');
  assert.equal(summary({ sheet: 'S', start_cell: 'A1', values: [[1, 2], [3, 4]] }), 'S!A1 · 写入 2×2');
  // 没有约定叫法的参数、纯对象都不上摘要
  assert.equal(summary({ sheet: 'S', margins: { top: 1 }, tolerance: 0.01 }), 'S');
});

test('agent: failed prerequisite skips pending writes and keeps all tool results paired', async () => {
  let rounds = 0;
  const executed = [];
  const history = [];
  const { runAgent } = agentHarness(async (_settings, messages) => {
    if (rounds++ === 0) return turn([call('read', 'read_range', { sheet: 'missing' }), call('write')]);
    assert.equal(JSON.parse(messages.at(-1).content).ok, false);
    assert.match(messages.at(-1).content, /本工具未执行/);
    return turn([], '读取失败，请核对表名。');
  }, async name => { executed.push(name); return { ok: false, error: 'Sheet not found' }; });
  await runAgent('读表后写入', history, settings, execCtx, callbacks, new AbortController().signal);
  assert.deepEqual(executed, ['read_range']);
  assert.deepEqual(history.filter(m => m.role === 'tool').map(m => m.tool_call_id), ['read', 'write']);
});

test('agent: identical failures stop re-executing even when argument key order changes', async () => {
  let rounds = 0, attempts = 0;
  const { runAgent } = agentHarness(async () => {
    rounds++;
    if (rounds > 3) return turn([], '需要核对输入。');
    return turn([call(`attempt-${rounds}`, 'read_range', rounds === 2 ? { range: 'A1', sheet: 'S' } : { sheet: 'S', range: 'A1' })]);
  }, async () => { attempts++; return { ok: false, error: 'invalid range' }; });
  const history = [];
  await runAgent('读取数据', history, settings, execCtx, callbacks, new AbortController().signal);
  assert.equal(attempts, 2);
  assert.match(history.filter(m => m.role === 'tool').at(-1).content, /相同工具和参数已失败两次/);
});

test('agent: oversized tool results preserve success and instruct the model to narrow its read', async () => {
  let rounds = 0;
  const { runAgent } = agentHarness(async (_settings, messages) => {
    if (rounds++ === 0) return turn([call('read', 'read_range', { sheet: 'S' })]);
    const result = JSON.parse(messages.at(-1).content);
    assert.equal(result.ok, true);
    assert.equal(result.truncated, true);
    assert.match(result.hint, /更小的范围/);
    return turn([], '请缩小查询范围。');
  }, async () => ({ ok: true, data: '值'.repeat(50000) }));
  await runAgent('读取数据', [], settings, execCtx, callbacks, new AbortController().signal);
  assert.equal(rounds, 2);
});

test('agent: successful workbook repair allows previously failed arguments to run again', async () => {
  let rounds = 0, readAttempts = 0, repaired = false;
  const { runAgent } = agentHarness(async () => {
    rounds++;
    if (rounds === 3) return turn([call('repair', 'create_sheet', { name: 'S' })]);
    if (rounds <= 4) return turn([call(`read-${rounds}`, 'read_range', { sheet: 'S' })]);
    return turn([], '读取完成');
  }, async name => {
    if (name === 'create_sheet') { repaired = true; return { ok: true }; }
    readAttempts++;
    return repaired ? { ok: true, data: [] } : { ok: false, error: 'sheet missing' };
  });
  await runAgent('读取 S，没有则创建', [], settings, execCtx, callbacks, new AbortController().signal);
  assert.equal(readAttempts, 3);
});

test('agent: tiny windows and oversized current instructions fail before calling the provider', async () => {
  let requests = 0;
  const { runAgent } = agentHarness(async () => { requests++; return turn([], 'done'); });
  await assert.rejects(runAgent('读取', [], { ...settings, contextWindowTokens: 4096 }, execCtx, callbacks, new AbortController().signal), /不足以容纳工具定义/);
  await assert.rejects(runAgent('值'.repeat(150000), [], settings, execCtx, callbacks, new AbortController().signal), /仍超过模型上下文预算/);
  assert.equal(requests, 0);
});

test('agent: context overflow compacts once, preserving the latest request without replaying writes', async () => {
  let requests = 0, executions = 0;
  const sent = [];
  const history = [
    { role: 'user', content: '旧'.repeat(18000) },
    { role: 'assistant', content: null, tool_calls: [call('committed')] },
    { role: 'tool', tool_call_id: 'committed', content: '{"ok":true,"data":"written"}' },
    { role: 'assistant', content: '已完成旧任务' },
  ];
  const { runAgent } = agentHarness(async (_settings, messages) => {
    sent.push(messages);
    if (requests++ === 0) throw Object.assign(new Error('context length exceeded'), { contextOverflow: true, contextLimit: 32000 });
    assert.ok(messages.some(m => m.role === 'user' && String(m.content).includes('继续统计')));
    return turn([], '统计完成');
  }, async () => { executions++; return { ok: true }; });
  await runAgent('继续统计', history, { ...settings, contextWindowTokens: 64000 }, execCtx, callbacks, new AbortController().signal);
  assert.equal(requests, 2);
  assert.equal(executions, 0);
  assert.ok(JSON.stringify(sent[1]).length < JSON.stringify(sent[0]).length);
  assert.equal(history[0].content.length, 18000);
  assert.equal(history.filter(m => m.tool_call_id === 'committed').length, 1);
});

test('agent: context overflow with no reducible history never loops', async () => {
  let requests = 0;
  const { runAgent } = agentHarness(async () => {
    requests++;
    throw Object.assign(new Error('context too long'), { contextOverflow: true });
  });
  await assert.rejects(runAgent('读取', [], settings, execCtx, callbacks, new AbortController().signal), /context too long/);
  assert.equal(requests, 1);
});

test('agent: UI failure preserves committed result and backfills only unexecuted calls', async () => {
  let writes = 0;
  const { runAgent } = agentHarness(async () => turn([call('a'), call('b')]), async () => { writes++; return { ok: true, data: 'written' }; });
  const history = [];
  await assert.rejects(runAgent('写入', history, settings, execCtx, { ...callbacks, onToolEnd() { throw new Error('UI failed'); } }, new AbortController().signal), /UI failed/);
  assert.equal(writes, 1);
  const results = history.filter(m => m.role === 'tool');
  assert.equal(results.length, 2);
  assert.equal(JSON.parse(results[0].content).ok, true);
  assert.equal(JSON.parse(results[1].content).ok, false);
});

test('agent: callback before tools still leaves complete tool history', async () => {
  const { runAgent } = agentHarness(async () => turn([call('a')]));
  const history = [];
  await assert.rejects(runAgent('写入', history, settings, execCtx, { ...callbacks, onAssistantDone() { throw new Error('UI failed'); } }, new AbortController().signal));
  assert.equal(history.at(-1).tool_call_id, 'a');
});

test('agent: native stream to dispatcher to next provider request closes the tool loop', async () => {
  const oldFetch = global.fetch;
  const requests = [];
  let writtenArgs;
  try {
    global.fetch = async (url, init) => {
      assert.equal(url, 'http://localhost:11434/v1/chat/completions');
      const body = JSON.parse(init.body);
      requests.push(body);
      if (requests.length === 1) return sse([delta({ tool_calls: [{ index: 0, ...call('a') }] }), delta({}, 'tool_calls')]);
      assert.equal(body.messages.at(-1).tool_call_id, 'a');
      assert.equal(JSON.parse(body.messages.at(-1).content).ok, true);
      return sse([delta({ content: '已写入 S!A1。' }, 'stop')]);
    };
    const runtime = sourceLoader({
      'src/core/tools/excelTools.ts': { TOOL_REGISTRY: {}, getQuickContext: async () => ({}), executeTool: async (_name, args, ctx) => { writtenArgs = JSON.parse(args); assert.ok(ctx.signal); return { ok: true, data: { rows: 1 } }; } },
      'src/core/skills.ts': { matchSkills: () => [] },
    })('src/core/agent.ts');
    const outcome = await runtime.runAgent('写入 11', [], settings, execCtx, callbacks, new AbortController().signal);
    assert.equal(requests.length, 2);
    assert.deepEqual(writtenArgs.values, [[11]]);
    assert.equal(outcome.incomplete, false);
  } finally { global.fetch = oldFetch; }
});

function steeringQueue() {
  const notes = [];
  return {
    notes,
    callbacks: { ...callbacks, hasSteering: () => notes.length > 0, takeSteering: () => notes.splice(0) },
  };
}
const userTexts = messages => messages.filter(m => m.role === 'user').map(m => typeof m.content === 'string' ? m.content : m.content.map(p => p.text || '').join(''));

test('agent: a note typed while a tool runs reaches the very next model request, after the tool result', async () => {
  const steer = steeringQueue();
  const requests = [];
  const { runAgent } = agentHarness(async (_s, messages) => {
    requests.push(messages);
    return requests.length === 1 ? turn([call('a')]) : turn([], '已按 B 列调整');
  }, async () => { steer.notes.push('改用 B 列'); return { ok: true, data: 'written' }; });
  const history = [];
  const outcome = await runAgent('写入', history, settings, execCtx, steer.callbacks, new AbortController().signal);
  assert.equal(requests.length, 2);
  const second = requests[1];
  const noteIndex = second.findIndex(m => m.role === 'user' && userTexts([m])[0].includes('[任务进行中用户补充] 改用 B 列'));
  const toolIndex = second.findIndex(m => m.role === 'tool' && m.tool_call_id === 'a');
  assert.ok(toolIndex > 0 && noteIndex > toolIndex, 'note must follow the tool result');
  assert.equal(steer.notes.length, 0);
  assert.equal(outcome.incomplete, false);
});

test('agent: a note that arrives while the model is finishing keeps the task going instead of being dropped', async () => {
  const steer = steeringQueue();
  const requests = [];
  const { runAgent } = agentHarness(async (_s, messages) => {
    requests.push(messages);
    if (requests.length === 1) { steer.notes.push('结果写到新表'); return turn([], '完成'); }
    return turn([], '已改为写到新表');
  });
  const history = [];
  await runAgent('汇总', history, settings, execCtx, steer.callbacks, new AbortController().signal);
  assert.equal(requests.length, 2);
  assert.ok(userTexts(requests[1]).some(t => t.includes('[任务进行中用户补充] 结果写到新表')));
  assert.deepEqual(history.filter(m => m.role === 'assistant').map(m => m.content), ['完成', '已改为写到新表']);
});

test('agent: truncated final answer is incomplete', async () => {
  const runtime = agentHarness(async () => ({ ...turn([], '结果'), truncated: true }));
  const outcome = await runtime.runAgent('分析', [], settings, execCtx, callbacks, new AbortController().signal);
  assert.equal(outcome.incomplete, true);
});

for (const arrival of ['model', 'tool']) {
  test(`agent: steering during ${arrival} skips stale pending calls before requesting a new plan`, async () => {
    const steer = steeringQueue();
    const requests = [];
    const executed = [];
    const { runAgent } = agentHarness(async (_s, messages) => {
      requests.push(messages);
      if (requests.length === 1) {
        if (arrival === 'model') steer.notes.push('不要覆盖原表，改写新表');
        return turn([call('a'), call('b')]);
      }
      return turn([], '已按新要求调整');
    }, async (_name, args) => {
      executed.push(args);
      steer.notes.push('不要覆盖原表，改写新表');
      return { ok: true, data: 'written' };
    });
    const history = [];
    await runAgent('写入', history, settings, execCtx, steer.callbacks, new AbortController().signal);
    assert.equal(executed.length, arrival === 'model' ? 0 : 1, 'old tool calls must not run after new instructions arrive');
    assert.equal(requests.length, 2);
    const results = requests[1].filter(m => m.role === 'tool');
    assert.deepEqual(results.map(m => m.tool_call_id), ['a', 'b'], 'every call still needs a protocol result');
    assert.equal(JSON.parse(results[0].content).ok, arrival === 'tool');
    assert.equal(JSON.parse(results[1].content).ok, false);
    assert.match(JSON.parse(results[1].content).error, /未执行/);
    assert.ok(userTexts(requests[1]).at(-1).includes('不要覆盖原表，改写新表'));
  });
}

test('agent: new steering supersedes the previous plan-only nudge', async () => {
  const steer = steeringQueue();
  const requests = [];
  const { runAgent } = agentHarness(async (_s, messages) => {
    requests.push(messages);
    if (requests.length === 1) {
      steer.notes.push('只解释步骤，先不要调用工具');
      return turn([], '我会先调用 read_range 读取数据，再继续处理。');
    }
    return turn([], '可以先按以下步骤检查。');
  });
  await runAgent('处理数据', [], settings, execCtx, steer.callbacks, new AbortController().signal);
  assert.equal(requests.length, 2);
  assert.match(userTexts(requests[1]).at(-1), /只解释步骤，先不要调用工具/);
  assert.ok(!userTexts(requests[1]).some(text => text.includes('现在请直接调用工具执行')));
});

test('batch: round limit and tool failures are never counted as done; emitted snapshots stay immutable', async () => {
  const snapshots = [];
  const runtime = sourceLoader({
    'src/core/agent.ts': { runAgent: async (_text, _history, _settings, _ctx, cb) => {
      cb.onToolStart('a', 'write_range', '写入', 'A1');
      cb.onToolEnd('a', { ok: false, error: 'failure' });
      return { hitRoundLimit: true };
    } },
    'src/core/tools/externalFiles.ts': { getExternalFiles: () => [{ name: 'large.csv', caption: 'large', sheets: [{ name: 'S', rows: Array.from({ length: 150000 }, () => [1]) }] }] },
  })('src/core/batch.ts');
  const result = await runtime.runBatch(['large.csv', 'missing.csv'], '汇总', settings, execCtx, s => snapshots.push(s), new AbortController().signal);
  assert.equal(result.states[0].status, 'incomplete');
  assert.equal(result.states[1].status, 'failed');
  assert.equal(snapshots.find(s => s.toolTrace.length).toolTrace[0].status, 'running');
});
