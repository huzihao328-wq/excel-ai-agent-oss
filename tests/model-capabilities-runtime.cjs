const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const load = sourceLoader();
const { DEFAULT_SETTINGS } = load('src/core/settings.ts');
const settings = { ...DEFAULT_SETTINGS, baseUrl: 'https://gateway.example/v1', model: 'custom-model', apiKey: 'test-only' };
const json = (body, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
const answer = (message, finish_reason = 'stop') => json({ choices: [{ message, finish_reason }] });

test('context: output limits and unrelated models never masquerade as the context window', () => {
  const { pickContextTokens, contextLimitFromError, guessContextWindow } = load('src/core/contextWindow.ts');
  assert.equal(pickContextTokens({ max_tokens: 32768, max_output_tokens: 65536 }), null);
  assert.equal(pickContextTokens({ max_completion_tokens: 32768 }), null);
  assert.equal(pickContextTokens({ context_length: 8192, model_info: { 'llama.context_length': 131072 } }), 8192);
  assert.equal(pickContextTokens({ context_length: 8192, context_window: 32768 }), 8192);
  assert.equal(pickContextTokens({ data: [{ id: 'other', context_length: 999000 }] }), null);
  assert.equal(pickContextTokens({ metadata: { max_input_tokens: '32k' } }), 32768);
  assert.equal(contextLimitFromError('maximum context length is 16,384 tokens, requested 23000'), 16384);
  assert.equal(contextLimitFromError('max_tokens is 8192; requested 23000 tokens'), null);
  assert.ok(guessContextWindow('gpt-5-custom') < 1000000);
  assert.equal(guessContextWindow('moonshot-v1-8k'), 8192);
});

test('context: cached evidence expires and is isolated by endpoint, model and credential', () => {
  const ctx = sourceLoader()('src/core/contextWindow.ts');
  const now = Date.now;
  try {
    Date.now = () => 1000;
    ctx.rememberContextWindow('https://host/v1/', 'm', 8192, 'detail', 'one');
    assert.equal(ctx.getCachedContextWindow('https://host/v1', 'm', 'one'), 8192);
    assert.equal(ctx.getCachedContextWindow('https://host/v1', 'm', 'two'), undefined);
    assert.equal(ctx.getCachedContextWindow('https://other/v1', 'm', 'one'), undefined);
    ctx.rememberContextWindows('https://host/v1', [{ id: 'm', contextWindow: 131072 }], 'one');
    assert.equal(ctx.getCachedContextWindow('https://host/v1', 'm', 'one'), 8192);
    Date.now = () => 302000;
    assert.equal(ctx.getCachedContextWindow('https://host/v1', 'm', 'one'), undefined);
  } finally { Date.now = now; }
});

test('context: compaction preserves the original task together with execution-time steering', () => {
  const { compactHistoryForRequest } = load('src/core/contextWindow.ts');
  const history = [
    { role: 'user', content: '旧任务'.repeat(4000) }, { role: 'assistant', content: '旧任务完成' },
    { role: 'user', content: '当前任务：保留完整说明'.repeat(1000) },
    { role: 'assistant', content: null, tool_calls: [{ id: 'one', type: 'function', function: { name: 'read_range', arguments: '{}' } }] },
    { role: 'tool', tool_call_id: 'one', content: '{"ok":true}' },
    { role: 'user', content: '[任务进行中用户补充] 结果写入新表' },
  ];
  const result = compactHistoryForRequest(history, { budgetTokens: 2000 });
  assert.ok(result.some(m => m.content === history[2].content));
  assert.ok(result.some(m => m.content === history[5].content));
  assert.equal(result.filter(m => m.role === 'tool').length, 1);
});

test('context: cancellation aborts an in-flight probe and prevents fallback requests', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  let requests = 0;
  const ctrl = new AbortController();
  global.fetch = async (_url, init) => {
    requests++;
    return new Promise((_resolve, reject) => init.signal.addEventListener('abort', () => reject(new DOMException('aborted', 'AbortError')), { once: true }));
  };
  const pending = sourceLoader()('src/core/llm.ts').probeContextWindowInfo(settings, { signal: ctrl.signal });
  ctrl.abort();
  await assert.rejects(pending, error => error.name === 'AbortError');
  assert.equal(requests, 1);
});

test('context: force probe bypasses cached estimates and reads the new deployment limit', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  const local = sourceLoader();
  const ctx = local('src/core/contextWindow.ts');
  ctx.rememberContextWindow(settings.baseUrl, settings.model, 131072, 'estimate', settings.apiKey);
  let requests = 0;
  global.fetch = async url => {
    requests++;
    assert.ok(url.endsWith('/models'));
    return json({ data: [{ id: 'other', context_length: 999000 }, { id: settings.model, context_length: 8192 }] });
  };
  const llm = local('src/core/llm.ts');
  assert.equal((await llm.probeContextWindowInfo(settings)).source, 'estimate');
  assert.equal(requests, 0);
  assert.deepEqual(await llm.probeContextWindowInfo(settings, { force: true }), { tokens: 8192, source: 'catalog' });
  assert.equal(requests, 1);
});

test('context: estimates remain labelled and failed probes do not invent a window', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  global.fetch = async () => json({}, 404);
  const { probeContextWindowInfo } = sourceLoader()('src/core/llm.ts');
  assert.equal(await probeContextWindowInfo(settings), null);
  const info = await probeContextWindowInfo({ ...settings, model: 'deepseek-chat' });
  assert.deepEqual(info, { tokens: 131072, source: 'estimate' });
});

test('context: Ollama uses the matching running model before any training limit', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  const requests = [];
  global.fetch = async url => {
    requests.push(url);
    return json({ models: [{ name: 'unrelated', context_length: 999000 }, { name: 'qwen:latest', context_length: 4096 }] });
  };
  const info = await sourceLoader()('src/core/llm.ts').probeContextWindowInfo({ ...settings, baseUrl: 'http://localhost:11434/v1', model: 'qwen' });
  assert.deepEqual(info, { tokens: 4096, source: 'ollama-running' });
  assert.deepEqual(requests, ['http://localhost:11434/api/ps']);
});

test('context: Ollama num_ctx takes precedence; training-only metadata stays unknown', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  let parameters = 'temperature 0.2\nnum_ctx 8192';
  global.fetch = async (url, init) => {
    if (url.endsWith('/api/ps')) return json({ models: [] });
    if (url.endsWith('/api/show')) {
      assert.deepEqual(JSON.parse(init.body), { model: 'qwen' });
      return json({ parameters, model_info: { 'qwen.context_length': 131072 } });
    }
    return json({}, 404);
  };
  const llm = sourceLoader()('src/core/llm.ts');
  const config = { ...settings, baseUrl: 'http://localhost:11434/v1', model: 'qwen' };
  assert.deepEqual(await llm.probeContextWindowInfo(config), { tokens: 8192, source: 'ollama-config' });
  parameters = '';
  assert.equal(await llm.probeContextWindowInfo(config, { force: true }), null);
});

for (const mode of ['native', 'text', 'wrong-arguments', 'wrong-receipt', 'unsupported', 'truncated']) {
  test(`connection: virtual tool round trip - ${mode}`, async t => {
    const previous = global.fetch;
    t.after(() => { global.fetch = previous; });
    let calls = 0;
    global.fetch = async (_url, init) => {
      const body = JSON.parse(init.body);
      assert.equal(body.tools.length, 1);
      assert.equal(body.tools[0].function.name, 'connection_probe');
      calls++;
      if (mode === 'unsupported') {
        // 还没确认接口接受 stream_options 时，首个 400 会去掉它重发一次；仍被拒绝才报未验证
        assert.equal('stream_options' in body, calls === 1);
        return json({ error: { message: 'tools not supported' } }, 400);
      }
      if (calls === 1) {
        assert.ok(!JSON.stringify(body).includes('sheet'));
        if (mode === 'text') return answer({ content: '<tool_call>{"name":"connection_probe","arguments":{"value":7}}</tool_call>' });
        return answer({ tool_calls: [{ id: 'p', type: 'function', function: { name: 'connection_probe', arguments: JSON.stringify({ value: mode === 'wrong-arguments' ? '7' : 7 }) } }] }, mode === 'truncated' ? 'length' : 'tool_calls');
      }
      const result = body.messages.at(-1);
      assert.equal(result.role, 'tool');
      assert.equal(result.tool_call_id, body.messages.at(-2).tool_calls[0].id);
      const receipt = JSON.parse(result.content).receipt;
      assert.ok(!body.messages[0].content.includes(receipt));
      return answer({ content: mode === 'wrong-receipt' ? 'pong' : receipt });
    };
    const result = await sourceLoader()('src/core/llm.ts').probeToolCalling(settings);
    assert.equal(result.status, mode === 'native' || mode === 'text' ? mode : 'unverified');
    assert.equal(calls, ['wrong-arguments', 'truncated'].includes(mode) ? 1 : 2);
  });
}

test('connection: successful plain chat cannot hide a failed tool capability test', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  global.fetch = async (url, init) => {
    if (!url.endsWith('/chat/completions')) return json({ data: [{ id: settings.model, context_length: 32000 }] });
    const body = JSON.parse(init.body);
    return body.tools ? json({ error: { message: 'tools not supported' } }, 400) : answer({ content: 'pong' });
  };
  const result = await sourceLoader()('src/core/llm.ts').testConnection(settings);
  assert.equal(result.toolCalling.status, 'unverified');
  assert.equal(result.contextInfo.source, 'catalog');
  assert.match(result.message, /对话连接成功.*工具调用未验证/s);
});

test('request: explicit context errors expose a recoverable limit and cache the evidence', async t => {
  const previous = global.fetch;
  t.after(() => { global.fetch = previous; });
  global.fetch = async () => json({ error: { message: 'maximum context length is 16,384 tokens, requested 23000' } }, 400);
  const local = sourceLoader();
  await assert.rejects(local('src/core/llm.ts').streamChatCompletion(settings, [], [], {}, new AbortController().signal), error => {
    assert.equal(error.contextLimit, 16384);
    assert.equal(error.contextOverflow, true);
    return true;
  });
  assert.deepEqual(local('src/core/contextWindow.ts').getCachedContextInfo(settings.baseUrl, settings.model, settings.apiKey), { tokens: 16384, source: 'error' });
});
