const assert = require('node:assert/strict');
const { test } = require('node:test');
const { sourceLoader } = require('./load-source.cjs');
const { explainRunError, providerMessage } = sourceLoader()('src/core/errorExplain.ts');

const http = (status, body, extra = {}) => Object.assign(new Error(`模型请求失败 (HTTP ${status}): ${body}`), { status, ...extra });

test('errors: provider JSON is reduced to one readable sentence', () => {
  assert.equal(providerMessage('{"error":{"message":"Incorrect API key provided","type":"auth"}}'), 'Incorrect API key provided');
  assert.equal(providerMessage('{"error":"model not found"}'), 'model not found');
  assert.equal(providerMessage('{"msg":"余额不足"}'), '余额不足');
  assert.equal(providerMessage('<html><body>502 Bad Gateway</body></html>'), '');
  assert.equal(providerMessage('plain text reason'), 'plain text reason');
  assert.ok(providerMessage('x'.repeat(900)).length <= 221);
});

test('errors: auth, balance, permission and missing model point at settings without leaking raw JSON', () => {
  const auth = explainRunError(http(401, '{"error":{"message":"Incorrect API key provided"}}'));
  assert.match(auth.text, /API Key 无效或已过期/);
  assert.equal(auth.action, 'settings');
  assert.equal(auth.detail, '服务商返回（HTTP 401）：Incorrect API key provided');
  assert.ok(!auth.text.includes('{'));

  const balance = explainRunError(http(402, '{"error":{"message":"Insufficient Balance"}}'));
  assert.match(balance.text, /余额不足/);
  // 有的服务商用 400/429 报欠费，不能说成「限流」
  assert.match(explainRunError(http(429, '{"error":{"message":"You exceeded your current quota"}}')).text, /余额不足/);
  assert.match(explainRunError(http(403, '{"error":{"message":"no access to model"}}')).text, /没有权限/);
  assert.match(explainRunError(http(404, '{"error":{"message":"The model `x` does not exist"}}')).text, /模型名不存在/);
  assert.match(explainRunError(http(400, '{"error":{"message":"this model does not support tools"}}')).text, /不支持工具调用/);
});

test('errors: overflow suggests a new chat; rate limits and outages only offer retry', () => {
  const overflow = explainRunError(http(400, '{"error":{"message":"maximum context length is 65536 tokens"}}', { contextOverflow: true }));
  assert.equal(overflow.action, 'new-chat');
  assert.match(overflow.text, /上下文窗口/);
  const limited = explainRunError(http(429, '{"error":{"message":"Rate limit reached"}}'));
  assert.match(limited.text, /限流/);
  assert.equal(limited.action, undefined);
  const down = explainRunError(http(502, '<html>Bad Gateway</html>'));
  assert.match(down.text, /暂时不可用/);
  assert.equal(down.detail, 'HTTP 502');
  assert.equal(down.action, undefined);
});

test('errors: network failures name the likely cause for remote and local endpoints', () => {
  const remote = explainRunError(new TypeError('Failed to fetch'));
  assert.match(remote.text, /连不上模型服务/);
  assert.equal(remote.action, 'settings');
  assert.match(explainRunError(new TypeError('Failed to fetch'), { loopback: true }).text, /Ollama/);
});

test('errors: settings-panel formats and already-readable messages pass through', () => {
  assert.match(explainRunError(new Error('HTTP 401: {"error":{"message":"bad key"}}')).text, /API Key 无效/);
  assert.match(explainRunError(new Error('获取模型列表失败 (HTTP 403): forbidden')).text, /没有权限/);
  const readable = explainRunError(new Error('模型响应在完成前中断，请重试；本轮工具尚未执行。'));
  assert.deepEqual(readable, { text: '模型响应在完成前中断，请重试；本轮工具尚未执行。' });
  assert.equal(explainRunError(new Error('压缩历史后，当前任务和最近工具结果仍超过模型上下文预算。')).action, 'new-chat');
  assert.ok(explainRunError(undefined).text);
});
