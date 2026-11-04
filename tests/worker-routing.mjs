import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const code = await readFile(new URL('../cloudflare/workers.js', import.meta.url), 'utf8');
const worker = (await import('data:text/javascript;base64,' + Buffer.from(code).toString('base64'))).default;
const origin = 'https://enddenmi128.github.io';
const env = {
  GAS_SHARED_SECRET: 'test-secret',
  GAS_WEB_APP_URL: 'https://example.invalid/exec',
  LINE_CHANNEL_SECRET: 'test-line-channel-secret',
  LINE_CHANNEL_ACCESS_TOKEN: 'test-line-channel-token'
};

function apiRequest(action, overrides = {}) {
  return new Request('https://worker.example/calendar-api', {
    method: 'POST',
    headers: {
      origin: overrides.origin ?? origin,
      authorization: overrides.authorization ?? 'Bearer test-access-token',
      'content-type': 'application/json'
    },
    body: JSON.stringify({ action, payload: {} })
  });
}

test('calendar bootstrap and managerSession forward the verified LINE token', async () => {
  const oldFetch = globalThis.fetch;
  const seen = [];
  globalThis.fetch = async (url, init) => {
    assert.equal(url, env.GAS_WEB_APP_URL);
    seen.push(JSON.parse(init.body));
    return new Response(JSON.stringify({ ok: true, viewer: { id: 'A03', manager: true } }), { status: 200 });
  };
  try {
    for (const action of ['bootstrap', 'managerSession']) {
      const response = await worker.fetch(apiRequest(action), env);
      assert.equal(response.status, 200);
      assert.equal((await response.json()).ok, true);
    }
    assert.deepEqual(seen.map(v => v.action), ['bootstrap', 'managerSession']);
    for (const body of seen) {
      assert.equal(body.kind, 'calendar_api');
      assert.equal(body.lineAccessToken, 'test-access-token');
      assert.equal(body.secret, env.GAS_SHARED_SECRET);
      assert.equal(Object.hasOwn(body, 'adminKey'), false);
    }
  } finally {
    globalThis.fetch = oldFetch;
  }
});

test('calendar API rejects invalid actions, origins and missing identity', async () => {
  const invalid = await worker.fetch(apiRequest('unsupported'), env);
  assert.equal(invalid.status, 400);
  const forbidden = await worker.fetch(apiRequest('bootstrap', { origin: 'https://evil.example' }), env);
  assert.equal(forbidden.status, 403);
  const identity = await worker.fetch(apiRequest('bootstrap', { authorization: 'Bearer ' }), env);
  assert.equal(identity.status, 401);
});

test('health endpoint and webhook signature guard remain intact', async () => {
  const health = await worker.fetch(new Request('https://worker.example/health'), env);
  assert.equal(health.status, 200);
  assert.equal(await health.text(), 'ok');
  const root = await worker.fetch(new Request('https://worker.example/'), env);
  assert.equal(root.status, 404);
  const webhook = await worker.fetch(new Request('https://worker.example/webhook', {
    method: 'POST',
    body: JSON.stringify({ events: [] }),
    headers: { 'content-type': 'application/json' }
  }), env);
  assert.equal(webhook.status, 401);
});
