import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHmac } from 'node:crypto';

const code = await readFile(new URL('../cloudflare/workers.js', import.meta.url), 'utf8');
const workerModule = await import('data:text/javascript;base64,' + Buffer.from(code).toString('base64'));
const worker = workerModule.default;
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

test('employee and manager postback actions stay distinct', () => {
  const id = '11111111-1111-1111-1111-111111111111';
  assert.deepEqual(workerModule.parsePostbackData('action=shift_reply&id='+id+'&v=2&choice=agree'), {id,version:2,choice:'agree'});
  assert.equal(workerModule.parsePostbackData('action=manager_shift&id='+id+'&v=2&decision=skip'), null);
  assert.deepEqual(workerModule.parseManagerPostbackData('action=manager_shift&id='+id+'&v=2&decision=replace'), {id,version:2,decision:'replace'});
  assert.equal(workerModule.parseManagerPostbackData('action=manager_shift&id='+id+'&v=2&decision=delete'), null);
});

test('signed replacement webhook forwards selection and preserves quick reply buttons', async()=>{
  const id='11111111-1111-1111-1111-111111111111';
  const body=JSON.stringify({events:[{type:'postback',replyToken:'reply',source:{type:'user',userId:'U'+'a'.repeat(32)},postback:{data:'action=manager_shift&id='+id+'&v=1&decision=select&staff=A05'}}]});
  const quickReply={items:[{type:'action',action:{type:'postback',label:'人員甲',data:'example'}}]};
  const old=globalThis.fetch,calls=[];
  globalThis.fetch=async(url,init)=>{
    const data=JSON.parse(init.body); calls.push(data);
    return new Response(JSON.stringify(url===env.GAS_WEB_APP_URL?{ok:true,message:'請選擇替補對象',quickReply}:{}));
  };
  try {
    const response=await worker.fetch(new Request('https://worker.example/webhook',{method:'POST',body,headers:{'x-line-signature':createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64')}}),env);
    assert.equal(response.status,200);
    assert.equal(calls[0].staffId,'A05');
    assert.equal(calls[0].kind,'manager_decision');
    assert.deepEqual(calls[1].messages[0].quickReply,quickReply);
  } finally {globalThis.fetch=old;}
});

const bindingUid = 'U' + 'a'.repeat(32);
const bindingMessage = '測試員工您好，LINE 帳號綁定成功！';
function bindingWebhook(replyToken) {
  const body = JSON.stringify({events: [{
    type: 'message', source: {type: 'user', userId: bindingUid}, replyToken,
    webhookEventId: 'binding-event', message: {type: 'text', text: '綁定 ABCDEF123456'}
  }]});
  const signature = createHmac('sha256', env.LINE_CHANNEL_SECRET).update(body).digest('base64');
  return new Request('https://worker.example/webhook', {method: 'POST', body, headers: {'x-line-signature': signature}});
}

test('first binding replies to the originating LINE conversation when reply is accepted', async () => {
  const oldFetch = globalThis.fetch;
  const work = [], calls = [];
  globalThis.fetch = async (url, init) => {
    calls.push({url, body: JSON.parse(init.body)});
    return new Response(JSON.stringify(url === env.GAS_WEB_APP_URL
      ? {ok: true, code: 'bound', message: bindingMessage} : {}));
  };
  try {
    const response = await worker.fetch(bindingWebhook('first-reply'), env, {waitUntil: promise => work.push(promise)});
    await Promise.all(work);
    assert.equal(response.status, 200);
    assert.equal(work.length, 1);
    assert.equal(calls.length, 2);
    assert.equal(calls[0].body.lineUid, bindingUid);
    assert.equal(calls[1].url, 'https://api.line.me/v2/bot/message/reply');
    assert.equal(calls[1].body.replyToken, 'first-reply');
    assert.equal(calls[1].body.messages[0].text, bindingMessage);
  } finally { globalThis.fetch = oldFetch; }
});

test('rejected first reply is privately pushed to the same user with an idempotent retry', async () => {
  const oldFetch = globalThis.fetch;
  const pushes = [];
  globalThis.fetch = async (url, init) => {
    if (url === env.GAS_WEB_APP_URL) return Response.json({ok: true, code: 'bound', message: bindingMessage});
    if (url.endsWith('/reply')) return new Response('', {status: 400});
    assert.equal(url, 'https://api.line.me/v2/bot/message/push');
    pushes.push({body: JSON.parse(init.body), key: init.headers['x-line-retry-key']});
    return pushes.length === 1 ? new Response('', {status: 500})
      : new Response('', {status: 409, headers: {'x-line-accepted-request-id': 'accepted'}});
  };
  try {
    assert.equal((await worker.fetch(bindingWebhook('expired'), env)).status, 200);
    assert.equal(pushes.length, 2);
    assert.equal(pushes[0].body.to, bindingUid);
    assert.equal(pushes[0].body.messages[0].text, bindingMessage);
    assert.match(pushes[0].key, /^[0-9a-f-]{36}$/);
    assert.deepEqual(pushes[1], pushes[0]);
  } finally { globalThis.fetch = oldFetch; }
});

test('an already-bound or rejected binding never sends the successful binding push', async () => {
  const oldFetch = globalThis.fetch;
  const calls = [];
  let code = 'already_bound';
  globalThis.fetch = async url => {
    calls.push(url);
    if (url === env.GAS_WEB_APP_URL) return Response.json({ok: true, code, message: '綁定狀態'});
    return new Response('', {status: 400});
  };
  try {
    assert.equal((await worker.fetch(bindingWebhook('second-reply'), env)).status, 200);
    code = 'invalid_code';
    assert.equal((await worker.fetch(bindingWebhook(), env)).status, 200);
    assert.deepEqual(calls, [env.GAS_WEB_APP_URL, 'https://api.line.me/v2/bot/message/reply', env.GAS_WEB_APP_URL]);
  } finally { globalThis.fetch = oldFetch; }
});
