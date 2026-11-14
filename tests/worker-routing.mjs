import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHmac } from 'node:crypto';

const code = await readFile(new URL('../cloudflare/workers.js', import.meta.url), 'utf8');
const workerModule = await import('data:text/javascript;base64,' + Buffer.from(code).toString('base64'));
const worker = workerModule.default;
const menuManifest = JSON.parse(await readFile(new URL('../rich-menu/menu.json',import.meta.url),'utf8'));
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
  assert.deepEqual(workerModule.parseManagerPostbackData('action=manager_shift&id='+id+'&v=2&decision=resend'), {id,version:2,decision:'resend'});
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
function bindingWebhook(replyToken, text = '測試員工\tABCDEF123456') {
  const body = JSON.stringify({events: [{
    type: 'message', source: {type: 'user', userId: bindingUid}, replyToken,
    webhookEventId: 'binding-event', message: {type: 'text', text}
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
    assert.equal(calls[0].body.text, '測試員工\tABCDEF123456');
    assert.equal(calls[1].url, 'https://api.line.me/v2/bot/message/reply');
    assert.equal(calls[1].body.replyToken, 'first-reply');
    assert.equal(calls[1].body.messages[0].text, bindingMessage);
  } finally { globalThis.fetch = oldFetch; }
});

test('unrelated LINE text does not start a binding request', async () => {
  const oldFetch = globalThis.fetch;
  globalThis.fetch = async (url,init) => {
    assert.equal(url,env.GAS_WEB_APP_URL);
    assert.equal(JSON.parse(init.body).kind,'contact_message');
    return Response.json({ok:true,code:'ignored'});
  };
  try {
    const response = await worker.fetch(bindingWebhook('chat', '您好，想問明天的班表'), env);
    assert.equal(response.status, 200);
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


test('signed acceptance preserves the Flex card and follow-ups are routed to GAS',async()=>{
  const id='11111111-1111-1111-1111-111111111111';
  const userId='U'+'a'.repeat(32);
  const card={type:'flex',altText:'已同意排班',contents:{type:'bubble'}};
  const old=globalThis.fetch,calls=[];
  globalThis.fetch=async(url,init)=>{
    const data=JSON.parse(init.body); calls.push(data);
    return new Response(JSON.stringify(url===env.GAS_WEB_APP_URL
      ? {ok:true,message:'已同意排班',...(data.kind==='shift_reply'?{replyMessage:card}:{})}
      : {}));
  };
  try {
    const events=[
      {type:'postback',postback:{data:'action=shift_reply&id='+id+'&v=1&choice=agree'}},
      {type:'postback',postback:{data:'action=shift_followup&id='+id+'&v=1&choice=leave_sick'}},
      {type:'message',message:{type:'text',text:'排班留言 ABCDEF01 測試'}}
    ];
    for (const event of events) {
      const body=JSON.stringify({events:[{...event,source:{type:'user',userId},replyToken:'reply'}]});
      const signature=createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64');
      const response=await worker.fetch(new Request('https://worker.example/webhook',{
        method:'POST',body,headers:{'x-line-signature':signature}
      }),env);
      assert.equal(response.status,200);
    }
    assert.deepEqual(calls[1].messages,[card]);
    assert.equal(calls[2].kind,'shift_followup');
    assert.equal(calls[2].choice,'leave_sick');
    assert.equal(calls[4].kind,'shift_followup');
    assert.equal(calls[4].text,'排班留言 ABCDEF01 測試');
  } finally {globalThis.fetch=old;}
});


test('a rejected receipt card falls back to text with the same action buttons',async()=>{
  const id='11111111-1111-1111-1111-111111111111';
  const source={type:'user',userId:'U'+'a'.repeat(32)};
  const body=JSON.stringify({events:[{type:'postback',source,replyToken:'reply',
    postback:{data:'action=shift_reply&id='+id+'&v=1&choice=agree'}}]});
  const quickReply={items:[{type:'action',action:{type:'postback',label:'我要請假',data:'test'}}]};
  const old=globalThis.fetch,replies=[];
  globalThis.fetch=async(url,init)=>{
    if(url===env.GAS_WEB_APP_URL) return new Response(JSON.stringify({ok:true,serial:'A-000001',
      message:'已同意排班\n流水號：A-000001',quickReply,replyMessage:{type:'flex'}}));
    replies.push(JSON.parse(init.body).messages[0]);
    return new Response('{}',{status:replies.length===1?400:200});
  };
  try {
    const signature=createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64');
    const response=await worker.fetch(new Request('https://worker.example/webhook',{
      method:'POST',body,headers:{'x-line-signature':signature}
    }),env);
    assert.equal(response.status,200);
    assert.deepEqual(replies[1],{type:'text',text:'已同意排班\n流水號：A-000001',quickReply});
  } finally {globalThis.fetch=old;}
});

test('approved-card actions use the GAS delivery and signal retry when both paths fail',async()=>{
  const id='11111111-1111-1111-1111-111111111111';
  const source={type:'user',userId:'U'+'a'.repeat(32)};
  const old=globalThis.fetch,calls=[];
  let replyStatus=200;
  globalThis.fetch=async(url,init)=>{
    const data=JSON.parse(init.body);
    calls.push(data);
    if(url!==env.GAS_WEB_APP_URL) return new Response('{}',{status:replyStatus});
    return Response.json(data.choice==='contact' ? {ok:true,code:'contact_open',delivered:true}
      : {ok:true,code:data.choice,message:'下一步',quickReply:{items:[]}});
  };
  const click=choice=>{
    const body=JSON.stringify({events:[{type:'postback',source,replyToken:'reply',webhookEventId:'evt-'+choice,
      postback:{data:'action=shift_followup&id='+id+'&v=1&choice='+choice}}]});
    const signature=createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64');
    return worker.fetch(new Request('https://worker.example/webhook',{
      method:'POST',body,headers:{'x-line-signature':signature}
    }),env);
  };
  try {
    assert.equal((await click('contact')).status,200);
    assert.equal(calls.length,1);
    assert.equal(calls[0].eventId,'evt-contact');
    assert.equal((await click('swap')).status,200);
    assert.equal(calls[2].messages[0].text,'下一步');
    replyStatus=401;
    assert.equal((await click('leave')).status,502);
    assert.deepEqual(calls.filter(call=>call.kind).map(call=>call.choice),['contact','swap','leave']);
  } finally {globalThis.fetch=old;}
});

test('signed contact, swap and leave review events reach the correct GAS handlers',async()=>{
  const id='11111111-1111-1111-1111-111111111111';
  const source={type:'user',userId:'U'+'a'.repeat(32)};
  const events=[
    {type:'message',message:{type:'text',text:'集合時間？',id:'line-message'}},
    {type:'postback',postback:{data:'action=shift_followup&id='+id+'&v=1&choice=contact_end'}},
    {type:'postback',postback:{data:'action=shift_followup&id='+id+'&v=1&choice=swap_date',params:{date:'2026-11-11'}}},
    {type:'postback',postback:{data:'action=shift_followup&id='+id+'&v=1&choice=swap&page=1'}},
    {type:'postback',postback:{data:'action=swap_response&id='+id+'&decision=agree'}},
    {type:'message',message:{type:'text',text:'審核請假 l-000001',id:'legacy-leave'}},
    {type:'postback',postback:{data:'action=leave_review&serial=L-000001&decision=approve'}},
    {type:'postback',postback:{data:'action=leave_review&serial=L-000001&decision=reject'}}
  ];
  const body=JSON.stringify({events:events.map(event=>({...event,source,replyToken:'reply',timestamp:123}))});
  const signature=createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64');
  const old=globalThis.fetch,calls=[];
  globalThis.fetch=async(url,init)=>{
    assert.equal(url,env.GAS_WEB_APP_URL);
    calls.push(JSON.parse(init.body));
    return Response.json({ok:true,code:'ignored'});
  };
  try {
    const response=await worker.fetch(new Request('https://worker.example/webhook',{
      method:'POST',body,headers:{'x-line-signature':signature}
    }),env);
    assert.equal(response.status,200);
    assert.equal(calls.find(call=>call.kind==='contact_message').eventId,'line-message');
    assert.equal(calls.find(call=>call.choice==='contact_end').timestamp,123);
    assert.equal(calls.find(call=>call.choice==='swap_date').date,'2026-11-11');
    assert.equal(calls.find(call=>call.choice==='swap').page,1);
    assert.equal(calls.find(call=>call.kind==='swap_response').decision,'agree');
    assert.deepEqual(calls.filter(call=>call.kind==='leave_review').map(call=>[call.serial,call.decision]),
      [['L-000001',undefined],['L-000001','approve'],['L-000001','reject']]);
  } finally {globalThis.fetch=old;}
});

test('all rich menu shortcuts and new reply buttons reach their respective routes',async()=>{
  const source={type:'user',userId:bindingUid};
  const token='11111111-1111-1111-1111-111111111111';
  const routes=[];
  const areas=menuManifest.menu.areas;
  const {width,height}=menuManifest.menu.size;
  for (const {bounds,action} of areas) {
    assert.ok(bounds.x>=0 && bounds.y>=0 && bounds.width>0 && bounds.height>0 &&
      bounds.x+bounds.width<=width && bounds.y+bounds.height<=height);
    if (action.type==='uri') assert.equal(new URL(action.uri).hostname,'liff.line.me');
    else routes.push({type:'message',message:{type:'text',text:action.text}});
  }
  for (const data of [
    'action=leave_replacement&serial=L-000001&decision=skip',
    'action=leave_replacement&serial=L-000001&decision=select&staff=A04',
    'action=calendar_menu&choice=leave_pick&value=0',
    'action=calendar_menu&choice=contact_reply&value='+token,
    'action=vacancy&choice=reject&id='+token
  ]) routes.push({type:'postback',postback:{data}});
  routes.push({type:'message',message:{type:'text',text:'團員回覆 '+token+' 測試回覆'}});
  const body=JSON.stringify({events:routes.map((event,index)=>({...event,source,
    replyToken:'reply-'+index,webhookEventId:'event-'+index}))});
  const signature=createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64');
  const old=globalThis.fetch, calls=[], receipts=[];
  globalThis.fetch=async(url,init)=>{
    const payload=JSON.parse(init.body);
    if (url===env.GAS_WEB_APP_URL) {
      calls.push(payload);
      return Response.json({ok:true,code:'done',message:payload.kind==='vacancy' && payload.choice==='list'?'現已滿班':'操作結果'});
    }
    assert.equal(url,'https://api.line.me/v2/bot/message/reply');
    receipts.push(payload);
    return Response.json({});
  };
  try {
    const response=await worker.fetch(new Request('https://worker.example/webhook',{
      method:'POST',body,headers:{'x-line-signature':signature}}),env);
    assert.equal(response.status,200);
    assert.deepEqual(calls.map(call=>[call.kind,call.choice || call.decision]),[
      ['calendar_menu_action','contact'],['calendar_menu_action','leave'],['calendar_menu_action','replies'],
      ['vacancy','list'],['personal_calendar_menu',undefined],['leave_replacement','skip'],
      ['leave_replacement','select'],['calendar_menu_action','leave_pick'],
      ['calendar_menu_action','contact_reply'],['vacancy','reject'],['calendar_menu_action','reply_send']
    ]);
    assert.equal(receipts.length,routes.length);
    assert.deepEqual(receipts[3].messages,[{type:'text',text:'現已滿班'}]);
  } finally {globalThis.fetch=old;}
});
