/**
 * 團長排班通知｜Cloudflare Worker（ES Module）
 * LINE 的 Webhook URL 必須填 https://你的-worker.workers.dev/webhook
 * 驗證原始 Body 的 x-line-signature，再安全轉送 Google Apps Script。
 */

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/calendar-api') return handleCalendarApi(request, env);
    if (request.method === 'GET' && url.pathname === '/health') {
      return new Response('ok', {status: 200, headers: {'content-type': 'text/plain; charset=utf-8'}});
    }
    if (request.method !== 'POST' || url.pathname !== '/webhook') {
      return new Response('Not found', {status: 404});
    }
    if (!env.LINE_CHANNEL_SECRET || !env.LINE_CHANNEL_ACCESS_TOKEN || !env.GAS_SHARED_SECRET || !env.GAS_WEB_APP_URL) {
      console.error('Required Worker environment variable missing');
      return new Response('Service unavailable', {status: 503});
    }

    const originalBody = await request.text();
    const signature = request.headers.get('x-line-signature');
    if (!await verifyLineSignature(originalBody, signature, env.LINE_CHANNEL_SECRET)) {
      return new Response('Invalid LINE signature', {status: 401});
    }

    let envelope;
    try { envelope = JSON.parse(originalBody); }
    catch { return new Response('Invalid JSON', {status: 400}); }
    if (!Array.isArray(envelope.events)) return new Response('Invalid events', {status: 400});

    // 有意等待 GAS 寫回結果，再回傳 200。後端暫時故障時回 502，搭配 LINE Webhook redelivery 重試。
    // GAS 的身分與版本檢查採冪等設計，避免 webhook 重送造成重複變更。
    const results = await Promise.allSettled(envelope.events.map(event => handleEvent(event, env)));
    if (results.some(result => result.status === 'rejected')) {
      for (const result of results) {
        if (result.status === 'rejected') console.error('Webhook processing failed:', String(result.reason?.message || result.reason));
      }
      return new Response('Temporary processing error', {status: 502});
    }
    return new Response('OK', {status: 200});
  }
};

/** LIFF calendar requests use a verified LINE access token; webhook handling is unchanged. */
export async function handleCalendarApi(request, env) {
  const allowedOrigin = env.CALENDAR_ALLOWED_ORIGIN || 'https://enddenmi128.github.io';
  if (request.headers.get('origin') !== allowedOrigin) {
    return new Response('Forbidden origin', {status: 403});
  }
  const headers = {
    'access-control-allow-origin': allowedOrigin,
    'access-control-allow-methods': 'POST, OPTIONS',
    'access-control-allow-headers': 'content-type, authorization',
    'cache-control': 'no-store', 'vary': 'Origin',
    'content-type': 'application/json; charset=utf-8'
  };
  const json = (value, status=200) => new Response(JSON.stringify(value), {status, headers});
  if (request.method === 'OPTIONS') return new Response(null, {status: 204, headers});
  if (request.method !== 'POST') return json({ok:false,error:'Method not allowed'},405);
  if (!env.GAS_SHARED_SECRET || !env.GAS_WEB_APP_URL) return json({ok:false,error:'Worker 尚未設定 GAS 連線'},503);
  const auth = request.headers.get('authorization') || '';
  const tokenMatch = /^Bearer\s+([^\s]+)$/i.exec(auth);
  if (!tokenMatch || tokenMatch[1].length > 4096) return json({ok:false,code:'unauthorized',error:'請透過 LINE 登入排班月曆'},401);
  let input;
  try {
    const body=await request.text();
    if(body.length>100000) return json({ok:false,error:'排班內容過大'},413);
    input=JSON.parse(body);
  } catch { return json({ok:false,error:'Invalid JSON'},400); }
  const actions=['bootstrap','managerSession','saveDay','clearDay','addStaff','deleteStaff','addShift','updateShift','deleteShift'];
  if(!input || !actions.includes(input.action) || !input.payload || typeof input.payload!=='object' || Array.isArray(input.payload)) return json({ok:false,error:'不支援的操作'},400);
  try {
    const response=await fetch(env.GAS_WEB_APP_URL,{
      method:'POST',headers:{'content-type':'application/json; charset=UTF-8'},redirect:'follow',
      body:JSON.stringify({kind:'calendar_api',action:input.action,payload:input.payload,lineAccessToken:tokenMatch[1],secret:env.GAS_SHARED_SECRET})
    });
    if(!response.ok) throw new Error('GAS upstream failed');
    const result=await response.json();
    if(!result || typeof result.ok!=='boolean') throw new Error('Invalid upstream response');
    return json(result,result.code==='unauthorized'?401:200);
  } catch {
    return json({ok:false,error:'GAS 回應失敗，請重新整理確認排班結果後再操作'},502);
  }
}

export async function verifyLineSignature(rawBody, signature, channelSecret) {
  if (!rawBody || !signature || !channelSecret) return false;
  try {
    const expected = Uint8Array.from(atob(signature), char => char.charCodeAt(0));
    if (expected.length !== 32) return false;
    const encoder = new TextEncoder();
    const key = await crypto.subtle.importKey(
      'raw', encoder.encode(channelSecret), {name: 'HMAC', hash: 'SHA-256'}, false, ['verify']
    );
    return await crypto.subtle.verify('HMAC', key, expected, encoder.encode(rawBody));
  } catch {
    return false;
  }
}

export function parsePostbackData(data) {
  if (typeof data !== 'string' || data.length > 300) return null;
  const qs = new URLSearchParams(data);
  const id = qs.get('id');
  const version = Number(qs.get('v'));
  const choice = qs.get('choice');
  if (qs.get('action') !== 'shift_reply' || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id || '')) return null;
  if (!Number.isSafeInteger(version) || version < 1 || !['agree', 'disagree'].includes(choice)) return null;
  return {id, version, choice};
}

async function handleEvent(event, env) {
  if (!event || event.source?.type !== 'user' || !/^U[0-9a-f]{32}$/i.test(event.source.userId || '')) return;
  const lineUid = event.source.userId;
  const eventId = event.webhookEventId || '';
  if (event.type === 'message' && event.message?.type === 'text') {
    const text = String(event.message.text || '').trim();
    if (text === '我的UID') {
      await replyLine(event.replyToken, '您的 LINE UID：\n' + lineUid + '\n\n若要接收臨時排班，請私下提供給團長。請勿公開張貼。', env);
      return;
    }
    if (!text.startsWith('綁定')) return;
    const result = await sendToGas({kind: 'bind_or_message', lineUid, text, eventId}, env);
    await replyLineSafe(event.replyToken, result.message, env);
    return;
  }
  if (event.type === 'postback') {
    const parsed = parsePostbackData(event.postback?.data);
    if (!parsed) return;
    const result = await sendToGas({kind: 'shift_reply', lineUid, eventId, ...parsed}, env);
    await replyLineSafe(event.replyToken, result.message, env);
  }
}

async function sendToGas(payload, env) {
  const response = await fetch(env.GAS_WEB_APP_URL, {
    method: 'POST',
    headers: {'content-type': 'application/json; charset=UTF-8'},
    body: JSON.stringify({...payload, secret: env.GAS_SHARED_SECRET}),
    redirect: 'follow'
  });
  if (!response.ok) throw new Error('GAS HTTP ' + response.status);
  let result;
  try { result = await response.json(); }
  catch { throw new Error('GAS response is not JSON'); }
  if (!result || result.ok !== true) throw new Error('GAS rejected the request: ' + String(result?.code || 'unknown'));
  return result;
}

async function replyLineSafe(replyToken, text, env) {
  if (!replyToken || !text) return;
  try { await replyLine(replyToken, text, env); }
  catch (error) {
    // GAS 可能已成功寫入；LINE 回覆訊息失敗不應反覆重送已完成事件。
    console.error('LINE confirmation reply failed:', String(error.message));
  }
}

async function replyLine(replyToken, text, env) {
  if (!replyToken) return;
  const response = await fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'POST',
    headers: {
      'content-type': 'application/json; charset=UTF-8',
      authorization: 'Bearer ' + env.LINE_CHANNEL_ACCESS_TOKEN
    },
    body: JSON.stringify({replyToken, messages: [{type: 'text', text: String(text).slice(0, 1800)}]})
  });
  if (!response.ok) throw new Error('LINE reply HTTP ' + response.status);
}

