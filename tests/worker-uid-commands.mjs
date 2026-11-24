import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {createHmac} from 'node:crypto';

const source=await readFile(new URL('../cloudflare/workers.js',import.meta.url),'utf8');
const worker=(await import('data:text/javascript;base64,'+Buffer.from(source).toString('base64'))).default;
const uid='U'+'a'.repeat(32);
const env={
  GAS_WEB_APP_URL:'https://example.invalid/exec',
  GAS_SHARED_SECRET:'test-secret',
  LINE_CHANNEL_SECRET:'test-secret-line',
  LINE_CHANNEL_ACCESS_TOKEN:'test-access-token',
  MASALU_CHANNEL:'backup'
};

function signedRequest(command){
  const body=JSON.stringify({events:[{
    type:'message',source:{type:'user',userId:uid},replyToken:'test-reply',
    webhookEventId:'uid-command-1',message:{type:'text',text:command}
  }]});
  return new Request('https://worker.test/webhook',{method:'POST',body,
    headers:{'content-type':'application/json',
      'x-line-signature':createHmac('sha256',env.LINE_CHANNEL_SECRET).update(body).digest('base64')}});
}

test('both LINE OAs recognize my UID case and spacing variants without GAS calls',async()=>{
  for(const command of ['我的UID','我的uid','我的 UID','我的 uid',' 我的 uid ','我的　ＵＩＤ']){
    const oldFetch=globalThis.fetch,seen=[];
    globalThis.fetch=async(url,init)=>{
      seen.push({url:String(url),body:JSON.parse(init.body)});
      return new Response('{}',{status:200});
    };
    try{
      const reply=await worker.fetch(signedRequest(command),env);
      assert.equal(reply.status,200,command);
      assert.equal(seen.length,1,command);
      assert.equal(seen[0].url,'https://api.line.me/v2/bot/message/reply');
      assert.ok(seen[0].body.messages[0].text.includes(uid),command);
      assert.equal(seen[0].body.replyToken,'test-reply');
    }finally{globalThis.fetch=oldFetch;}
  }
});
test('other commands still route to GAS rather than revealing UID',async()=>{
  const oldFetch=globalThis.fetch,seen=[];
  globalThis.fetch=async(url,init)=>{
    seen.push({url:String(url),body:JSON.parse(init.body)});
    return new Response(JSON.stringify(String(url)===env.GAS_WEB_APP_URL
      ? {ok:true,message:'其他訊息'} : {}),{status:200});
  };
  try{
    const r=await worker.fetch(signedRequest('個人班表'),env);
    assert.equal(r.status,200);
    assert.equal(seen[0].url,env.GAS_WEB_APP_URL);
    assert.equal(seen[0].body.kind,'personal_calendar_menu');
    assert.equal(seen[0].body.channel,'backup');
    assert.equal(seen[1].url,'https://api.line.me/v2/bot/message/reply');
  }finally{globalThis.fetch=oldFetch;}
});
