import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';

const code=await readFile(new URL('../GAS/程式碼.js',import.meta.url),'utf8');
const uid='U'+'a'.repeat(32);
const other='U'+'b'.repeat(32);
const quota=JSON.stringify({message:'You have reached your monthly limit.'});
const http=(status,body='')=>({getResponseCode:()=>status,getContentText:()=>body,getAllHeaders:()=>({})});

function prepare(settings, primary=http(429,quota), secondary=http(200)) {
  const calls=[],log=[];
  const ctx={
    console:{log:(...args)=>log.push(args.join(' ')),error:(...args)=>log.push(args.join(' '))},
    PropertiesService:{getScriptProperties:()=>({getProperty:key=>settings[key]||''})},
    UrlFetchApp:{fetch(_url,options){calls.push(options);return options.headers.Authorization==='Bearer primary'?primary:secondary;}}
  };
  vm.createContext(ctx);vm.runInContext(code,ctx);
  ctx.masaluLineRetryKey_=identity=>'key:'+identity;
  return {ctx,calls,log};
}
const settings={
  LINE_CHANNEL_ACCESS_TOKEN:'primary',
  MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN:'backup',
  MASALU_BACKUP_FRIEND_UIDS:uid,
  MASALU_BACKUP_OA_ENABLED:'true'
};
const message=[{type:'text',text:'測試'}];

test('fallback is disabled until secondary webhook is explicitly enabled',()=>{
  const h=prepare({...settings,MASALU_BACKUP_OA_ENABLED:''});
  const res=h.ctx.masaluPushWithQuotaFallback_(uid,message);
  assert.equal(res.channel,'primary');
  assert.deepEqual(h.calls.map(x=>x.headers.Authorization),['Bearer primary']);
});
test('eligible backup friend switches on verified monthly limit only',()=>{
  const h=prepare(settings);
  const res=h.ctx.masaluPushWithQuotaFallback_(uid,message);
  assert.equal(res.channel,'backup');assert.equal(res.response.getResponseCode(),200);
  assert.deepEqual(h.calls.map(x=>x.headers.Authorization),['Bearer primary','Bearer backup']);
  assert.equal(h.calls[0].payload,h.calls[1].payload);
});
test('nonquota 429 and other primary statuses never switch channel',()=>{
  for(const primary of [http(429,'rate limit'),http(500,'temporary'),http(200,'ok'),http(403,'forbidden')]){
    const h=prepare(settings,primary);
    assert.equal(h.ctx.masaluPushWithQuotaFallback_(uid,message).channel,'primary');
    assert.equal(h.calls.length,1);
  }
});
test('missing backup token or missing exact UID will not switch',()=>{
  for(const conf of [
    {...settings,MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN:''},
    {...settings,MASALU_BACKUP_FRIEND_UIDS:other},
    {...settings,MASALU_BACKUP_FRIEND_UIDS:uid.slice(0,-1)},
    {...settings,MASALU_BACKUP_FRIEND_UIDS:''}
  ]){
    const h=prepare(conf);
    assert.equal(h.ctx.masaluPushWithQuotaFallback_(uid,message).channel,'primary');
    assert.equal(h.calls.length,1);
  }
});
test('backup failure is returned as a failure, never silently marked successful',()=>{
  const h=prepare(settings,http(429,quota),http(403,'not a friend'));
  const res=h.ctx.masaluPushWithQuotaFallback_(uid,message);
  assert.equal(res.channel,'backup');assert.equal(res.response.getResponseCode(),403);
});
test('retry identity is separated for primary and secondary and no tokens logged',()=>{
  const h=prepare(settings);
  h.ctx.masaluPushWithQuotaFallback_(uid,message,'shift:123');
  const keys=h.calls.map(x=>x.headers['X-Line-Retry-Key']);
  assert.deepEqual(keys,['key:primary:shift:123:'+uid,'key:backup:shift:123:'+uid]);
  assert.equal(h.log.join(' ').includes('primary'),false);
  assert.equal(h.log.join(' ').includes('backup'),false);
});
