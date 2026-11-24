import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';

const account=await readFile(new URL('../GAS/AccountLink.js',import.meta.url),'utf8');
const menu=await readFile(new URL('../GAS/CalendarRichMenu.js',import.meta.url),'utf8');
const uid='U'+'a'.repeat(32);
const other='U'+'b'.repeat(32);

function prepare({roster={A10:{uid,name:'A10'}},status=200,profileUid=uid,friends='',throwNetwork=false,channel='backup'}={}) {
  const state={LINE_CHANNEL_ACCESS_TOKEN:'primary-token',
    MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN:'backup-token',
    MASALU_BACKUP_OA_ENABLED:'true',MASALU_BACKUP_FRIEND_UIDS:friends};
  const props={getProperty:k=>state[k]||'',setProperty:(k,v)=>{state[k]=v},getProperties:()=>({...state})};
  const calls=[];
  const ctx={
    PropertiesService:{getScriptProperties:()=>props},
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    UrlFetchApp:{fetch(url,options){calls.push({url,options});if(throwNetwork)throw Error('network');return {
      getResponseCode:()=>status,
      getContentText:()=>JSON.stringify({userId:profileUid,displayName:'T'})};}},
    validLineUid_:u=>/^U[0-9a-f]{32}$/i.test(String(u||'')),
    employeeIndex_:()=>roster,
    console:{warn(){},log(){},error(){}}
  };
  vm.createContext(ctx);
  vm.runInContext(account,ctx,{filename:'AccountLink.js'});
  vm.runInContext(menu,ctx,{filename:'CalendarRichMenu.js'});
  let synced=false;
  ctx.masaluCalendarMenuSettings_=()=>({
    properties:props,assets:'mock',page:{start:'2026-11-22',end:'2026-12-05'},
    channels:[{name:channel,token:channel==='backup'?'backup-token':'primary-token',
      prefix:channel==='backup'?'MASALU_BACKUP_PERSONAL_MENU_':'MASALU_PERSONAL_MENU_',
      ...(channel==='backup'?{uids:new Set((state.MASALU_BACKUP_FRIEND_UIDS||'').split(/[\\s,]+/).filter(Boolean))}:{})}]
  });
  ctx.masaluEnsureDefaultCalendarMenu_=()=>{};
  ctx.masaluSyncPersonalMenus_=(_page,_assets,_token,_props,who,_prefix,allow)=>{synced=true;
    if(allow&&!allow.has(who))throw Error('not allowed');
    return {status:'empty',entries:[]};
  };
  return {ctx,state,calls,wasSynced:()=>synced};
}

test('existing bound backup friend self-recovers without repeat follow event',()=>{
  const h=prepare();const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'backup');
  assert.equal(r.code,'empty');assert.equal(h.state.MASALU_BACKUP_FRIEND_UIDS,uid);
  assert.equal(h.calls.length,1);
  assert.equal(h.calls[0].url,'https://api.line.me/v2/bot/profile/'+uid);
  assert.equal(h.calls[0].options.headers.Authorization,'Bearer backup-token');
  assert.equal(h.wasSynced(),true);
});
test('already verified friend is not re-queried',()=>{
  const h=prepare({friends:uid});const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'backup');
  assert.equal(r.code,'empty');assert.equal(h.calls.length,0);
});
test('unbound or duplicate identity cannot be allowed',()=>{
  for(const roster of [{},{A10:{uid},A11:{uid}}]){
    const h=prepare({roster});const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'backup');
    assert.equal(r.code,'menu_unbound');assert.equal(h.calls.length,0);
    assert.equal(h.state.MASALU_BACKUP_FRIEND_UIDS,'');assert.equal(h.wasSynced(),false);
  }
});
test('unverified, mismatched, or offline profile never registers a friend',()=>{
  for(const options of [{status:404},{profileUid:other},{throwNetwork:true}]){
    const h=prepare(options);const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'backup');
    assert.equal(r.code,'menu_unbound');assert.equal(h.state.MASALU_BACKUP_FRIEND_UIDS,'');
    assert.equal(h.wasSynced(),false);
  }
});
test('primary channel cannot modify secondary friend allowlist',()=>{
  const h=prepare({channel:'primary'});const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'primary');
  assert.equal(r.code,'empty');assert.equal(h.calls.length,0);
  assert.equal(h.state.MASALU_BACKUP_FRIEND_UIDS,'');
});
test('recovery preserves existing verified other users',()=>{
  const h=prepare({friends:other});const r=h.ctx.masaluRefreshOwnCalendarMenu_(uid,'backup');
  assert.equal(r.code,'empty');
  assert.deepEqual(new Set(h.state.MASALU_BACKUP_FRIEND_UIDS.split('\n')),new Set([other,uid]));
});
