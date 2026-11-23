import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
const base=new URL('../',import.meta.url);
function harness(){
 const rows=[['A01','一','','U'+'1'.repeat(32),''],['A02','二','','U'+'2'.repeat(32),''],['A03','三','','','']];
 const props=new Map(),sent=[];let valid=true,fail=false;
 const ctx={Date,console,LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},SpreadsheetApp:{flush(){}},PropertiesService:{getScriptProperties:()=>({getProperty:k=>props.get(k),setProperty:(k,v)=>props.set(k,v)})}};
 vm.createContext(ctx);
 vm.runInContext(readFileSync(new URL('GAS/CalendarApi.js',base),'utf8')+'\n'+readFileSync(new URL('GAS/AccountLink.js',base),'utf8'),ctx);
 ctx.SHIFT_APP={EMPLOYEES:'employees',EMP_HEADERS:[]};ctx.validLineUid_=v=>/^U[0-9a-f]{32}$/.test(v);
 ctx.masaluVerifiedLineProfile_=()=>valid?{userId:'U'+'3'.repeat(32)}:null;
 ctx.masaluTable_=()=>({rows,sheet:{getRange:(r)=>({setValues:([v])=>{rows[r-2][3]=v[0];rows[r-2][4]=v[1];}})}});
 ctx.employeeIndex_=()=>Object.fromEntries(rows.map(r=>[r[0],{uid:r[3],name:r[1]}]));ctx.safeLog_=()=>{};
 ctx.masaluPushWithQuotaFallback_=(uid,messages,key)=>{if(fail)throw Error('offline');sent.push({uid,messages,key});return {response:{getResponseCode:()=>200}};};
 return {ctx,rows,sent,invalid:()=>valid=false,offline:v=>fail=v,bind:id=>ctx.masaluNfcBind_({employeeId:id},'token')};
}
test('NFC A03 binds verified identity, sends success and all three manager notices once',()=>{
 const h=harness(),res=h.bind('A03');assert.equal(res.code,'bound');assert.equal(h.rows[2][3],'U'+'3'.repeat(32));assert.equal(h.sent.length,4);assert.match(h.sent[0].messages[0].text,/liff.line.me/);assert.equal(h.bind('A03').code,'already_bound');assert.equal(h.sent.length,4);
 assert.equal(h.bind('A01').employeeId,'A03');assert.equal(h.rows[0][3],'U'+'1'.repeat(32));
});
test('invalid identity, unknown and occupied employee IDs never write or notify',()=>{
 for(const id of ['A01','A99','../A03']){const h=harness();assert.equal(h.bind(id).ok,false);assert.equal(h.rows[2][3],'');assert.equal(h.sent.length,0);}
 const h=harness();h.invalid();assert.equal(h.bind('A03').code,'unauthorized');assert.equal(h.rows[2][3],'');
});
test('failed notifications preserve binding and retry only pending notices on next scan',()=>{
 const h=harness();h.offline(true);assert.equal(h.bind('A03').notifications.failed,4);assert.ok(h.rows[2][3]);h.offline(false);assert.equal(h.bind('A03').notifications.sent,4);h.bind('A03');assert.equal(h.sent.length,4);
});
test('LIFF direct and OAuth nested emp parameters are handled',()=>{
 for(const search of ['?emp=A03','?liff.state=%3Femp%3DA03','?liff.state=%2F%3Femp%3DA03']){const p=new URLSearchParams(search),s=new URLSearchParams((p.get('liff.state')||'').replace(/^\/?\?/,''));assert.equal(p.get('emp')||s.get('emp'),'A03');}
});
