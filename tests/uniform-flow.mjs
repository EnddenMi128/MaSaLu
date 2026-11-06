import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';

const sources=await Promise.all(['程式碼.js','CalendarApi.js','UniformNotifications.js']
  .map(name=>readFile(new URL('../GAS/'+name,import.meta.url),'utf8')));
const uid='U'+'2'.repeat(32);

function harness(){
  const rows=[],pushes=[],logs=[];
  const sheet={
    appendRow(row){rows.push([...row]);},
    getRange(r,c,_n=1,_width=1){
      return {
        setValues(values){values[0].forEach((value,i)=>rows[r-2][c-1+i]=value);},
        setValue(value){rows[r-2][c-1]=value;}
      };
    }
  };
  const ctx={console,Date,Set,Map,
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    PropertiesService:{getScriptProperties:()=>({getProperty:()=> 'mock-token'})},
    UrlFetchApp:{fetch(_url,options){pushes.push(JSON.parse(options.payload));return {getResponseCode:()=>200};}},
    Utilities:{formatDate(_date,_tz,format){return format==='yyyy-MM-dd'?'2026-11-19':'12:00';}}
  };
  vm.createContext(ctx);
  sources.forEach(source=>vm.runInContext(source,ctx));
  ctx.masaluDate_=value=>String(value);
  ctx.masaluUniformTable_=()=>({sheet,rows});
  ctx.masaluLoad_=()=>({uniforms:{sheet,rows}});
  ctx.employeeIndex_=()=>({A02:{name:'A02',uid}});
  ctx.safeLog_=(...args)=>logs.push(args);
  return {ctx,rows,sheet,pushes,logs};
}

test('blank uniform starts one five-minute wait and unchanged edits do not restart it',()=>{
  const {ctx,rows,sheet}=harness();
  const data={uniforms:{sheet,rows}};
  const first=ctx.masaluSetUniformForDate_(data,'2026-11-19','',false,true);
  assert.equal(first.deferNotifications,true);
  assert.equal(first.requestA02,true);
  assert.equal(rows[0][2],'pending');
  assert.ok(rows[0][3] instanceof Date);
  const deadline=rows[0][3].getTime();
  assert.ok(deadline>Date.now()+4*60*1000);
  const second=ctx.masaluSetUniformForDate_(data,'2026-11-19','',false,true);
  assert.equal(second.requestA02,false);
  assert.equal(rows[0][3].getTime(),deadline);
  assert.equal(rows[0][6],1);
});

test('only bound A02 may choose; choice preserves existing shift replies',()=>{
  const {ctx,rows,sheet}=harness();
  ctx.masaluSetUniformForDate_({uniforms:{sheet,rows}},'2026-11-19','',false,true);
  let uniformUpdates=0, released=0;
  ctx.masaluNotifyUniformChange_=()=>{uniformUpdates++;return {sent:0,failed:0};};
  ctx.masaluSendDeferredUniform_=()=>{released++;return {sent:2,failed:0};};
  const input={lineUid:uid,date:'2026-11-19',uniform:'black',version:1};
  assert.equal(ctx.handleUniformChoice_({...input,lineUid:'U'+'3'.repeat(32)}).code,'unauthorized');
  assert.equal(ctx.handleUniformChoice_(input).code,'confirmed');
  assert.equal(rows[0][1],'black');
  assert.equal(rows[0][2],'confirmed');
  assert.equal(uniformUpdates,1);
  assert.equal(released,1);
  assert.equal(ctx.handleUniformChoice_(input).code,'expired');
});

test('five-minute expiry sends 未擇衣; at 12:00 Taipei A02 is reminded only once',()=>{
  const {ctx,rows,sheet,pushes}=harness();
  ctx.masaluSetUniformForDate_({uniforms:{sheet,rows}},'2026-11-19','',false,true);
  rows[0][3]=new Date(Date.now()-1000);
  let releases=0;
  ctx.masaluSendDeferredUniform_=()=>{releases++;return {sent:2,failed:0};};
  const first=ctx.masaluUniformTick_();
  assert.equal(rows[0][2],'not_selected');
  assert.equal(first.expired,1);
  assert.equal(first.released,2);
  assert.equal(first.reminded,1);
  assert.equal(rows[0][5],'2026-11-19');
  assert.equal(pushes.length,1);
  assert.equal(pushes[0].to,uid);
  assert.match(pushes[0].messages[0].altText,/團服待確認/);
  const second=ctx.masaluUniformTick_();
  assert.equal(second.reminded,0);
  assert.equal(pushes.length,1);
  assert.equal(releases,2); // retries any undelivered rows; sent rows are skipped in production
});

test('clearing a day ends its pending uniform request',()=>{
  const {ctx,rows,sheet}=harness();
  const data={uniforms:{sheet,rows}};
  ctx.masaluSetUniformForDate_(data,'2026-11-19','',false,true);
  const result=ctx.masaluSetUniformForDate_(data,'2026-11-19','',true,false);
  assert.equal(result.deferNotifications,false);
  assert.equal(result.requestA02,false);
  assert.equal(rows[0][2],'cleared');
});
