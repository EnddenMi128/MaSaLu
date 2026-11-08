import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';

const code = await readFile(new URL('../GAS/程式碼.js', import.meta.url), 'utf8');
const calendarCode = await readFile(new URL('../GAS/CalendarApi.js', import.meta.url), 'utf8');
const uuid = '11111111-1111-1111-1111-111111111111';
const uid = 'U'+'a'.repeat(32);
const managerUid = 'U'+'b'.repeat(32);

function harness() {
  const row = [false,uuid,'A04','2026-11-05','','','玉門','待回覆','','已發送','',1,'','','','',''];
  const logs=[], notices=[];
  const sheet = {
    getRange(r,c,n,m) {
      assert.equal(r,2);
      return {
        getValues:()=>[row.slice(c-1,c-1+(m||1))],
        setValues(v){ for(let i=0;i<v[0].length;i++)row[c-1+i]=v[0][i]; },
        setValue(v){row[c-1]=v;},
        clearContent(){row[c-1]='';}
      };
    }
  };
  const ctx = {
    console,
    Date,
    SHIFT_APP:null,
    LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    SpreadsheetApp:{flush(){}},
    PropertiesService:{getScriptProperties:()=>({getProperty:()=>''})},
    Utilities:{},
    UrlFetchApp:{fetch(){throw Error('no outbound line calls expected')}}
  };
  vm.createContext(ctx);
  vm.runInContext(code,ctx);
  vm.runInContext(calendarCode,ctx);
  ctx.shiftSheet_=()=>sheet;
  ctx.findShiftRow_=()=>2;
  ctx.employeeIndex_=()=>({A04:{name:'員工甲',uid},A01:{name:'團長',uid:managerUid}});
  ctx.masaluDate_=x=>String(x);
  ctx.safeLog_=(...items)=>logs.push(items);
  ctx.masaluNotifyManagers_=x=>{notices.push(x);return {sent:1,failed:0};};
  return {ctx,row,logs,notices};
}

test('agree writes exact response, notifies managers once, repeated clicks are idempotent',()=>{
  const h=harness();
  const input={lineUid:uid,id:uuid,version:1,choice:'agree',eventId:'evt'};
  const first=h.ctx.handleShiftReply_(input);
  assert.equal(first.code,'confirmed');
  assert.equal(h.row[7],'同意');
  assert.equal(h.notices.length,1);
  assert.equal(h.notices[0].status,'同意');
  const second=h.ctx.handleShiftReply_(input);
  assert.equal(second.code,'already');
  assert.equal(h.notices.length,1);
});

test('disagree persists original row with different status, only owner may respond',()=>{
  const h=harness();
  const wrong=h.ctx.handleShiftReply_({lineUid:managerUid,id:uuid,version:1,choice:'disagree'});
  assert.equal(wrong.code,'wrong_user');
  assert.equal(h.row[7],'待回覆');
  const result=h.ctx.handleShiftReply_({lineUid:uid,id:uuid,version:1,choice:'disagree'});
  assert.equal(result.code,'confirmed');
  assert.equal(h.row[7],'不同意');
  assert.equal(h.notices.length,1);
  assert.equal(h.notices[0].status,'不同意');
});

test('a reply never gets rolled back if manager push fails',()=>{
  const h=harness();
  h.ctx.masaluNotifyManagers_=()=>{throw Error('LINE push failed');};
  const result=h.ctx.handleShiftReply_({lineUid:uid,id:uuid,version:1,choice:'agree'});
  assert.equal(result.code,'confirmed');
  assert.equal(result.managerNotice.failed,1);
  assert.equal(h.row[7],'同意');
});

test('manager skip preserves declined record and rejects unbound manager',()=>{
  const h=harness();
  h.row[7]='不同意';
  const action={lineUid:managerUid,id:uuid,version:1,decision:'skip'};
  const bad=h.ctx.handleManagerDecision_({...action,lineUid:uid});
  assert.equal(bad.code,'unauthorized');
  const good=h.ctx.handleManagerDecision_(action);
  assert.equal(good.code,'skipped');
  assert.equal(h.row[7],'不同意');
  assert.match(h.row[16],/管理已略過/);
  assert.equal(h.ctx.handleManagerDecision_(action).code,'already');
});

test('declined assignments disappear from the visible calendar but remain in source data',()=>{
  const h=harness();
  h.ctx.masaluRevision_=()=>'revision';
  h.ctx.masaluDate_=x=>String(x);
  const data={
    employees:{rows:[['A04','員工甲','','','']]},
    shifts:{rows:[['ym','玉門',false]]},
    display:{rows:[]},
    schedule:{rows:[h.row]}
  };
  assert.equal(h.ctx.masaluBootstrap_(data).schedules.length,1);
  h.row[7]='不同意';
  assert.equal(h.ctx.masaluBootstrap_(data).schedules.length,0);
  assert.equal(data.schedule.rows[0][7],'不同意');
});

function editHarness(){
  const h=harness(), old=[false,uuid,'A04','2026-11-05','','','玉門','同意','','已發送','',1,'','','','',''];
  const records=[old];
  const edits=[];
  const sheet={
    appendRow(row){records.push(row);},
    getLastRow(){return records.length+1;},
    getRange(row,col,num=1,width=1){
      return {
        getValues(){return records.slice(row-2,row-2+num).map(r=>r.slice(col-1,col-1+width));},
        clearContent(){records[row-2].fill('');edits.push(row);},
      };
    }
  };
  const data={
    employees:{rows:[['A04','員工甲','','',''],['A05','員工乙','','','']]},
    shifts:{rows:[['ym','玉門',false]]},display:{rows:[]},
    schedule:{rows:records,sheet}
  };
  h.ctx.Utilities.getUuid=()=>uuid;
  h.ctx.masaluRevision_=()=>'rev';
  h.ctx.shiftSheet_=()=>sheet;
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'mock-token'})};
  return {h,data,records,edits};
}

test('clear day removes already agreed rows from active schedule, preserving staff roster',()=>{
  const {h,data,records,edits}=editHarness();
  const result=h.ctx.masaluSaveDay_(data,{date:'2026-11-05',assignments:{}},true);
  assert.equal(result.clear,true);
  assert.equal(result.removed.length,1);
  assert.equal(records[0][1],'');
  assert.deepEqual(edits,[2]);
  assert.equal(data.employees.rows.length,2);
});

test('saving unchanged pending roster does not generate a new notification',()=>{
  const {h,data,records,edits}=editHarness();
  records[0][7]='待回覆';
  const result=h.ctx.masaluSaveDay_(data,{date:'2026-11-05',assignments:{ym:['A04']}},false);
  assert.equal(result.addedIds.length,0);
  assert.equal(result.removed.length,0);
  assert.equal(result.retained,1);
  assert.equal(edits.length,0);
});

test('removing just one staff member reports a targeted cancellation',()=>{
  const {h,data,records}=editHarness();
  records.push([false,'22222222-2222-2222-2222-222222222222','A05','2026-11-05','','','玉門','待回覆','','已發送','',1,'','','','','']);
  const result=h.ctx.masaluSaveDay_(data,{date:'2026-11-05',assignments:{ym:['A04']}},false);
  assert.equal(result.retained,1);
  assert.equal(result.removed.length,1);
  assert.equal(result.removed[0].staffId,'A05');
});

test('clear notification broadcasts to bound staff; non-clear only notifies removed user',()=>{
  const {h}=editHarness(), deliveries=[];
  const uid1='U'+'1'.repeat(32),uid2='U'+'2'.repeat(32);
  h.ctx.employeeIndex_=()=>({A01:{name:'團長',uid:uid1},A04:{name:'員工甲',uid:uid2},A05:{name:'未綁定',uid:''}});
  h.ctx.UrlFetchApp={fetch(_url,options){deliveries.push(JSON.parse(options.payload));return {getResponseCode:()=>200};}};
  const delta={date:'2026-11-05',clear:true,addedIds:[],removed:[{staffId:'A04',shift:'玉門'}]};
  const all=h.ctx.masaluNotifyCancellations_(delta);
  assert.equal(all.sent,2);
  assert.equal(all.skipped,1);
  assert.equal(deliveries.length,2);
  deliveries.length=0;
  const targeted=h.ctx.masaluNotifyCancellations_({...delta,clear:false});
  assert.equal(targeted.sent,1);
  assert.equal(deliveries[0].to,uid2);
});

function replacementHarness(count=3) {
  const h=harness(); h.row[7]='不同意';
  const records=[h.row];
  const roster={A04:{name:'無法出席',uid},A01:{name:'團長',uid:managerUid}};
  for(let i=0;i<count;i++) roster['E'+i]={name:'人員'+i,uid:'U'+'c'.repeat(32)};
  roster.unbound={name:'未綁定',uid:''};
  const sheet={
    getLastRow:()=>records.length+1,
    appendRow:r=>records.push(r),
    getRange(r,c,n=1,m=1){return {
      getValues:()=>records.slice(r-2,r-2+n).map(row=>row.slice(c-1,c-1+m)),
      setValue:v=>{records[r-2][c-1]=v;}
    };}
  };
  h.ctx.shiftSheet_=()=>sheet;
  h.ctx.employeeIndex_=()=>roster;
  h.ctx.findShiftRow_=(_s,id)=>{const i=records.findIndex(r=>r[1]===id);return i<0?0:i+2;};
  h.ctx.Utilities.getUuid=()=> '22222222-2222-2222-2222-222222222222';
  const sends=[];
  h.ctx.sendShiftRow_=(_row,_force,shiftId)=>{
    sends.push(shiftId);
    records.find(r=>r[1]===shiftId)[9]='已發送';
    return {code:'sent'};
  };
  return {...h,records,roster,sends,input:{lineUid:managerUid,id:uuid,version:1,decision:'replace'}};
}

test('replacement picker lists same-day unscheduled bound people across all shifts',()=>{
  const h=replacementHarness();
  h.records.push([false,'other','E0','2026-11-05','','','田園','同意']);
  h.records.push([false,'next-day','E1','2026-11-06','','','田園','同意']);
  const result=h.ctx.handleManagerDecision_(h.input);
  assert.match(result.message,/^請選擇替補對象/);
  const labels=result.quickReply.items.map(i=>i.action.label);
  assert.deepEqual([...labels],['團長','人員1','人員2','取消']);
  assert.equal(h.records.length,3);
  assert.equal(h.sends.length,0);
});

test('replacement picker paginates within LINE limits and handles empty/cancel',()=>{
  const h=replacementHarness(25);
  const first=h.ctx.handleManagerDecision_(h.input);
  assert.ok(first.quickReply.items.length<=13);
  assert.ok(first.quickReply.items.some(i=>i.action.label==='下一頁'));
  const next=h.ctx.handleManagerDecision_({...h.input,page:1});
  assert.equal(next.quickReply.items.length,13);
  assert.equal(h.ctx.handleManagerDecision_({...h.input,decision:'cancel'}).code,'cancelled');
  h.ctx.employeeIndex_=()=>({A04:h.roster.A04,A01:h.roster.A01});
  h.records.push([false,'other','A01','2026-11-05','','','田園','待回覆']);
  assert.equal(h.ctx.handleManagerDecision_(h.input).code,'no_candidates');
});

test('selection rechecks availability, saves original shift, sends once and rejects duplicate managers',()=>{
  const h=replacementHarness();
  h.records.push([false,'other','E0','2026-11-05','','','田園','待回覆']);
  const stale=h.ctx.handleManagerDecision_({...h.input,decision:'select',staffId:'E0'});
  assert.match(stale.message,/該人員目前無法替補。/);
  assert.ok(stale.quickReply.items.every(i=>i.action.label!=='人員0'));
  assert.equal(h.sends.length,0);
  const result=h.ctx.handleManagerDecision_({...h.input,decision:'select',staffId:'E1'});
  assert.equal(result.code,'replaced');
  assert.equal(h.records[2][2],'E1');
  assert.equal(h.records[2][6],'玉門');
  assert.equal(h.records[0][7],'不同意');
  assert.match(h.records[0][16],/管理已略過/);
  assert.equal(h.sends.length,1);
  assert.equal(h.ctx.handleManagerDecision_({...h.input,decision:'select',staffId:'E2'}).code,'already');
  assert.equal(h.sends.length,1);
  assert.equal(h.records.length,3);
});

test('failed replacement delivery can retry the same person without creating another assignment',()=>{
  const h=replacementHarness();
  h.ctx.sendShiftRow_=()=>({code:'failed'});
  const input={...h.input,decision:'select',staffId:'E0'};
  const failure=h.ctx.handleManagerDecision_(input);
  assert.equal(failure.notification,'failed');
  assert.match(failure.quickReply.items[0].action.data,/decision=select&staff=E0$/);
  assert.equal(h.ctx.handleManagerDecision_(input).notification,'failed');
  assert.equal(h.records.length,2);
  assert.equal(h.ctx.handleManagerDecision_({...input,lineUid:uid}).code,'unauthorized');
});

test('retry reuses the replacement if saving the source link was interrupted',()=>{
  const h=replacementHarness();
  const target='22222222-2222-2222-2222-222222222222';
  h.records.push([false,target,'E0','2026-11-05','','','玉門','待回覆','','未發送','',0,'','','','','替補來源：'+uuid]);
  const result=h.ctx.handleManagerDecision_({...h.input,decision:'select',staffId:'E0'});
  assert.equal(result.notification,'sent');
  assert.equal(h.records.length,2);
  assert.deepEqual(h.sends,[target]);
  assert.ok(h.row[16].includes('[替補ID:'+target+']'));
});

test('unauthenticated GET cannot read the roster or calendar through legacy JSONP',()=>{
  const {ctx}=harness();
  ctx.ContentService={
    MimeType:{JSON:'json'},
    createTextOutput:text=>({setMimeType:type=>({text,type})})
  };
  ctx.masaluBootstrap_=()=>{throw Error('public GET must never load scheduling data');};
  const response=ctx.doGet({parameter:{action:'bootstrap',callback:'loadCalendar'}});
  assert.equal(response.type,'json');
  assert.equal(JSON.parse(response.text).code,'unauthorized');
  assert.equal(JSON.parse(ctx.doGet({parameter:{}}).text).ok,true);
});
