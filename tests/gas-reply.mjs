import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
import {createHash,createHmac} from 'node:crypto';

const code = await readFile(new URL('../GAS/程式碼.js', import.meta.url), 'utf8');
const calendarCode = await readFile(new URL('../GAS/CalendarApi.js', import.meta.url), 'utf8');
const swapCode = await readFile(new URL('../GAS/ShiftSwap.js', import.meta.url), 'utf8');
const menuActionsCode = await readFile(new URL('../GAS/CalendarMenuActions.js', import.meta.url), 'utf8');
const uuid = '11111111-1111-1111-1111-111111111111';
const uid = 'U'+'a'.repeat(32);
const managerUid = 'U'+'b'.repeat(32);

function harness({mockManagerNotice=true}={}) {
  const row = [false,uuid,'A04','2026-11-05','','','玉門','待回覆','','已發送','',1,'','','','',''];
  const logs=[], notices=[], serialRows=[];
  const sheet = {
    getLastRow:()=>2,
    getRange(r,c,n,m) {
      assert.equal(r,2);
      return {
        getValues:()=>[row.slice(c-1,c-1+(m||1))],
        getDisplayValues:()=>[row.slice(c-1,c-1+(m||1)).map(String)],
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
    LockService:{getScriptLock:()=>({waitLock(){},tryLock(){return true;},releaseLock(){}})},
    SpreadsheetApp:{flush(){}},
    PropertiesService:{getScriptProperties:()=>({getProperty:()=>''})},
    Utilities:{},
    UrlFetchApp:{fetch(){throw Error('no outbound line calls expected')}}
  };
  vm.createContext(ctx);
  vm.runInContext(code,ctx);
  vm.runInContext(calendarCode,ctx);
  vm.runInContext(swapCode,ctx);
  const ledger={getDataRange:()=>({getValues:()=>[vm.runInContext('MASALU_SERIAL_HEADERS',ctx),...serialRows]}),
    appendRow:values=>serialRows.push(Array.from(values))};
  ctx.openBook_=()=>({getSheetByName:name=>name==='MaSaLu流水號'?ledger:null});
  ctx.shiftSheet_=()=>sheet;
  ctx.findShiftRow_=()=>2;
  ctx.employeeIndex_=()=>({A04:{name:'員工甲',uid},A01:{name:'團長',uid:managerUid}});
  const parseCalendarDate=ctx.masaluDate_;
  ctx.masaluDate_=x=>String(x);
  ctx.safeLog_=(...items)=>logs.push(items);
  if (mockManagerNotice) ctx.masaluNotifyManagers_=x=>{notices.push(x);return {sent:1,failed:0};};
  return {ctx,row,logs,notices,serialRows,ledger,parseCalendarDate};
}

test('LINE binding requires the name and unused code from the same employee row', () => {
  const {ctx} = harness();
  const rows = [
    ['A04', '測試員工', 'ABCDEF123456', '', ''],
    ['A05', '測試人員', 'FEDCBA654321', '', '']
  ];
  const sheet = {
    getLastRow: () => 1 + rows.length,
    getRange(row, column, count, width) {
      return {
        getDisplayValues: () => rows.slice(row - 2, row - 2 + count)
          .map(values => values.slice(column - 1, column - 1 + width)),
        setValues: ([values]) => rows[row - 2].splice(column - 1, width, ...values)
      };
    }
  };
  ctx.openBook_ = () => ({getSheetByName: () => sheet});
  const bind = (text, lineUid = uid) => ctx.handleLineText_({text, lineUid});

  assert.equal(bind('綁定 ABCDEF123456').code, 'invalid');
  assert.equal(bind('ABCDEF123456').code, 'invalid');
  assert.equal(bind('測試員工 FEDCBA654321').code, 'invalid_code');
  assert.equal(rows[0][2], 'ABCDEF123456');
  assert.equal(rows[1][2], 'FEDCBA654321');
  assert.equal(bind('測試員工\tabcdef123456').code, 'bound');
  assert.deepEqual(rows[0].slice(2, 4), ['', uid]);
  assert.equal(bind('測試員工 ABCDEF123456').code, 'already_bound');
  assert.equal(bind('測試員工 ABCDEF123456', managerUid).code, 'invalid_code');
  assert.equal(bind('測試人員FEDCBA654321', managerUid).code, 'bound');
  assert.deepEqual(rows[1].slice(2, 4), ['', managerUid]);
});

test('agree writes exact response, notifies managers once, repeated clicks are idempotent',()=>{
  const h=harness();
  const input={lineUid:uid,id:uuid,version:1,choice:'agree',eventId:'evt'};
  const first=h.ctx.handleShiftReply_(input);
  assert.equal(first.code,'confirmed');
  assert.equal(first.serial,'A-000001');
  assert.equal(first.message,'✅ 員工甲 已同意排班\n2026-11-05 ・ 玉門\nA-000001');
  assert.equal(first.replyMessage.altText,'✅ 員工甲 已同意排班｜2026-11-05 ・ 玉門');
  const [title,serial]=first.replyMessage.contents.body.contents[0].contents;
  assert.equal(title.text,'✅ 員工甲 已同意排班');
  assert.equal(serial.text,'A-000001');
  assert.equal(serial.align,'end');
  assert.equal(serial.size,'xs');
  assert.deepEqual(Array.from(first.replyMessage.contents.footer.contents,b=>b.action.label),['聯絡團長','調班請求','我要請假']);
  assert.equal(h.row[7],'同意');
  assert.equal(h.notices.length,1);
  assert.equal(h.notices[0].status,'同意');
  const second=h.ctx.handleShiftReply_(input);
  assert.equal(second.code,'already');
  assert.equal(second.serial,first.serial);
  assert.equal(h.serialRows.length,1);
  assert.equal(second.replyMessage.contents.footer.contents.length,3);
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
  assert.equal(result.serial,'D-000001');
  assert.match(result.message,/流水號：D-000001/);
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

test('calendar dates use the loaded spreadsheet time zone and batched LINE checks still enforce the channel',()=>{
  const {ctx,row,parseCalendarDate}=harness();
  row[3]=new Date('2026-11-04T16:30:00Z');
  row[7]='同意';
  ctx.openBook_=()=>{throw Error('calendar rendering should reuse its loaded time zone');};
  ctx.Utilities.formatDate=(date,timeZone,format)=>{
    assert.equal(timeZone,'Asia/Taipei');
    assert.equal(format,'yyyy-MM-dd');
    return new Intl.DateTimeFormat('en-CA',{timeZone,year:'numeric',month:'2-digit',day:'2-digit'}).format(date);
  };
  ctx.masaluDate_=parseCalendarDate;
  const employees={rows:[['A04','員工甲','',uid]]};
  ctx.masaluTable_=()=>employees;
  ctx.masaluLoad_=()=>({employees,shifts:{rows:[['ym','玉門',false]]},display:{rows:[]},
    schedule:{rows:[row]},timeZone:'Asia/Taipei'});
  ctx.masaluRevision_=()=> 'revision';
  const response=(data,status=200)=>({getResponseCode:()=>status,getContentText:()=>JSON.stringify(data)});
  let channel='2010020470', batches=0;
  ctx.UrlFetchApp.fetchAll=requests=>{
    batches++;
    assert.equal(requests.length,2);
    assert.match(requests[0].url,/\/verify\?access_token=test-token$/);
    assert.equal(requests[1].headers.Authorization,'Bearer test-token');
    return [response({client_id:channel,expires_in:60}),response({userId:uid})];
  };
  assert.equal(ctx.masaluApi('bootstrap',{},'test-token').schedules[0].date,'2026-11-05');
  assert.equal(batches,1);
  channel='different-channel';
  assert.equal(ctx.masaluApi('bootstrap',{},'test-token').code,'unauthorized');
});

test('an employee sees the confirmed crew only on their own work dates',()=>{
  const {ctx,row}=harness();
  const date='2026-11-05', otherDate='2026-11-06';
  const coworker=[...row];
  coworker[2]='A05';
  coworker[6]='田園';
  const declined=[...row];
  declined[2]='A06';
  declined[7]='不同意';
  const elsewhere=[...coworker];
  elsewhere[3]=otherDate;
  elsewhere[7]='同意';
  const dayOff=[...declined];
  dayOff[6]='休假';
  dayOff[7]='待回覆';
  row[7]='同意';
  const data={
    employees:{rows:[['A04','員工甲'],['A05','員工乙'],['A06','員工丙']]},
    shifts:{rows:[['ym','玉門',false],['ty','田園',false],['off','休假',true]]},
    display:{rows:[]},
    schedule:{rows:[row,coworker,declined,elsewhere,dayOff]}
  };
  ctx.masaluLoad_=()=>data;
  ctx.masaluRevision_=()=>'revision';
  ctx.masaluViewer_=()=>({id:'A04',manager:false});
  const read=()=>ctx.masaluApi('bootstrap',{},'test-token');
  const roster=result=>Array.from(result.staff,person=>person.id);
  assert.deepEqual(roster(read()),['A04']);
  coworker[7]='同意';
  assert.deepEqual(roster(read()),['A04']);
  declined[16]='[管理已略過]';
  const published=read();
  assert.deepEqual(roster(published),['A04','A05']);
  assert.deepEqual(Array.from(published.schedules,r=>[r.date,r.shiftId]),[[date,'ym'],[date,'ty']]);
  coworker[7]='待回覆';
  assert.deepEqual(roster(read()),['A04']);
  ctx.masaluViewer_=()=>({id:'A03',manager:true});
  assert.equal(read().schedules.length,4);
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

test('one person cannot be at both main locations but can also have a morning shift elsewhere',()=>{
  const {h,data,records,edits}=editHarness();
  data.shifts.rows.push(['ty','田園',false],['am','早班・其他地點',false]);
  assert.throws(()=>h.ctx.masaluSaveDay_(data,{
    date:'2026-11-05',assignments:{ym:['A04'],ty:['A04']}
  },false),/只能安排玉門或田園其中一邊/);
  assert.equal(records.length,1);
  assert.equal(edits.length,0);
  const result=h.ctx.masaluSaveDay_(data,{
    date:'2026-11-05',assignments:{ym:['A04'],am:['A04']}
  },false);
  assert.equal(result.retained,1);
  assert.equal(result.addedIds.length,1);
  assert.equal(records[1][6],'早班・其他地點');
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


function followupHarness() {
  const h=harness(), requests=[], cache=new Map(), pushes=[];
  h.row[7]='同意';
  let reviewHeaders=['',''];
  const sheet={
    getLastRow:()=>requests.length+1,
    appendRow:row=>requests.push(Array.from(row)),
    getRange:(row,col)=>row===1 && col===14
      ? {getValues:()=>[reviewHeaders],setValues:([headers])=>{reviewHeaders=headers;}}
      : {setValue(value){requests[row-2][col-1]=value;}}
  };
  const readTable=h.ctx.masaluTable_;
  h.ctx.openBook_=()=>({getSheetByName:name=>name==='MaSaLu流水號'?h.ledger:sheet});
  h.ctx.masaluTable_=(name,headers,book)=>name==='MaSaLu流水號'?readTable(name,headers,book):{sheet,rows:requests};
  h.ctx.CacheService={getScriptCache:()=>({get:key=>cache.get(key),put:(key,value)=>cache.set(key,value)})};
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-secret'})};
  h.ctx.masaluLineRetryKey_=value=>value;
  h.ctx.Utilities={
    getUuid:()=> 'ABCDEF01-1111-1111-1111-111111111111',
    computeHmacSha256Signature:(value,key)=>createHmac('sha256',key).update(value).digest(),
    base64EncodeWebSafe:value=>value.toString('base64url')
  };
  const fail=new Set();
  h.ctx.UrlFetchApp={fetch(_url,options){
    const body=JSON.parse(options.payload);
    pushes.push(body);
    return {getResponseCode:()=>fail.has(body.to)?500:200};
  }};
  return {...h,requests,pushes,fail,input:{lineUid:uid,id:uuid,version:1}};
}

test('the three approved-card actions deliver their prompts to the employee through GAS',()=>{
  const h=followupHarness();
  h.ctx.masaluDate_=value=>value instanceof Date?'2026-11-05':String(value);
  h.ctx.ContentService={MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>JSON.parse(text)})};
  const run=(choice,eventId)=>h.ctx.doPost({postData:{contents:JSON.stringify({
    ...h.input,secret:'test-secret',kind:'shift_followup',choice,eventId,timestamp:123
  })}});
  assert.equal(run('contact','contact-event').delivered,true);
  assert.equal(run('swap','swap-event').delivered,true);
  assert.equal(run('leave','leave-event').delivered,true);
  assert.deepEqual(h.pushes.map(push=>push.to),[uid,uid,uid]);
  assert.deepEqual(h.pushes.map(push=>push.messages[0].quickReply.items[0].action.label),['輸入訊息','重新查詢','事假']);
  assert.match(h.pushes[1].messages[0].text,/原本的班｜2026-11-05 ・ 玉門/);
  h.fail.add(uid);
  const fallback=run('contact','contact-event');
  assert.equal(fallback.code,'contact_open');
  assert.equal(fallback.quickReply.items[1].action.label,'結束對話');
});

test('member contact cards support private replies in both directions and reject a copied reply token',()=>{
  const {ctx,pushes}=followupHarness();
  vm.runInContext(menuActionsCode,ctx);
  const roster=ctx.employeeIndex_();
  const otherUid='U'+'c'.repeat(32);
  roster.A05={name:'團員乙',uid:otherUid};
  ctx.employeeIndex_=()=>roster;
  ctx.openBook_=()=>({getSpreadsheetTimeZone:()=> 'Asia/Taipei'});
  ctx.Utilities.formatDate=()=> '2026-11-05';
  const run=input=>ctx.masaluCalendarMenuAction_(input);
  const action=()=>pushes.at(-1).messages[0].contents.footer.contents[0].action;
  assert.equal(run({lineUid:managerUid,choice:'send',value:'A04',text:'幾點到？',eventId:'first'}).code,'sent');
  assert.equal(pushes[0].to,uid);
  assert.equal(action().label,'回覆');
  assert.equal(action().inputOption,'openKeyboard');
  const [,target,proof]=action().fillInText.trim().split(' ');
  assert.equal(target,'A01');
  const reply={choice:'reply_contact',value:target,proof,text:'八點到',eventId:'second'};
  assert.equal(run({...reply,lineUid:otherUid}).code,'unavailable');
  assert.equal(run({...reply,lineUid:uid,text:''}).code,'invalid');
  assert.equal(run({...reply,lineUid:uid}).message,'回覆已傳給 團長。');
  assert.equal(pushes[1].to,managerUid);
  assert.match(pushes[1].messages[0].contents.body.contents[1].text,/八點到/);
  const [,returnTarget,returnProof]=action().fillInText.trim().split(' ');
  assert.equal(run({...reply,lineUid:managerUid,value:returnTarget,proof:returnProof,text:'收到',eventId:'third'}).code,'sent');
  assert.deepEqual(pushes.map(push=>push.to),[uid,managerUid,uid]);
});

test('leave supports type and optional note, validates owner and version, and deduplicates',()=>{
  const h=followupHarness();
  const run=fields=>h.ctx.handleShiftFollowup_({...h.input,...fields});
  assert.equal(run({choice:'leave',lineUid:managerUid}).code,'wrong_user');
  assert.equal(run({choice:'leave',version:2}).code,'expired');
  const options=run({choice:'leave'}).quickReply.items;
  assert.deepEqual(Array.from(options,i=>i.action.label),['事假','病假','其他']);
  const note=run({choice:'leave_sick'});
  assert.equal(note.quickReply.items[0].action.fillInText,'排班留言 ABCDEF01 ');
  const submitted=run({text:'排班留言 ABCDEF01 =not a formula'});
  assert.equal(submitted.code,'request_sent');
  assert.match(submitted.message,/流水號：L-000001/);
  assert.equal(h.requests[0][9],'病假');
  assert.equal(h.requests[0][10],"'=not a formula");
  assert.equal(h.requests[0][11],'待處理');
  const card=h.pushes[0].messages[0];
  assert.match(card.contents.body.contents[0].text,/員工甲[\s\S]*2026-11-05[\s\S]*=not a formula[\s\S]*是否核准/);
  assert.deepEqual(card.contents.footer.contents.map(button=>button.action.label),['核准','不核准']);
  const repeat=run({choice:'leave_personal_send'});
  assert.equal(repeat.code,'request_already');
  assert.equal(repeat.serial,submitted.serial);
  assert.equal(h.requests.length,1);
  assert.equal(h.pushes.length,1);
  assert.equal(h.row[7],'同意');
});

test('managers can review a legacy leave request once and retry an undelivered result',()=>{
  const h=followupHarness();
  h.ledger.appendRow([uuid+':1:leave','L-000001']);
  h.requests.push([uuid+':1:leave',new Date(),uuid,1,'A04','員工甲','2026-11-05','玉門','請假申請','病假','感冒','待處理','A01']);
  const review=(decision,lineUid=managerUid)=>h.ctx.masaluLeaveReview_({lineUid,serial:'L-000001',decision});
  assert.equal(review('approve',uid).code,'unauthorized');
  assert.equal(review().quickReply.items.length,2);
  h.fail.add(uid);
  assert.equal(review('approve').code,'leave_result_pending');
  assert.equal(h.requests[0][11],'已核准');
  assert.equal(review('reject').code,'already');
  assert.equal(review().quickReply.items[0].action.label,'重試通知');
  h.fail.delete(uid);
  assert.equal(review('approve').code,'leave_reviewed');
  assert.match(h.pushes.at(-1).messages[0].text,/L-000001[\s\S]*已核准/);
  const deliveries=h.pushes.length;
  assert.equal(review('approve').code,'leave_reviewed');
  assert.equal(h.pushes.length,deliveries);
});

test('leave approval immediately offers available names, skip and add, and rechecks before assigning',()=>{
  const h=followupHarness();
  const replacement=replacementHarness();
  h.ctx.shiftSheet_=replacement.ctx.shiftSheet_;
  h.ctx.findShiftRow_=replacement.ctx.findShiftRow_;
  h.ctx.employeeIndex_=replacement.ctx.employeeIndex_;
  h.ctx.sendShiftRow_=replacement.ctx.sendShiftRow_;
  replacement.roster.E0.name='勇翰';
  replacement.roster.E1.name='王玲';
  replacement.row[7]='同意';
  replacement.records.push([false,'other','E2','2026-11-05','','','田園','同意']);
  h.ledger.appendRow([uuid+':1:leave','L-000001']);
  h.requests.push([uuid+':1:leave',new Date(),uuid,1,'A04','員工甲','2026-11-05','玉門','請假申請','事假','','待處理']);
  const input={lineUid:managerUid,serial:'L-000001'};
  const act=(decision,extra)=>h.ctx.masaluLeaveReplacement_({...input,decision,...extra});
  const approved=h.ctx.masaluLeaveReview_({...input,decision:'approve'});
  assert.deepEqual(Array.from(approved.quickReply.items,item=>item.action.label),['略過','團長','勇翰','王玲','新增']);
  assert.equal(act('select',{staffId:'E2'}).code,'leave_replacement_choices');
  const add=approved.quickReply.items.at(-1).action;
  assert.equal(add.type,'uri');
  assert.equal(add.uri,'https://liff.line.me/2010020470-VBPZRUQ9?open=add-staff&date=2026-11-05');
  assert.equal(act('add').quickReply.items[0].action.uri,add.uri);
  assert.equal(act('select',{staffId:'E1'}).code,'leave_replaced');
  assert.equal(replacement.row[7],'已請假');
  assert.deepEqual(Array.from(replacement.records.at(-1).slice(2,8)),['E1','2026-11-05','','','玉門','待回覆']);
  assert.equal(act('skip').code,'already');
  assert.equal(replacement.sends.length,1);
});

test('contact sends only messages in the active conversation to the three managers',()=>{
  const h=followupHarness();
  const manager2='U'+'c'.repeat(32);
  const manager3='U'+'d'.repeat(32);
  const roster=h.ctx.employeeIndex_();
  roster.A02={name:'副團長',uid:manager2};
  roster.A03={name:'副團長',uid:manager3};
  h.ctx.employeeIndex_=()=>roster;
  const run=fields=>h.ctx.handleShiftFollowup_({...h.input,...fields});
  const text={kind:'contact_message',text:'請問幾點集合？',eventId:'msg-1'};
  assert.equal(run({...text,timestamp:1}).code,'ignored');
  const start=run({choice:'contact',timestamp:2});
  assert.equal(start.code,'contact_open');
  assert.deepEqual(Array.from(start.quickReply.items,item=>item.action.label),['輸入訊息','結束對話']);
  assert.equal(run({...text,timestamp:1}).code,'ignored');
  assert.equal(run({...text,timestamp:3}).code,'contact_sent');
  assert.deepEqual(h.pushes.map(p=>p.to),[managerUid,manager2,manager3]);
  assert.match(h.pushes[0].messages[0].text,/員工甲[\s\S]*請問幾點集合？/);
  assert.equal(run({choice:'contact_end',timestamp:4}).code,'contact_ended');
  assert.equal(run({choice:'contact',timestamp:2}).code,'ignored');
  assert.equal(run({...text,timestamp:5}).code,'ignored');
  assert.equal(h.pushes.length,3);
});

test('swap dates list only eligible assignments, paginate and recover from obsolete date pickers',()=>{
  const h=followupHarness();
  const otherUid='U'+'c'.repeat(32);
  const roster={A04:{uid},A05:{uid:otherUid},A06:{uid:''}};
  const shift={id:uuid,version:1,date:'2026-11-05',shift:'玉門'};
  const row=(date,staff='A05',status='同意')=>[false,'22222222-2222-2222-2222-222222222222',staff,date,'','','田園',status,'','','',1];
  let rows=[h.row,row('2026-11-01'),row('2026-12-01','A05','待回覆'),row('2026-12-02','A05','不同意'),
    row('2026-12-03','A06'),row('2026-12-04','A04'),row('bad-date'),
    ...Array.from({length:15},(_,i)=>row('2026-11-'+String(5+i).padStart(2,'0'))),row('2026-11-05')];
  h.ctx.masaluDate_=value=>value instanceof Date?'2026-11-05':/^\d{4}-\d{2}-\d{2}$/.test(String(value))?String(value):'';
  const sheet={getLastRow:()=>rows.length+1,getRange:()=>({getValues:()=>rows})};
  const choose=input=>h.ctx.masaluSwapSelection_(input,shift,roster,uid,sheet);
  const first=choose({choice:'swap'});
  assert.deepEqual(Array.from(first.quickReply.items,item=>item.action.label),[
    '2026-11-05','2026-11-06','2026-11-07','2026-11-08','2026-11-09',
    '2026-11-10','2026-11-11','2026-11-12','2026-11-13','2026-11-14','下一頁']);
  const second=choose({choice:'swap',page:1});
  assert.deepEqual(Array.from(second.quickReply.items,item=>item.action.label),[
    '2026-11-15','2026-11-16','2026-11-17','2026-11-18','2026-11-19','上一頁']);
  const stale=choose({choice:'swap_date',date:'2026-12-01'});
  assert.match(stale.message,/2026-12-01.*沒有可調換/);
  assert.equal(stale.quickReply.items[0].action.label,'2026-11-05');
  assert.match(stale.quickReply.items[0].action.data,/choice=swap_date&date=2026-11-05/);
  rows=[h.row];
  const empty=choose({choice:'swap'});
  assert.equal(empty.code,'swap_no_candidates');
  assert.match(empty.message,/已綁定 LINE/);
  assert.deepEqual(Array.from(empty.quickReply.items,item=>item.action.label),['重新查詢','聯絡團長']);
});

for (const decision of ['agree','decline']) test('swap '+decision+' only registers a case after the intended recipient agrees',()=>{
  const h=followupHarness(), invites=[];
  h.row[4]='14:00';
  h.row[5]='17:00';
  const targetId='22222222-2222-2222-2222-222222222222';
  const targetUid='U'+'c'.repeat(32);
  const roster=h.ctx.employeeIndex_();
  roster.A05={name:'對方',uid:targetUid};
  h.ctx.employeeIndex_=()=>roster;
  h.ctx.masaluDate_=value=>value instanceof Date?'2026-11-05':String(value);
  const target=[false,targetId,'A05','2026-11-06','18:00','20:00','田園','同意','','已發送','',2];
  h.ctx.findShiftRow_=(_sheet,id)=>[uuid,targetId].indexOf(id)+2;
  h.ctx.shiftSheet_=()=>({getLastRow:()=>3,getRange:(row,_column,count)=>({
    getValues:()=>count===2?[h.row,target]:[row===2?h.row:target]
  })});
  const inviteSheet={appendRow:values=>invites.push(Array.from(values)),getRange:(row,column)=>({
    setValue:value=>{invites[row-2][column-1]=value;}
  })};
  const readTable=h.ctx.masaluTable_;
  h.ctx.masaluTable_=(name,headers,book)=>name==='MaSaLu調班邀請'
    ? {sheet:inviteSheet,rows:invites}:readTable(name,headers,book);
  const getBook=h.ctx.openBook_;
  h.ctx.openBook_=()=>({getSheetByName:name=>name==='MaSaLu調班邀請'?inviteSheet:getBook().getSheetByName(name)});

  const choose=fields=>h.ctx.handleShiftFollowup_({...h.input,...fields});
  const dates=choose({choice:'swap'});
  assert.equal(dates.quickReply.items[0].action.label,'2026-11-06');
  assert.match(dates.message,/原本的班｜2026-11-05 ・ 玉門/);
  assert.match(dates.message,/14:00–17:00/);
  const date='2026-11-06';
  const shifts=choose({choice:'swap_date',date});
  assert.equal(shifts.quickReply.items[0].action.label,'田園 ・ 18:00–20:00');
  assert.match(shifts.message,/原本的班｜2026-11-05 ・ 玉門[\s\S]*想換的日期｜2026-11-06/);
  const people=choose({choice:'swap_shift',date,target:targetId});
  assert.equal(people.quickReply.items[0].action.label,'對方');
  assert.match(people.message,/原本的班｜2026-11-05 ・ 玉門[\s\S]*想換的班｜2026-11-06 ・ 田園/);
  const submit={choice:'swap_target',date,target:targetId,targetVersion:2};
  assert.equal(choose({...submit,targetVersion:1}).code,'swap_expired');
  const invited=choose(submit);
  assert.equal(invited.code,'swap_invited');
  assert.match(invited.message,/原本的班｜2026-11-05 ・ 玉門[\s\S]*想換的班｜2026-11-06 ・ 田園[\s\S]*調班對象｜對方/);
  assert.match(invited.message,/18:00–20:00/);
  assert.equal(h.requests.length,0);
  assert.equal(h.pushes[0].to,targetUid);
  const actions=h.pushes[0].messages[0].contents.footer.contents;
  assert.deepEqual(actions.map(item=>item.action.label),['同意調班','不同意調班']);
  const answer={lineUid:targetUid,id:invites[0][0],decision};
  assert.equal(h.ctx.masaluSwapResponse_({...answer,lineUid:managerUid}).code,'wrong_user');
  assert.equal(h.ctx.masaluSwapResponse_(answer).code,'swap_answered');
  assert.equal(h.requests.length,decision==='agree'?1:0);
  assert.equal(h.pushes.filter(push=>push.to===managerUid).length,decision==='agree'?1:0);
  if (decision==='agree') assert.equal(h.requests[0][11],'待團長審核');
  assert.equal(h.ctx.masaluSwapResponse_(answer).code,'swap_answered');
  assert.equal(h.requests.length,decision==='agree'?1:0);
  assert.equal(h.pushes.filter(push=>push.to===managerUid).length,decision==='agree'?1:0);
  assert.equal(target[2],'A05');
  assert.equal(h.row[2],'A04');
});

test('temporary staff can use their card after their UID is cleared, copied buttons are rejected',()=>{
  const h=followupHarness();
  h.row[2]='';
  h.row[7]='待回覆';
  h.row[12]='臨時員工';
  h.row[13]=uid;
  h.ctx.masaluDeliverShiftReceipt_=()=>false;
  const result=h.ctx.handleShiftReply_({...h.input,choice:'agree'});
  assert.equal(h.row[13],'');
  const proof=result.replyMessage.contents.footer.contents[0].action.data.split('&proof=')[1];
  assert.match(proof,/^[A-Za-z0-9_-]{43}$/);
  assert.equal(h.ctx.handleShiftFollowup_({...h.input,lineUid:managerUid,choice:'contact',token:proof}).code,'wrong_user');
  assert.equal(h.ctx.handleShiftFollowup_({...h.input,choice:'contact',token:proof}).code,'contact_open');
  assert.ok(!JSON.stringify(h.requests).includes(uid));
});


test('employee receipt is sent to the actual employee before the separate manager notice',()=>{
  const h=harness(), deliveries=[];
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-token'})};
  h.ctx.Utilities={
    DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},
    computeDigest:(algorithm,value)=>Array.from(createHash(algorithm).update(value).digest())
  };
  h.ctx.UrlFetchApp={fetch(_url,options){
    deliveries.push({body:JSON.parse(options.payload),key:options.headers['X-Line-Retry-Key']});
    return deliveries.length===1?{getResponseCode:()=>200}:
      {getResponseCode:()=>409,getAllHeaders:()=>({'x-line-accepted-request-id':'existing'})};
  }};
  h.ctx.masaluNotifyManagers_=()=>{
    assert.equal(deliveries.length,1);
    return {sent:1,failed:0};
  };
  const input={id:uuid,version:1,choice:'agree',lineUid:uid,eventId:'line-event'};
  const result=h.ctx.handleShiftReply_(input);
  assert.equal(result.serial,'A-000001');
  assert.equal(result.message,undefined);
  assert.equal(deliveries[0].body.to,uid);
  assert.equal(deliveries[0].body.messages[0].type,'flex');
  const repeated=h.ctx.handleShiftReply_(input);
  assert.equal(repeated.message,undefined);
  assert.equal(deliveries[0].key,deliveries[1].key);
});

test('a manager agreeing to their own shift gets one confirmation card and still receives the day summary',()=>{
  const h=harness({mockManagerNotice:false}), deliveries=[];
  h.row[2]='A01';
  h.ctx.employeeIndex_=()=>({A01:{name:'小米',uid},A02:{name:'團長乙',uid:managerUid}});
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-token'})};
  h.ctx.masaluLineRetryKey_=value=>value;
  h.ctx.UrlFetchApp={fetch(_url,options){
    deliveries.push(JSON.parse(options.payload));
    return {getResponseCode:()=>200};
  }};
  h.ctx.handleShiftReply_({id:uuid,version:1,choice:'agree',lineUid:uid,eventId:'event'});
  const ownMessages=deliveries.filter(delivery=>delivery.to===uid).flatMap(delivery=>delivery.messages);
  assert.deepEqual(ownMessages.map(message=>message.type),['flex','text']);
  assert.match(ownMessages[0].altText,/小米 已同意排班/);
  assert.match(ownMessages[1].text,/總排班已全部確認/);
  const otherMessages=deliveries.find(delivery=>delivery.to===managerUid).messages;
  assert.equal(otherMessages.length,2);
  assert.match(otherMessages[0].text,/小米 已同意排班/);
  assert.equal(otherMessages[1].text,ownMessages[1].text);
});

test('a manager can resend a declined shift once, failed delivery preserves decline and a new answer gets a new serial',()=>{
  const h=harness(), sends=[];
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-token'})};
  h.ctx.Utilities={
    DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},
    computeDigest:(algorithm,value)=>Array.from(createHash(algorithm).update(value).digest())
  };
  h.ctx.masaluDeliverShiftReceipt_=()=>false;
  const response={lineUid:uid,id:uuid,version:1,choice:'disagree'};
  assert.equal(h.ctx.handleShiftReply_(response).serial,'D-000001');
  assert.equal(h.ctx.handleShiftReply_({...response,choice:'agree'}).code,'already');
  h.row[16]='原備註｜[管理已略過]';
  const resend={lineUid:managerUid,id:uuid,version:1,decision:'resend'};
  assert.equal(h.ctx.handleManagerDecision_({...resend,lineUid:uid}).code,'unauthorized');
  h.row[16]+='[替補ID:22222222-2222-2222-2222-222222222222]';
  assert.equal(h.ctx.handleManagerDecision_(resend).code,'replacement_exists');
  h.row[16]='原備註｜[管理已略過]';
  h.ctx.UrlFetchApp={fetch(_url,options){
    sends.push(options);
    return {getResponseCode:()=>sends.length===1?500:200,getContentText:()=>''};
  }};
  assert.equal(h.ctx.handleManagerDecision_(resend).code,'failed');
  assert.equal(h.row[7],'不同意');
  assert.equal(h.row[11],1);
  assert.equal(h.ctx.handleManagerDecision_(resend).code,'sent');
  assert.equal(sends[0].headers['X-Line-Retry-Key'],sends[1].headers['X-Line-Retry-Key']);
  const card=JSON.parse(sends[1].payload).messages[0];
  assert.match(card.contents.footer.contents[0].action.data,/v=2&choice=agree$/);
  assert.equal(h.row[7],'待回覆');
  assert.equal(h.row[16],'原備註');
  assert.equal(h.ctx.handleManagerDecision_(resend).code,'expired');
  assert.equal(h.ctx.handleShiftReply_(response).code,'expired');
  assert.equal(h.ctx.handleShiftReply_({...response,version:2}).serial,'D-000002');
  assert.equal(h.serialRows.length,2);
});


test('LINE rejects the rich receipt, so GAS sends the same serial and buttons as text',()=>{
  const h=harness(), messages=[];
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-token'})};
  h.ctx.masaluLineRetryKey_=value=>value;
  h.ctx.UrlFetchApp={fetch(_url,options){
    const body=JSON.parse(options.payload);
    assert.equal(body.to,uid);
    messages.push(body.messages[0]);
    return {getResponseCode:()=>messages.length===1?400:200};
  }};
  const result=h.ctx.handleShiftReply_({id:uuid,version:1,choice:'agree',lineUid:uid,eventId:'event'});
  assert.equal(result.message,undefined);
  assert.equal(messages[0].type,'flex');
  assert.match(messages[1].text,/\nA-000001$/);
  assert.deepEqual(messages[1].quickReply.items.map(item=>item.action.label),['聯絡團長','調班請求','我要請假']);
});
