import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
import {createHash,createHmac} from 'node:crypto';

const code = await readFile(new URL('../GAS/程式碼.js', import.meta.url), 'utf8');
const calendarCode = await readFile(new URL('../GAS/CalendarApi.js', import.meta.url), 'utf8');
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
  const ledger={getDataRange:()=>({getValues:()=>[vm.runInContext('MASALU_SERIAL_HEADERS',ctx),...serialRows]}),
    appendRow:values=>serialRows.push(Array.from(values))};
  ctx.openBook_=()=>({getSheetByName:name=>name==='MaSaLu流水號'?ledger:null});
  ctx.shiftSheet_=()=>sheet;
  ctx.findShiftRow_=()=>2;
  ctx.employeeIndex_=()=>({A04:{name:'員工甲',uid},A01:{name:'團長',uid:managerUid}});
  ctx.masaluDate_=x=>String(x);
  ctx.safeLog_=(...items)=>logs.push(items);
  if (mockManagerNotice) ctx.masaluNotifyManagers_=x=>{notices.push(x);return {sent:1,failed:0};};
  return {ctx,row,logs,notices,serialRows,ledger};
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


function followupHarness() {
  const h=harness(), requests=[], cache=new Map(), pushes=[];
  h.row[7]='同意';
  const sheet={
    getLastRow:()=>requests.length+1,
    appendRow:row=>requests.push(Array.from(row)),
    getRange:(row,col)=>({setValue(value){requests[row-2][col-1]=value;}})
  };
  const readTable=h.ctx.masaluTable_;
  h.ctx.openBook_=()=>({getSheetByName:name=>name==='MaSaLu流水號'?h.ledger:sheet});
  h.ctx.masaluTable_=(name,headers,book)=>name==='MaSaLu流水號'?readTable(name,headers,book):{sheet,rows:requests};
  h.ctx.CacheService={getScriptCache:()=>({get:key=>cache.get(key),put:(key,value)=>cache.set(key,value)})};
  h.ctx.PropertiesService={getScriptProperties:()=>({getProperty:()=> 'test-secret'})};
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
  assert.match(h.pushes[0].messages[0].text,/員工甲[\s\S]*2026-11-05[\s\S]*=not a formula/);
  const repeat=run({choice:'leave_personal_send'});
  assert.equal(repeat.code,'request_already');
  assert.equal(repeat.serial,submitted.serial);
  assert.equal(h.requests.length,1);
  assert.equal(h.pushes.length,1);
  assert.equal(h.row[7],'同意');
});

test('contact and swap are independent; retry only resends to unnotified managers',()=>{
  const h=followupHarness();
  const manager2='U'+'c'.repeat(32);
  const roster=h.ctx.employeeIndex_();
  roster.A02={name:'副團長',uid:manager2};
  h.ctx.employeeIndex_=()=>roster;
  h.fail.add(manager2);
  const run=choice=>h.ctx.handleShiftFollowup_({...h.input,choice});
  assert.equal(run('contact').code,'request_saved');
  h.fail.clear();
  assert.equal(run('contact').code,'request_sent');
  assert.deepEqual(h.pushes.map(p=>p.to),[managerUid,manager2,manager2]);
  assert.equal(run('swap').code,'request_note');
  assert.equal(run('swap_send').code,'request_sent');
  assert.deepEqual(h.requests.map(r=>r[8]),['聯絡團長','調班請求']);
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
  assert.equal(h.ctx.handleShiftFollowup_({...h.input,choice:'contact',token:proof}).code,'request_sent');
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
