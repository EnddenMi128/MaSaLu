/**
 * MaSaLu 團服確認：白色／黑色／暫不選。
 * Dates and the noon reminder use Asia/Taipei (UTC+08:00).
 * A Worker Cron calls masaluUniformTick_ once per minute; no user webhook can invoke it
 * without the existing GAS_SHARED_SECRET.
 */
const MASALU_UNIFORM_TZ = 'Asia/Taipei';

function masaluUniformLabel_(uniform) {
  return uniform==='white'?'白色團服':uniform==='black'?'黑色團服':'未擇衣';
}

function masaluUniformRecord_(date) {
  const table=masaluUniformTable_();
  const index=table.rows.findIndex(r=>masaluDate_(r[0])===date);
  return {table,index,row:index>=0?table.rows[index]:null};
}

function masaluUniformDisplay_(date) {
  if(!date)return '';
  const record=masaluUniformRecord_(date).row;
  if(!record || record[2]==='cleared' || record[2]==='pending')return '';
  return masaluUniformLabel_(String(record[1]||''));
}

function masaluLinePushUniform_(uid,message) {
  const token=PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if(!token)throw new Error('缺少 LINE_CHANNEL_ACCESS_TOKEN');
  if(!validLineUid_(uid))throw new Error('尚未綁定有效 LINE 帳號');
  const response=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
    method:'post',contentType:'application/json; charset=UTF-8',
    headers:{Authorization:'Bearer '+token},
    payload:JSON.stringify({to:uid,messages:[message]}),muteHttpExceptions:true
  });
  if(response.getResponseCode()<200||response.getResponseCode()>=300)
    throw new Error('LINE HTTP '+response.getResponseCode());
}

/** Send a single A02 choice card. Initial and noon reminders have separate dedupe fields. */
function masaluNotifyA02Uniform_(date,reason) {
  const record=masaluUniformRecord_(date),r=record.row;
  if(!r||!['pending','not_selected'].includes(String(r[2])))return {sent:false,skipped:true};
  const version=Number(r[6])||0;
  if(reason==='initial'&&Number(r[4])===version)return {sent:false,skipped:true};
  if(reason==='noon'&&String(r[5])===date)return {sent:false,skipped:true};
  const a02=employeeIndex_().A02;
  if(!a02||!validLineUid_(a02.uid))throw new Error('A02 尚未綁定 LINE，無法通知選團服');
  const button=(name,color,value)=>({
    type:'button',style:'primary',color,
    action:{type:'postback',label:name,displayText:'團服：'+name,
      data:'action=uniform_choice&date='+date+'&v='+version+'&uniform='+value}
  });
  const body=reason==='noon'
    ? '今天 12:00 再次提醒：團服仍未選擇，請確認當日團服。'
    : '請於 5 分鐘內選擇團服；逾時會先發送「未擇衣」排班。';
  const message={type:'flex',altText:'👕 '+date+' 團服待確認',contents:{
    type:'bubble',size:'kilo',body:{type:'box',layout:'vertical',spacing:'md',contents:[
      {type:'text',text:'👕 團服待確認',weight:'bold',size:'lg',color:'#D96B24'},
      {type:'text',text:'日期：'+date,wrap:true,size:'sm'},
      {type:'text',text:body,wrap:true,size:'sm'}
    ]},footer:{type:'box',layout:'horizontal',spacing:'sm',contents:[
      button('白色團服','#64748B','white'),button('黑色團服','#1F2937','black')
    ]}
  }};
  masaluLinePushUniform_(a02.uid,message);
  // Only mark after a successful API response; a failed send can be retried.
  const cell=reason==='noon'?6:5;
  record.table.sheet.getRange(record.index+2,cell).setValue(reason==='noon'?date:version);
  // A date created after noon already has its first A02 reminder; don't
  // immediately send a second copy on the next minute tick.
  if(reason==='initial' && Utilities.formatDate(new Date(),MASALU_UNIFORM_TZ,'yyyy-MM-dd')===date &&
      Utilities.formatDate(new Date(),MASALU_UNIFORM_TZ,'HH:mm')>='12:00')
    record.table.sheet.getRange(record.index+2,6).setValue(date);
  safeLog_(reason==='noon'?'團服中午提醒':'團服待選通知','', 'A02','成功',date,'');
  return {sent:true};
}

function masaluNotifyUniformChange_(date,uniform) {
  const sheet=shiftSheet_(),last=sheet.getLastRow(),roster=employeeIndex_();
  const summary={sent:0,failed:0,skipped:0};
  if(last<2)return summary;
  const delivered=new Set();
  sheet.getRange(2,1,last-1,SHIFT_APP.SHIFT_HEADERS.length).getValues().forEach(row=>{
    if(masaluDate_(row[3])!==date || String(row[7])==='不同意'||String(row[9])!=='已發送')return;
    const person=roster[String(row[2]||'')];
    const uid=person?.uid || String(row[13]||'');
    if(!validLineUid_(uid)||delivered.has(uid)){summary.skipped++;return;}
    delivered.add(uid);
    try {
      masaluLinePushUniform_(uid,{type:'text',text:'👕 團服異動通知\n日期：'+date+'\n最新團服：'+masaluUniformLabel_(uniform)+'\n原排班同意狀態不變。'});
      summary.sent++;
    } catch(err) {
      summary.failed++;
      safeLog_('團服異動通知','',String(row[2]||''),'失敗',String(err.message).slice(0,160),'');
    }
  });
  return summary;
}

/** Release any previously deferred unsent rows, preserving answered rows and versions. */
function masaluSendDeferredUniform_(date) {
  return masaluNotifyAddedForDate_({date,addedIds:[],releasePending:true,deferNotifications:false});
}

/** Only the bound A02 LINE identity can choose via the reminder's postback. */
function handleUniformChoice_(input) {
  const uid=String(input.lineUid||''),date=masaluDate_(input.date),uniform=String(input.uniform||'');
  const version=Number(input.version);
  if(!validLineUid_(uid)||!date||!['white','black'].includes(uniform)||!Number.isSafeInteger(version)||version<1)
    return {ok:true,code:'invalid',message:'團服選擇格式不正確。'};
  if(employeeIndex_().A02?.uid!==uid)
    return {ok:true,code:'unauthorized',message:'只有 A02 可以透過此訊息選擇團服。'};
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  let changed=false;
  try {
    const data=masaluLoad_(),old=masaluUniformFor_(data,date);
    if(!old.row||old.status==='cleared')return {ok:true,code:'missing',message:'這天沒有待確認的排班。'};
    if(version!==old.version)return {ok:true,code:'expired',message:'這張團服通知已更新，請使用最新的通知。'};
    if(old.status==='confirmed')return {ok:true,code:'already',message:'當日團服已確認為'+masaluUniformLabel_(old.uniform)+'。'};
    masaluSetUniformForDate_(data,date,uniform,false,true);
    changed=true;
    safeLog_('A02選擇團服','', 'A02','成功',date+' '+masaluUniformLabel_(uniform),String(input.eventId||''));
  } finally {lock.releaseLock();}
  if(!changed)return {ok:true,code:'already',message:'團服狀態未變更。'};
  const update=masaluNotifyUniformChange_(date,uniform);
  const released=masaluSendDeferredUniform_(date);
  return {ok:true,code:'confirmed',message:'✅ '+date+' 已設定'+masaluUniformLabel_(uniform)+'。\n新排班通知 '+released.sent+' 人，團服異動通知 '+update.sent+' 人。',update,released};
}

/** Called by Cloudflare every minute: five-minute expiry and 12:00 Taiwan reminder. */
function masaluUniformTick_() {
  const now=new Date(),today=Utilities.formatDate(now,MASALU_UNIFORM_TZ,'yyyy-MM-dd');
  const hhmm=Utilities.formatDate(now,MASALU_UNIFORM_TZ,'HH:mm');
  const expired=[],pendingSend=[],noon=[];
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const table=masaluUniformTable_();
    table.rows.forEach((row,index)=>{
      const date=masaluDate_(row[0]),status=String(row[2]||'');
      if(!date||!['pending','not_selected'].includes(status))return;
      const deadline=row[3] instanceof Date?row[3].getTime():new Date(row[3]).getTime();
      if(status==='pending'&&Number.isFinite(deadline)&&now.getTime()>=deadline){
        table.sheet.getRange(index+2,3).setValue('not_selected');
        expired.push({date,previousUniform:String(row[7]||'')});
      }
      if(status==='not_selected'||expired.some(x=>x.date===date))pendingSend.push(date);
      if(date===today&&hhmm>='12:00'&&String(row[5]||'')!==today)noon.push(date);
    });
  } finally {lock.releaseLock();}
  const summary={ok:true,expired:expired.length,released:0,reminded:0,failed:0};
  expired.forEach(item=>{
    if(item.previousUniform){
      const result=masaluNotifyUniformChange_(item.date,'');
      summary.failed+=result.failed;
    }
  });
  [...new Set(pendingSend)].forEach(date=>{
    const result=masaluSendDeferredUniform_(date);
    summary.released+=result.sent;
    summary.failed+=result.failed;
  });
  [...new Set(noon)].forEach(date=>{
    try {if(masaluNotifyA02Uniform_(date,'noon').sent)summary.reminded++;}
    catch(err){summary.failed++;safeLog_('團服中午提醒','', 'A02','失敗',String(err.message).slice(0,160),'');}
  });
  return summary;
}
