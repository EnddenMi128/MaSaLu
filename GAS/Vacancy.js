const MASALU_VACANCY_SHEET = 'MaSaLu補班申請';
const MASALU_VACANCY_HEADERS = ['申請ID','空缺識別碼','員工編號','建立時間','狀態','日期','班別'];

function masaluVacancyButton_(label, action, id) {
  return {type:'action',action:{type:'postback',label:label.slice(0,20),
    data:'action=vacancy&choice='+action+'&id='+id}};
}

function masaluVacancySlots_(rows, approvedLeaves, offNames, today, timeZone) {
  const replacements = new Map(rows.map(row=>[String(row[16] || ''),row]));
  const leaves = new Map(approvedLeaves.map(({row,serial})=>[String(row[2]),{row,serial}]));
  const slots = [];
  for (const row of rows) {
    const date = masaluDate_(row[3],timeZone);
    const shift = String(row[6] || '').trim();
    const id = String(row[1] || '');
    if (date < today || !shift || offNames.has(shift) || !/^[a-f0-9-]{36}$/i.test(id)) continue;
    const note = String(row[16] || '');
    if (note.includes(MASALU_REASSIGNED_NOTE)) continue;
    const leave = leaves.get(id);
    let marker;
    let closed;
    let key;
    if (row[7]==='不同意') {
      marker = '替補來源：'+id;
      closed = note.includes('[管理已略過]');
      key = id+':'+Number(row[11]);
    } else if (leave && ['同意','已請假'].includes(row[7]) &&
        Number(leave.row[3])===Number(row[11]) && masaluDate_(leave.row[6],timeZone)===date &&
        leave.row[7]===row[6] && String(leave.row[4] || '')===String(row[2] || '')) {
      marker = '請假替補來源：'+leave.serial;
      closed = note.includes('[請假略過:'+leave.serial+']') || note.includes('[請假已替補:'+leave.serial+']');
      key = leave.serial;
    } else continue;
    const replacement = replacements.get(marker);
    if (replacement && ['同意','不同意','已請假'].includes(replacement[7])) continue;
    if (!replacement && closed) continue;
    slots.push({id,key,date,shift,row,marker,leave:marker.startsWith('請假')?leave:null,replacement});
  }
  return slots.sort((a,b)=>a.date.localeCompare(b.date)||a.shift.localeCompare(b.shift));
}

function masaluVacancy_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let result;
  try {
    result = masaluVacancyLocked_(input);
  } finally {
    lock.releaseLock();
  }
  if (!result.shiftId) return result;
  const sent = sendShiftRow_(null,false,result.shiftId);
  if (['sent','skipped'].includes(sent.code)) {
    return {ok:true,code:'approved',message:'補班已核准，通知已送出'};
  }
  return {ok:true,code:'approved_pending',message:'補班已登記，排班通知尚未送出。',
    quickReply:{items:[masaluVacancyButton_('重試發送','approve',result.shiftId)]}};
}

function masaluVacancyLocked_({lineUid,choice,id}) {
  const reply = (code,message,items=[])=>({ok:true,code,message,
    ...(items.length?{quickReply:{items}}:{})});
  const roster = employeeIndex_();
  const accounts = Object.keys(roster).filter(key=>roster[key].uid===lineUid);
  if (!validLineUid_(lineUid) || accounts.length!==1) return reply('unbound','請先綁定團員帳號。');
  const employeeId = accounts[0];
  const managers = ['A01','A02','A03'];
  const manager = managers.includes(employeeId);
  const book = openBook_();
  const timeZone = book.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(),timeZone,'yyyy-MM-dd');
  const sheet = shiftSheet_();
  const rows = sheet.getLastRow()<2?[]:sheet.getRange(2,1,sheet.getLastRow()-1,SHIFT_APP.SHIFT_HEADERS.length).getValues();
  const config = book.getSheetByName(MASALU.SHIFTS);
  const offNames = new Set(['休假']);
  if (config) for (const row of config.getDataRange().getValues().slice(1)) {
    if (String(row[2]).toLowerCase()==='true') offNames.add(String(row[1]).trim());
  }
  const leaves = [];
  const followupName = 'MaSaLu排班請求';
  if (book.getSheetByName(followupName) && book.getSheetByName(MASALU_SERIAL_SHEET)) {
    const ledger = new Map(masaluTable_(MASALU_SERIAL_SHEET,MASALU_SERIAL_HEADERS,book).rows.map(row=>[row[0],row[1]]));
    for (const row of masaluTable_(followupName,MASALU_FOLLOWUP_HEADERS,book).rows) {
      if (row[8]==='請假申請' && row[11]==='已核准' && ledger.has(row[0])) leaves.push({row,serial:ledger.get(row[0])});
    }
  }
  const slots = masaluVacancySlots_(rows,leaves,offNames,today,timeZone);
  const busy = (staffId,date)=>rows.some(row=>String(row[2] || '').trim()===staffId &&
    masaluDate_(row[3],timeZone)===date && row[6] && !offNames.has(String(row[6]).trim()) &&
    !['不同意','已請假'].includes(row[7]));
  const eligible = (slot,staffId)=>!slot.replacement && String(slot.row[2] || '').trim()!==staffId &&
    !busy(staffId,slot.date);
  let applications = book.getSheetByName(MASALU_VACANCY_SHEET);
  const requests = applications ? masaluTable_(MASALU_VACANCY_SHEET,MASALU_VACANCY_HEADERS,book).rows : [];
  const label = slot=>slot.date.slice(5)+' '+slot.shift;
  if (choice==='list') {
    const visible = manager ? slots : slots.filter(slot=>eligible(slot,employeeId));
    const pending = manager ? requests.filter(row=>row[4]==='待審核' && slots.some(slot=>slot.key===row[1] &&
      eligible(slot,String(row[2])))) : [];
    const groups = new Map();
    const ownRequests = new Map(requests.filter(row=>row[2]===employeeId).map(row=>[row[1],row[4]]));
    for (const slot of visible) {
      const key = slot.date+' '+slot.shift;
      if (!groups.has(key)) groups.set(key,{slot,open:0,waiting:0});
      const group = groups.get(key);
      if (ownRequests.get(group.slot.key)==='不核准' && !ownRequests.has(slot.key)) group.slot = slot;
      if (slot.replacement) group.waiting++;
      else group.open++;
    }
    const choices = [
      ...pending.map(row=>({text:masaluDate_(row[5],timeZone)+' '+row[6]+'｜'+(roster[row[2]]?.name || row[2]),
        button:masaluVacancyButton_('審核 '+row[6]+' '+(roster[row[2]]?.name || row[2]),'review',row[0])})),
      ...[...groups].map(([key,{slot,open,waiting}])=>({text:key+'｜'+[open?'缺 '+open+' 人':'',waiting?'待確認 '+waiting+' 人':''].filter(Boolean).join(' ・ '),
        button:manager?null:masaluVacancyButton_('申請 '+label(slot),'apply',slot.key)}))
    ];
    const page = Math.min(Math.max(0,Number(id)||0),Math.max(0,Math.ceil(choices.length/10)-1));
    const shown = choices.slice(page*10,page*10+10);
    const buttons = shown.map(entry=>entry.button).filter(Boolean);
    if (page>0) buttons.push(masaluVacancyButton_('上一頁','list',page-1));
    if ((page+1)*10<choices.length) buttons.push(masaluVacancyButton_('下一頁','list',page+1));
    return reply('vacancies',shown.length?'缺班查詢\n\n'+shown.map(entry=>entry.text).join('\n'):
      '現已滿班',buttons);
  }
  if (choice==='apply') {
    const slot = slots.find(item=>item.key===id);
    if (!slot || !eligible(slot,employeeId)) return reply('unavailable','這個空缺已變更，或你當天已有排班。');
    const existing = requests.find(row=>row[1]===id && row[2]===employeeId);
    if (existing && existing[4]!=='待審核') return reply('already','這筆補班申請已處理。');
    const requestId = existing?.[0] || Utilities.getUuid();
    if (!existing) {
      applications = applications || ensureSheet_(book,MASALU_VACANCY_SHEET,MASALU_VACANCY_HEADERS,false);
      applications.appendRow([requestId,id,employeeId,new Date(),'待審核',slot.date,slot.shift]);
      SpreadsheetApp.flush();
    }
    const review = masaluVacancyButton_('審核補班','review',requestId);
    const notice = {type:'text',text:'補班申請｜'+roster[employeeId].name+'\n'+slot.date+' '+slot.shift,
      quickReply:{items:[review]}};
    const recipients = [...new Set(managers.map(key=>roster[key]?.uid).filter(validLineUid_))];
    const delivered = recipients.map(uid=>masaluPushWorkflowMessage_(uid,notice,'vacancy-review:'+requestId+':'+uid));
    const message = '補班申請已登記\n'+slot.date+' ・ '+slot.shift;
    return delivered.some(Boolean) ? reply('submitted',message) :
      reply('submitted_notice_pending',message+'\n通知未送達',
        [masaluVacancyButton_('重試通知','apply',id)]);
  }
  if (!manager) return reply('unauthorized','只有管理員可以審核補班。');
  const index = requests.findIndex(row=>row[0]===id);
  if (index<0) return reply('missing','找不到這筆補班申請。');
  const request = requests[index];
  const staffId = String(request[2]);
  const details = masaluDate_(request[5],timeZone)+' '+request[6]+'｜'+(roster[staffId]?.name || staffId);
  const slot = slots.find(item=>item.key===request[1]);
  const alreadyScheduled = rows.find(row=>row[1]===id && String(row[2])===staffId);
  if (choice==='approve' && alreadyScheduled) {
    if (request[4]==='已核准') return {shiftId:id};
    if (request[4]==='待審核' && slot?.replacement===alreadyScheduled) {
      masaluCloseVacancySource_(sheet,rows,slot,id);
      applications.getRange(index+2,5).setValue('已核准');
      SpreadsheetApp.flush();
      return {shiftId:id};
    }
    return reply('unavailable','補班紀錄已變更，請重新查詢。');
  }
  if (choice==='reject' && ['待審核','不核准'].includes(request[4])) {
    if (request[7]) return reply('already',details+'｜不核准，已通知申請人。');
    if (request[4]==='待審核') {
      applications.getRange(index+2,5).setValue('不核准');
      SpreadsheetApp.flush();
    }
    const sent = validLineUid_(roster[staffId]?.uid) && masaluPushWorkflowMessage_(roster[staffId].uid,
      {type:'text',text:'補班申請未核准\n'+details},'vacancy-rejected:'+id);
    if (sent) {
      applications.getRange(1,8).setValue('拒絕通知時間');
      applications.getRange(index+2,8).setValue(new Date());
      SpreadsheetApp.flush();
    }
    return sent ? reply('rejected','補班未核准') :
      reply('rejected_notice_pending','補班未核准，通知未送達',[
        masaluVacancyButton_('重試通知','reject',id)
      ]);
  }
  if (request[4]!=='待審核') return reply('already',details+'｜'+request[4]);
  if (!slot || !eligible(slot,staffId) || !validLineUid_(roster[staffId]?.uid)) {
    return reply('unavailable',details+'｜空缺已變更或申請人當天已有排班。');
  }
  if (choice==='review') return reply('review',details+'\n是否核准補班？',[
    masaluVacancyButton_('核准','approve',id),masaluVacancyButton_('拒絕','reject',id)]);
  if (choice!=='approve') return reply('invalid','請從缺班查詢重新操作。');
  const source = slot.row;
  sheet.appendRow([false,id,staffId,source[3],source[4],source[5],source[6],
    '待回覆','','未發送','',0,'','','','',slot.marker]);
  masaluCloseVacancySource_(sheet,rows,slot,id);
  applications.getRange(index+2,5).setValue('已核准');
  SpreadsheetApp.flush();
  return {shiftId:id};
}

function masaluCloseVacancySource_(sheet,rows,slot,id) {
  const source = slot.row;
  const number = rows.indexOf(source)+2;
  const closeMarker = slot.leave?'[請假已替補:'+slot.leave.serial+']':'[管理已略過][替補ID:'+id+']';
  sheet.getRange(number,17).setValue(String(source[16] || '')+closeMarker);
  if (slot.leave) sheet.getRange(number,8).setValue('已請假');
}
