function masaluCalendarActionButton_(label, action, value, options) {
  return {type:'action',action:{type:'postback',label:label.slice(0,20),
    data:'action=calendar_menu&choice='+action+'&value='+encodeURIComponent(value),...options}};
}

function masaluCalendarActionPage_(entries, value, action, title, empty) {
  const limit = 10;
  const page = Math.min(Math.max(0,Number(value)||0),Math.max(0,Math.ceil(entries.length/limit)-1));
  const selected = entries.slice(page*limit,(page+1)*limit);
  const items = selected.map(entry=>entry.button).filter(Boolean);
  if (page>0) items.push(masaluCalendarActionButton_('上一頁',action,page-1));
  if ((page+1)*limit<entries.length) items.push(masaluCalendarActionButton_('下一頁',action,page+1));
  return {ok:true,code:action,message:selected.length?title+'\n\n'+selected.map(entry=>entry.text).join('\n'):empty,
    ...(items.length?{quickReply:{items}}:{})};
}

function masaluCalendarMenuAction_(input) {
  const {lineUid,choice,value,text,eventId} = input;
  const reply = (code,message,items=[])=>({ok:true,code,message,...(items.length?{quickReply:{items}}:{})});
  const roster = employeeIndex_();
  const matches = Object.keys(roster).filter(id=>roster[id].uid===lineUid);
  if (!validLineUid_(lineUid) || matches.length!==1) return reply('unbound','請先綁定團員帳號。');
  const employeeId = matches[0];
  const manager = ['A01','A02','A03'].includes(employeeId);
  const book = openBook_();
  const timeZone = book.getSpreadsheetTimeZone();
  if (choice==='contact_reply' || choice==='reply_send') {
    const cache = CacheService.getScriptCache();
    const thread = JSON.parse(cache.get('masalu:member-reply:'+lineUid+':'+value) || 'null');
    const target = thread?.id;
    if (!target || !validLineUid_(thread.uid) || roster[target]?.uid!==thread.uid || thread.uid===lineUid)
      return reply('expired','這則回覆已失效，請重新開啟聯絡團員。');
    if (choice==='contact_reply') {
      const prefix = '團員回覆 '+value+' ';
      return reply('contact_reply','回覆給 '+(roster[target].name || target),[
        masaluCalendarActionButton_('輸入回覆','keyboard','',{inputOption:'openKeyboard',fillInText:prefix})
      ]);
    }
    return masaluSendMemberMessage_(employeeId,target,roster,text,eventId,'回覆團員');
  }
  if (choice==='leave' || choice==='leave_detail') {
    const name = 'MaSaLu排班請求';
    const requests = book.getSheetByName(name) && book.getSheetByName(MASALU_SERIAL_SHEET)
      ? masaluTable_(name,MASALU_FOLLOWUP_HEADERS,book).rows : [];
    if (choice==='leave_detail') {
      if (!manager) return reply('unauthorized','只有管理員可以審核請假。');
      return masaluLeaveReview_({lineUid,serial:value,eventId});
    }
    const serials = requests.length ? new Map(masaluTable_(MASALU_SERIAL_SHEET,MASALU_SERIAL_HEADERS,book).rows
      .map(row=>[row[0],row[1]])) : new Map();
    const entries = requests.filter(row=>row[8]==='請假申請' && (manager
      ? ['待處理','待團長審核'].includes(row[11]) : String(row[4]).trim()===employeeId))
      .reverse().map(row=>({text:[masaluDate_(row[6],timeZone),row[7],manager?row[5]:row[9],row[11]].join('｜'),
        button:manager && serials.has(row[0])
          ? masaluCalendarActionButton_('審核 '+String(row[5] || '')+' '+serials.get(row[0]),'leave_detail',serials.get(row[0])) : null}));
    return masaluCalendarActionPage_(entries,value,'leave','請假審核',manager?'目前沒有待審核的請假。':'你目前沒有請假申請。');
  }
  if (choice==='keyboard') return reply('contact','請在輸入欄接著填寫要傳送的訊息。');

  const today = Utilities.formatDate(new Date(),timeZone,'yyyy-MM-dd');
  const sheet = shiftSheet_();
  const rows = sheet.getLastRow()<2 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,SHIFT_APP.SHIFT_HEADERS.length).getValues();
  if (choice==='leave_pick') {
    const offNames = new Set(['休假']);
    const config = book.getSheetByName(MASALU.SHIFTS);
    if (config) for (const row of config.getDataRange().getValues().slice(1)) {
      if (String(row[2]).toLowerCase()==='true') offNames.add(String(row[1]).trim());
    }
    const entries = rows.filter(row=>String(row[2]||'').trim()===employeeId && row[7]==='同意' &&
      /^[a-f0-9-]{36}$/i.test(String(row[1])) && Number.isSafeInteger(Number(row[11])) &&
      Number(row[11])>0 && row[6] && !offNames.has(String(row[6]).trim()) &&
      masaluDate_(row[3],timeZone)>=today).map(row=>{
        const date=masaluDate_(row[3],timeZone);
        const label=date+'｜'+row[6];
        const action=masaluFollowupAction_({id:row[1],version:Number(row[11])},'leave',label.slice(0,20));
        return {date,text:label,button:{type:'action',action}};
      }).sort((a,b)=>a.date.localeCompare(b.date));
    return masaluCalendarActionPage_(entries,value,'leave_pick','請選擇要請假的日期與班別','目前沒有可請假的已確認班次。');
  }
  const upcoming = rows.map(row=>({date:masaluDate_(row[3],timeZone),shift:String(row[6] || '').trim(),
    staffId:String(row[2] || '').trim(),status:String(row[7] || '').trim(),name:String(row[12] || '')}))
    .filter(row=>row.date>=today && row.shift && row.shift!=='休假' && (row.staffId || row.name))
    .sort((a,b)=>a.date.localeCompare(b.date)||a.shift.localeCompare(b.shift));

  if (choice==='replies') {
    const entries = upcoming
      .filter(row=>(!row.status || row.status==='待回覆') && (manager || row.staffId===employeeId))
      .map(row=>({text:[row.date,row.shift,manager?(roster[row.staffId]?.name || row.name || row.staffId):'']
        .filter(Boolean).join('｜')}));
    return masaluCalendarActionPage_(entries,value,'replies','尚未回覆','排班皆已回覆');
  }

  const ownConfirmed = new Set(upcoming.filter(row=>row.staffId===employeeId && row.status==='同意')
    .map(row=>row.date+'｜'+row.shift));
  const coworkers = new Set(upcoming.filter(row=>row.status==='同意' && ownConfirmed.has(row.date+'｜'+row.shift))
    .map(row=>row.staffId));
  const contacts = Object.keys(roster).filter(id=>id!==employeeId && validLineUid_(roster[id].uid) &&
    roster[id].uid!==lineUid && (manager || coworkers.has(id)));
  if (choice==='contact') {
    const entries = contacts.map(id=>({text:roster[id].name || id,
      button:masaluCalendarActionButton_(roster[id].name || id,'contact_target',id)}));
    return masaluCalendarActionPage_(entries,value,'contact','選擇收件人','目前沒有可聯絡的團員。');
  }
  if (!contacts.includes(value)) return reply('unavailable','目前無法聯絡這位團員。請重新選擇。');
  if (choice==='contact_target') {
    const prefix = '團員留言 '+value+' ';
    return reply('contact','傳給 '+(roster[value].name || value),[
      masaluCalendarActionButton_('輸入訊息','keyboard','',{inputOption:'openKeyboard',fillInText:prefix})
    ]);
  }
  if (choice!=='send') return reply('invalid','請重新開啟聯絡團員。');
  return masaluSendMemberMessage_(employeeId,value,roster,text,eventId,'聯絡團員');
}

function masaluSendMemberMessage_(senderId,targetId,roster,text,eventId,title) {
  if (!text || text.length>1500 || !eventId)
    return {ok:true,code:'invalid',message:'請輸入 1 至 1500 字的訊息。'};
  const target = roster[targetId];
  const token = masaluLineRetryKey_('member-reply:'+eventId+':'+target.uid);
  CacheService.getScriptCache().put('masalu:member-reply:'+target.uid+':'+token,
    JSON.stringify({id:senderId,uid:roster[senderId].uid}),21600);
  const sent = masaluPushWorkflowMessage_(target.uid,{
    type:'text',text:'💬 '+(roster[senderId].name || senderId)+'｜'+title+'\n\n'+text,
    quickReply:{items:[masaluCalendarActionButton_('回覆','contact_reply',token,
      {inputOption:'openKeyboard',fillInText:'團員回覆 '+token+' '})]}
  },'menu-contact:'+eventId);
  return {ok:true,code:sent?'sent':'failed',message:sent
    ? '訊息已傳給 '+(target.name || targetId)+'。' : '訊息尚未送達，請稍後重試。'};
}
