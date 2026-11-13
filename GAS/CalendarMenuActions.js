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
  return {ok:true,code:action,message:selected.length?title+'\n'+selected.map(entry=>entry.text).join('\n'):empty,
    ...(items.length?{quickReply:{items}}:{})};
}

function masaluCalendarMenuAction_(input) {
  const {lineUid,choice,value,text,eventId} = input;
  const reply = (code,message,items=[])=>({ok:true,code,message,...(items.length?{quickReply:{items}}:{})});
  const roster = employeeIndex_();
  const matches = Object.keys(roster).filter(id=>roster[id].uid===lineUid);
  if (!validLineUid_(lineUid) || matches.length!==1) return reply('unbound','請先綁定團員帳號。');
  const employeeId = matches[0];
  if (choice==='reply_contact') {
    const target = roster[value];
    if (!validLineUid_(target?.uid) || target.uid===lineUid ||
        !constantTimeEqual_(String(input.proof || ''),masaluContactReplyProof_(employeeId,lineUid,value,target.uid))) {
      return reply('unavailable','這則對話已無法回覆，請重新聯絡團員。');
    }
    return masaluSendMemberContact_(roster,employeeId,value,text,eventId,true);
  }
  const manager = ['A01','A02','A03'].includes(employeeId);
  const book = openBook_();
  const timeZone = book.getSpreadsheetTimeZone();
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
  const upcoming = rows.map(row=>({date:masaluDate_(row[3],timeZone),shift:String(row[6] || '').trim(),
    staffId:String(row[2] || '').trim(),status:String(row[7] || '').trim(),name:String(row[12] || '')}))
    .filter(row=>row.date>=today && row.shift && row.shift!=='休假' && (row.staffId || row.name))
    .sort((a,b)=>a.date.localeCompare(b.date)||a.shift.localeCompare(b.shift));

  if (choice==='replies') {
    const entries = upcoming
      .filter(row=>(!row.status || row.status==='待回覆') && (manager || row.staffId===employeeId))
      .map(row=>({text:[row.date,row.shift,manager?(roster[row.staffId]?.name || row.name || row.staffId):'']
        .filter(Boolean).join('｜')}));
    return masaluCalendarActionPage_(entries,value,'replies','回覆確認｜尚未回覆','目前沒有尚未回覆的排班。');
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
    return masaluCalendarActionPage_(entries,value,'contact','聯絡團員｜請選擇收件人','目前沒有可聯絡的團員。');
  }
  if (!contacts.includes(value)) return reply('unavailable','目前無法聯絡這位團員。請重新選擇。');
  if (choice==='contact_target') {
    const prefix = '團員留言 '+value+' ';
    return reply('contact','傳給 '+(roster[value].name || value)+'\n點「輸入訊息」，也可手動輸入「'+prefix+'你的訊息」。',[
      masaluCalendarActionButton_('輸入訊息','keyboard','',{inputOption:'openKeyboard',fillInText:prefix})
    ]);
  }
  if (choice!=='send') return reply('invalid','請從「聯絡團員」重新選擇收件人。');
  return masaluSendMemberContact_(roster,employeeId,value,text,eventId,false);
}

function masaluContactReplyProof_(recipientId,recipientUid,senderId,senderUid) {
  const secret = PropertiesService.getScriptProperties().getProperty('GAS_SHARED_SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(
    JSON.stringify(['member-contact-reply',recipientId,recipientUid,senderId,senderUid]),secret
  )).replace(/=+$/,'');
}

function masaluSendMemberContact_(roster,senderId,recipientId,text,eventId,isReply) {
  const message = String(text || '').trim();
  if (!message || message.length>1500 || !eventId) {
    return {ok:true,code:'invalid',message:'請輸入 1 至 1500 字的訊息。'};
  }
  const sender = roster[senderId], recipient = roster[recipientId];
  const title = '💬 '+(sender.name || senderId)+'｜'+(isReply?'回覆':'團員對話');
  const proof = masaluContactReplyProof_(recipientId,recipient.uid,senderId,sender.uid);
  const action = masaluCalendarActionButton_('回覆','keyboard','',{
    inputOption:'openKeyboard',fillInText:'團員回覆 '+senderId+' '+proof+' '
  }).action;
  const card = {type:'flex',altText:title.slice(0,400),contents:{type:'bubble',
    body:{type:'box',layout:'vertical',spacing:'md',contents:[
      {type:'text',text:title,weight:'bold',wrap:true},
      {type:'text',text:message,wrap:true}
    ]},
    footer:{type:'box',layout:'vertical',contents:[{type:'button',style:'secondary',height:'sm',action}]}
  }};
  const sent = masaluPushWorkflowMessage_(recipient.uid,card,'menu-contact:'+eventId);
  return sent ? {ok:true,code:'sent',message:(isReply?'回覆':'訊息')+'已傳給 '+(recipient.name || recipientId)+'。'} :
    {ok:true,code:'failed',message:'訊息尚未送達，請稍後重試。'};
}
