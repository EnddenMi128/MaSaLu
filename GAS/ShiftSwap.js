const MASALU_SWAP_INVITES = 'MaSaLu調班邀請';
const MASALU_SWAP_HEADERS = ['邀請ID','建立時間','原排班ID','原排班版次','原員工編號','原LINE_UID',
  '對方排班ID','對方排班版次','對方員工編號','對方LINE_UID','狀態','原日期','原班別','對方日期','對方班別'];

function masaluSwapShiftText_(row) {
  const times=[row[4],row[5]].map(value=>value instanceof Date
    ? Utilities.formatDate(value,openBook_().getSpreadsheetTimeZone(),'HH:mm')
    : String(value || '').trim());
  return String(row[6] || '')+' ・ '+(times.every(Boolean)?times.join('–'):'時段待確認');
}

function masaluSwapAction_(shift, choice, label, params={}) {
  const action=masaluFollowupAction_(shift,choice,label.slice(0,20));
  action.data+=Object.entries(params).map(([key,value])=>'&'+key+'='+encodeURIComponent(value)).join('');
  return {type:'action',action};
}

function masaluSwapDates_(input, shift, candidates, message='') {
  const dates=[...new Set(candidates.map(row=>masaluDate_(row[3])))].sort();
  if (!dates.length) return {ok:true,code:'swap_no_candidates',
    message:'目前沒有可調班的班次',
    quickReply:{items:[masaluSwapAction_(shift,'swap','重新查詢'),masaluSwapAction_(shift,'contact','聯絡團長')]}};
  const page=Math.max(0,Math.min(Number(input.page)||0,Math.floor((dates.length-1)/10)));
  const items=dates.slice(page*10,page*10+10).map(date=>masaluSwapAction_(shift,'swap_date',date,{date}));
  if (page>0) items.push(masaluSwapAction_(shift,'swap','上一頁',{page:page-1}));
  if ((page+1)*10<dates.length) items.push(masaluSwapAction_(shift,'swap','下一頁',{page:page+1}));
  return {ok:true,code:'swap_date',message:(message?message+'\n':'')+
    '請選擇日期'+
    (dates.length>10?'\n第 '+(page+1)+'／'+Math.ceil(dates.length/10)+' 頁':''),quickReply:{items}};
}

function masaluSwapSelection_(input, shift, roster, uid, sheet) {
  const today=masaluDate_(new Date());
  if (!shift.date || !shift.shift || shift.date<today)
    return {ok:true,code:'swap_expired',message:'班次已變更或過期，請重新選擇'};
  const last=sheet.getLastRow();
  const rows=last<2?[]:sheet.getRange(2,1,last-1,SHIFT_APP.SHIFT_HEADERS.length).getValues();
  const available=rows.filter(row=>{
    const recipient=roster[String(row[2]||'').trim()]?.uid;
    const date=masaluDate_(row[3]);
    return row[1] && row[1]!==shift.id && row[6] && row[7]==='同意' && date && date>=today &&
      validLineUid_(recipient) && recipient!==uid;
  });
  if (input.choice==='swap') return masaluSwapDates_(input,shift,available);
  const date=masaluDate_(input.date);
  if (!date || date<today) return masaluSwapDates_({},shift,available,'請選擇今天或之後的日期。');
  const candidates=available.filter(row=>masaluDate_(row[3])===date);
  const selected=candidates.find(row=>row[1]===input.target);
  if (input.choice==='swap_target') {
    if (!selected || Number(selected[11])!==input.targetVersion)
      return {...masaluSwapDates_({},shift,available,'對方的排班已變更，請重新選擇。'),code:'swap_expired'};
    return masaluSendSwapInvite_(shift,uid,selected,roster);
  }

  const byShift=input.choice==='swap_date';
  if (!candidates.length || (!byShift && !selected))
    return masaluSwapDates_({},shift,available,date+' 無可調班對象');
  const options=byShift ? [...new Map(candidates.map(row=>[masaluSwapShiftText_(row),row])).values()]
    : candidates.filter(row=>masaluSwapShiftText_(row)===masaluSwapShiftText_(selected));
  const page=Math.max(0,Math.min(Number(input.page)||0,Math.floor((options.length-1)/10)));
  const items=options.slice(page*10,page*10+10).map(row=>{
    if (byShift) return masaluSwapAction_(shift,'swap_shift',masaluSwapShiftText_(row),{date,target:row[1]});
    const staffId=String(row[2]).trim();
    return masaluSwapAction_(shift,'swap_target',roster[staffId]?.name || staffId,
      {date,target:row[1],tv:row[11]});
  });
  const nav=next=>masaluSwapAction_(shift,input.choice,next>page?'下一頁':'上一頁',
    {date,page:next,...(byShift?{}:{target:input.target})});
  if (page>0) items.push(nav(page-1));
  if ((page+1)*10<options.length) items.push(nav(page+1));
  items.push(byShift?masaluSwapAction_(shift,'swap','重選日期'):masaluSwapAction_(shift,'swap_date','重選班別',{date}));
  return {ok:true,code:byShift?'swap_shift':'swap_target',message:byShift?'請選擇班別\n'+date
    :'請選擇調班對象\n'+date+' ・ '+masaluSwapShiftText_(selected),quickReply:{items}};
}

function masaluSendSwapInvite_(shift, uid, target, roster) {
  const staffId=String(target[2]).trim();
  const recipient=roster[staffId];
  const book=openBook_();
  const sheet=book.getSheetByName(MASALU_SWAP_INVITES)||ensureSheet_(book,MASALU_SWAP_INVITES,MASALU_SWAP_HEADERS,false);
  const {rows}=masaluTable_(MASALU_SWAP_INVITES,MASALU_SWAP_HEADERS,book);
  const targetDate=masaluDate_(target[3]);
  const existing=rows.find(row=>row[2]===shift.id && Number(row[3])===shift.version && row[5]===uid &&
    row[6]===target[1] && Number(row[7])===Number(target[11]) && row[9]===recipient.uid &&
    masaluDate_(row[11])===shift.date && row[12]===shift.shift && masaluDate_(row[13])===targetDate && row[14]===target[6]);
  if (existing && existing[10]!=='待對方回覆')
    return {ok:true,code:'swap_already',message:'對方已回覆這筆調班邀請，狀態為「'+existing[10]+'」。'};
  const inviteId=existing ? existing[0] : Utilities.getUuid();
  if (!existing) {
    sheet.appendRow([inviteId,new Date(),shift.id,shift.version,shift.employeeId,uid,
      target[1],Number(target[11]),staffId,recipient.uid,'待對方回覆',shift.date,shift.shift,targetDate,target[6]]);
    SpreadsheetApp.flush();
  }
  const buttons=[['同意調班','agree'],['不同意調班','decline']].map(([label,decision])=>({
    type:'button',height:'sm',style:'secondary',action:{
      type:'postback',label,displayText:label,data:'action=swap_response&id='+inviteId+'&decision='+decision
    }
  }));
  const text=['🔄 調班詢問',shift.name+' 想和你調班。',
    '對方的班｜'+shift.date+' ・ '+(shift.shiftLabel || shift.shift),'你的班｜'+targetDate+' ・ '+masaluSwapShiftText_(target),
    '同意後會交由團長審核。'].join('\n');
  const card={type:'flex',altText:'調班詢問｜'+shift.name,contents:{type:'bubble',
    body:{type:'box',layout:'vertical',contents:[{type:'text',text,wrap:true}]},
    footer:{type:'box',layout:'vertical',spacing:'sm',contents:buttons}
  }};
  const sent=masaluPushWorkflowMessage_(recipient.uid,card,'swap-invite:'+inviteId);
  const details=targetDate+' ・ '+masaluSwapShiftText_(target)+'\n'+(recipient.name||staffId);
  return sent ? {ok:true,code:'swap_invited',message:'調班邀請已送出\n'+details}
    : {ok:true,code:'swap_not_sent',message:'調班邀請未送出，請重試\n'+details};
}

function masaluSwapResponse_(input) {
  const {id,decision,lineUid}=input;
  if (!validLineUid_(lineUid) || !/^[a-f0-9-]{36}$/i.test(id||'') || !['agree','decline'].includes(decision))
    return {ok:true,code:'invalid',message:'調班邀請格式錯誤。'};
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const book=openBook_();
    if (!book.getSheetByName(MASALU_SWAP_INVITES)) return {ok:true,code:'swap_missing',message:'找不到這筆調班邀請。'};
    const {sheet,rows}=masaluTable_(MASALU_SWAP_INVITES,MASALU_SWAP_HEADERS,book);
    const index=rows.findIndex(row=>row[0]===id);
    if (index<0) return {ok:true,code:'swap_missing',message:'找不到這筆調班邀請。'};
    const [, ,sourceId,sourceVersion,sourceStaff,sourceUid,targetId,targetVersion,targetStaff,targetUid,status,
      sourceDate,sourceShift,targetDate,targetShift]=rows[index];
    const roster=employeeIndex_();
    if (lineUid!==targetUid || roster[targetStaff]?.uid!==lineUid)
      return {ok:true,code:'wrong_user',message:'僅限受邀的本人回覆調班。'};
    const next=decision==='agree'?'待團長審核':'對方不同意';
    if (status!=='待對方回覆' && status!==next)
      return {ok:true,code:'swap_already',message:'這筆調班邀請已回覆。'};
    const shifts=shiftSheet_();
    const getRow=(shiftId,version,staff)=>{
      const number=findShiftRow_(shifts,shiftId);
      if (!number) return null;
      const row=shifts.getRange(number,1,1,SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
      return row[7]==='同意' && Number(row[11])===Number(version) && String(row[2]||'').trim()===String(staff) ? row : null;
    };
    const source=getRow(sourceId,sourceVersion,sourceStaff);
    const target=getRow(targetId,targetVersion,targetStaff);
    const today=masaluDate_(new Date());
    if (!source || !target || (sourceStaff && roster[sourceStaff]?.uid!==sourceUid) ||
        masaluDate_(source[3])!==masaluDate_(sourceDate) || source[6]!==sourceShift ||
        masaluDate_(target[3])!==masaluDate_(targetDate) || target[6]!==targetShift ||
        masaluDate_(source[3])<today || masaluDate_(target[3])<today)
      return {ok:true,code:'swap_expired',message:'排班已變更或過期，請重新發起調班。'};

    // Persist consent before handing off. A repeat of the same decision can finish a failed handoff.
    if (status!==next) {
      sheet.getRange(index+2,11).setValue(next);
      SpreadsheetApp.flush();
    }
    const other=roster[targetStaff]?.name || targetStaff;
    let managerPending=false;
    if (decision==='agree') {
      const shift={id:sourceId,version:sourceVersion,employeeId:sourceStaff,date:masaluDate_(source[3]),
        shift:source[6],name:roster[sourceStaff]?.name || source[12] || '臨時人員'};
      const details='對方已同意｜'+other+' ・ '+masaluDate_(target[3])+' ・ '+masaluSwapShiftText_(target)+
        '\n原班時段｜'+masaluSwapShiftText_(source);
      const result=masaluSubmitFollowup_(shift,'swap','',details,roster,id);
      managerPending=result.code==='request_saved';
    }
    const text=decision==='agree' ? other+' 已同意調班。案件已登記，待團長審核。'
      : other+' 不同意這次調班。';
    const sourceDetails='原本的班｜'+masaluDate_(source[3])+' ・ '+masaluSwapShiftText_(source);
    const targetDetails='想換的班｜'+masaluDate_(target[3])+' ・ '+masaluSwapShiftText_(target);
    masaluPushWorkflowMessage_(sourceUid,{type:'text',text:[text,sourceDetails,targetDetails].join('\n')},'swap-result:'+id);
    const answer=decision==='decline'?'已回覆不同意調班。'
      : '已同意並立案，待團長審核。'+(managerPending?'部分團長通知尚未送達，可再點「同意調班」重試。':'');
    return {ok:true,code:'swap_answered',message:[answer,'你的班｜'+masaluDate_(target[3])+' ・ '+masaluSwapShiftText_(target),
      '對方的班｜'+masaluDate_(source[3])+' ・ '+masaluSwapShiftText_(source)].join('\n')};
  } finally { lock.releaseLock(); }
}
