/** The two Messaging API channels belong to the same LINE provider and share user IDs. */
function masaluAccountLinkIdentity_(uid) {
  if (!validLineUid_(uid)) return null;
  const roster = employeeIndex_();
  const matches = Object.keys(roster).filter(id => roster[id].uid === uid);
  return matches.length === 1 ? {id:matches[0],name:roster[matches[0]].name} : null;
}

function masaluAccountLinkStart_(input) {
  const uid = String(input.lineUid || '');
  const person = masaluAccountLinkIdentity_(uid);
  if (!person) return {ok:true,code:'unbound',
    message:'尚未找到這個 LINE 身分的團員綁定。請先在任一帳號完成首次身分驗證，再按「帳號綁定」。'};
  PropertiesService.getScriptProperties().setProperty('MASALU_LINK_PENDING_'+uid,
    JSON.stringify({channel:input.channel === 'backup' ? 'backup' : 'primary',time:Date.now()}));
  return {ok:true,code:'await_employee_id',message:'請輸入你的工號，例如 A03。只需要工號，不用綁定碼。'};
}

function masaluAccountLinkEmployee_(input) {
  const uid = String(input.lineUid || '');
  const properties = PropertiesService.getScriptProperties();
  const key = 'MASALU_LINK_PENDING_'+uid;
  let pending;
  try { pending = JSON.parse(properties.getProperty(key) || 'null'); } catch (_) {}
  if (!pending || Date.now() - pending.time > 10*60*1000 || pending.channel !== input.channel) {
    properties.deleteProperty(key);
    return {ok:true,code:'link_expired',message:'請先按圖文選單的「帳號綁定」，再輸入工號 Axx。'};
  }
  const person = masaluAccountLinkIdentity_(uid);
  if (!person || person.id !== String(input.employeeId || '').toUpperCase())
    return {ok:true,code:'employee_mismatch',message:'工號與目前已綁定的 LINE 身分不符，請確認後重新輸入。'};
  properties.deleteProperty(key);
  if (input.channel === 'backup') masaluRememberBackupFriend_(uid);
  const other = input.channel === 'backup' ? 'primary' : 'backup';
  const token = properties.getProperty(other === 'backup' ? 'MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN' : 'LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return {ok:true,code:'channel_unavailable',message:'另一個官方帳號尚未設定完成，請稍後再試。'};
  const profile = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/'+encodeURIComponent(uid),{
    headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
  if (profile.getResponseCode() === 200) {
    if (other === 'backup') masaluRememberBackupFriend_(uid);
    return {ok:true,code:'already_friends',message:person.id+' 已與兩個官方帳號完成對應，之後無需再操作。'};
  }
  const info = UrlFetchApp.fetch('https://api.line.me/v2/bot/info',{
    headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
  if (info.getResponseCode() !== 200) return {ok:true,code:'friend_link_unavailable',
    message:'已核對工號，但暫時無法取得另一個官方帳號的加好友入口。'};
  const basicId = JSON.parse(info.getContentText()).basicId;
  if (!/^@[A-Za-z0-9._-]{3,40}$/.test(basicId || '')) return {ok:true,code:'friend_link_unavailable',
    message:'已核對工號，但另一個官方帳號尚未提供加好友入口。'};
  const url = 'https://line.me/R/ti/p/'+basicId;
  return {ok:true,code:'add_friend',message:person.id+' 已核對。請點連結加入另一個官方帳號；加入後會自動完成，無需再輸入資料。\n'+url,
    quickReply:{items:[{type:'action',action:{type:'uri',label:'加入另一個帳號',uri:url}}]}};
}

function masaluAccountLinkFollow_(input) {
  const uid = String(input.lineUid || '');
  const person = masaluAccountLinkIdentity_(uid);
  if (!person) return {ok:true,code:'welcome_unbound',message:'歡迎使用 MaSaLu 排班小幫手！請先在任一帳號完成首次團員身分驗證。'};
  if (input.channel === 'backup') masaluRememberBackupFriend_(uid);
  return {ok:true,code:'linked',message:person.name+'（'+person.id+'）已自動對應這個官方帳號，無需再輸入工號。'};
}

function masaluRememberBackupFriend_(uid) {
  if (!validLineUid_(uid)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const names = new Set((properties.getProperty('MASALU_BACKUP_FRIEND_UIDS') || '')
      .split(/[\s,]+/).filter(validLineUid_));
    if (!names.has(uid)) {
      names.add(uid);
      properties.setProperty('MASALU_BACKUP_FRIEND_UIDS',[...names].join('\n'));
    }
  } finally { lock.releaseLock(); }
}

/**
 * Safely recover the second OA friend record when a user joined before
 * the backup follow webhook was installed. Never trust an employee ID
 * typed into chat as proof of identity.
 * Requires a unique already-bound LINE UID and a successful profile lookup
 * against the second OA's own Channel Access Token.
 */
function masaluVerifyAndRememberBackupFriend_(uid, backupToken) {
  if (!backupToken || !masaluAccountLinkIdentity_(uid)) return false;
  try {
    const response = UrlFetchApp.fetch(
      'https://api.line.me/v2/bot/profile/' + encodeURIComponent(uid), {
        method: 'get',
        headers: {Authorization: 'Bearer ' + backupToken},
        muteHttpExceptions: true
      });
    if (response.getResponseCode() !== 200) return false;
    const profile = JSON.parse(response.getContentText());
    if (profile.userId !== uid) return false;
    masaluRememberBackupFriend_(uid);
    return true;
  } catch (_) {
    console.error('第二 LINE 帳號好友驗證暫時失敗');
    return false;
  }
}

/** NFC employee IDs are enrollment instructions, not proof of employee identity.
 * Only an empty roster UID can be enrolled; an existing UID is never replaced.
 * LINE token verification always precedes roster access and mutations.
 */
function masaluNfcBind_(payload, accessToken) {
  const profile=masaluVerifiedLineProfile_(accessToken);
  if(!profile) return {ok:false,code:'unauthorized',error:'LINE 登入驗證失敗，請重新開啟'};
  const employeeId=String(payload && payload.employeeId || '').trim().toUpperCase();
  if(!/^A\d{2,3}$/.test(employeeId)) return {ok:false,error:'NFC 工號格式錯誤'};
  const uid=profile.userId,lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const table=masaluTable_(SHIFT_APP.EMPLOYEES,SHIFT_APP.EMP_HEADERS);
    const own=table.rows.filter(r=>String(r[3]||'').trim()===uid);
    if(own.length>1) return {ok:false,error:'此 LINE 有重複綁定，請聯絡管理員'};
    // An enrolled user always opens their own calendar, even on another NFC tag.
    if(own.length===1) {
      const ownId=String(own[0][0]).trim().toUpperCase();
      const notices=masaluNfcNotices_(ownId,uid,String(own[0][1]||''),false);
      return {ok:true,code:'already_bound',message:'已開啟個人月曆',employeeId:ownId,notifications:notices};
    }
    const matches=table.rows.map((row,i)=>({row,i})).filter(x=>String(x.row[0]).trim().toUpperCase()===employeeId);
    if(matches.length!==1) return {ok:false,error:'找不到唯一員工工號，請聯絡管理員'};
    const target=matches[0];
    if(String(target.row[3]||'').trim()) return {ok:false,code:'binding_conflict',error:'此工號已綁定其他 LINE，請聯絡管理員'};
    table.sheet.getRange(target.i+2,4,1,2).setValues([[uid,new Date()]]);
    SpreadsheetApp.flush();
    const notices=masaluNfcNotices_(employeeId,uid,String(target.row[1]||''),true);
    safeLog_('NFC綁定','',employeeId,'成功','LINE 身分驗證完成','');
    return {ok:true,code:'bound',employeeId,message:'綁定成功，已開啟個人月曆'+(notices.failed?'；部分 LINE 通知待重試':''),notifications:notices};
  } finally { lock.releaseLock(); }
}

/** Persist each recipient separately; deterministic LINE retry keys handle lost responses. */
function masaluNfcNotices_(id,uid,name,create) {
  const props=PropertiesService.getScriptProperties(),key='MASALU_NFC_NOTICE_'+id+'_'+uid;
  let record;
  try { record=JSON.parse(props.getProperty(key)||'null'); } catch(_) {}
  if(!record && !create) return {sent:0,failed:0};
  if(!record) {
    const roster=employeeIndex_();
    const targets=[{uid,text:name+'（'+id+'）LINE 綁定成功！個人月曆：https://liff.line.me/2010020470-VBPZRUQ9?view=personal'}];
    ['A01','A02','A03'].forEach(manager=>{
      if(roster[manager] && validLineUid_(roster[manager].uid)) targets.push({uid:roster[manager].uid,text:'NFC 綁定通知：'+name+'（'+id+'）已完成 LINE 綁定。'});
    });
    record={targets,done:[]};
    props.setProperty(key,JSON.stringify(record));
  }
  let sent=0,failed=0;
  record.targets.forEach((target,i)=>{
    if(record.done.includes(i)) return;
    try {
      const response=masaluPushWithQuotaFallback_(target.uid,[{type:'text',text:target.text}],'nfc:'+id+':'+uid+':'+i).response;
      const status=response.getResponseCode();
      if(!(status>=200 && status<300 || status===409 && masaluLineRetryAccepted_(response))) throw new Error('LINE notification failed');
      record.done.push(i);props.setProperty(key,JSON.stringify(record));sent++;
    } catch(_) { failed++; }
  });
  return {sent,failed};
}
