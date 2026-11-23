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
