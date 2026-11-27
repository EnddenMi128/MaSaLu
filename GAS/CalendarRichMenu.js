/** Run from the bound spreadsheet to publish the calendar menu for the next two weeks. */
function masaluPublishCalendarRichMenu() {
  const page = masaluUpdateCalendarRichMenu_(true);
  const result = page.failures ? '，有 '+page.failures+' 個個人選單待重試。' : '。';
  show_(page.start+' 至 '+page.end+' 排班月曆圖文選單已發布'+result);
}

function masaluEnableWeeklyCalendarRichMenu() {
  masaluEnsureWeeklyMenuTrigger_();
  masaluPublishCalendarRichMenu();
}

function masaluEnsureWeeklyMenuTrigger_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const handler = 'masaluRefreshWeeklyCalendarRichMenu';
    if (!ScriptApp.getProjectTriggers().some(trigger=>trigger.getHandlerFunction()===handler)) {
      ScriptApp.newTrigger(handler).timeBased().everyMinutes(15).create();
    }
  } finally {
    lock.releaseLock();
  }
}

function masaluEnableWeeklyMenuFromLine_(uid) {
  const roster = employeeIndex_();
  if (!validLineUid_(uid) || !['A01','A02','A03'].some(id=>roster[id]?.uid===uid)) {
    return {ok:true,code:'unauthorized',message:'只有已綁定的 A01～A03 團長可以啟用每週更新。'};
  }
  try {
    masaluEnsureWeeklyMenuTrigger_();
  } catch (error) {
    console.error('無法建立 LINE 月曆定時觸發器');
    return {ok:true,code:'menu_trigger_failed',message:'每週更新尚未啟用。請由 Apps Script 專案擁有者執行 masaluEnableWeeklyCalendarRichMenu 並完成授權。'};
  }
  try {
    const page = masaluUpdateCalendarRichMenu_(true);
    const details = page.failures ? '，有 '+page.failures+' 個個人選單將自動重試' : '';
    return {ok:true,code:'menu_enabled',message:'已啟用每週更新，並發布 '+page.start+' 至 '+page.end+' 的月曆'+details+'。'};
  } catch (error) {
    console.error('已啟用每週更新，本次 LINE 月曆發布失敗');
    return {ok:true,code:'menu_enabled_pending',message:'已啟用每週更新。本次發布尚未完成，系統會自動重試。'};
  }
}

function masaluRefreshWeeklyCalendarRichMenu() {
  masaluUpdateCalendarRichMenu_(false);
}

function masaluUpdateCalendarRichMenu_(manual) {
  const settings = masaluCalendarMenuSettings_();
  const {properties, channels, assets, page} = settings;
  if (!page) {
    if (manual) throw new Error('本週月曆圖片正在更新，稍後會自動重試。');
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    page.failures = 0;
    for (const channel of channels) {
      masaluEnsureDefaultCalendarMenu_(page,assets,channel.token);
      page.failures += masaluSyncPersonalMenus_(page,assets,channel.token,properties,
        undefined,channel.prefix,channel.uids);
    }
  } finally {
    lock.releaseLock();
  }
  return page;
}

function masaluCalendarMenuSettings_() {
  const today = Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd');
  const sunday = new Date(today+'T00:00:00Z');
  sunday.setUTCDate(sunday.getUTCDate()-sunday.getUTCDay());
  const start = sunday.toISOString().slice(0,10);
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('缺少 LINE_CHANNEL_ACCESS_TOKEN');
  const channels = [{name:'primary',token,prefix:'MASALU_PERSONAL_MENU_'}];
  if (properties.getProperty('MASALU_BACKUP_OA_ENABLED') === 'true') {
    const backup = properties.getProperty('MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN');
    if (!backup) throw new Error('缺少 MASALU_BACKUP_LINE_CHANNEL_ACCESS_TOKEN');
    const uids = new Set((properties.getProperty('MASALU_BACKUP_FRIEND_UIDS') || '')
      .split(/[\s,]+/).filter(validLineUid_));
    channels.push({name:'backup',token:backup,prefix:'MASALU_BACKUP_PERSONAL_MENU_',uids});
  }
  const assets = 'https://raw.githubusercontent.com/EnddenMi128/MaSaLu/main/rich-menu/';
  const manifest = UrlFetchApp.fetch(assets+'menu.json?week='+start,{muteHttpExceptions:true});
  if (manifest.getResponseCode() !== 200) throw new Error('兩週月曆圖文選單尚未發布');
  const page = JSON.parse(manifest.getContentText());
  if (!page?.menu || !page.alias || !page.image || !page.personal) throw new Error('月曆選單資料不完整');
  return {start, properties, channels, assets, page:page.start === start ? page : null};
}

function masaluRefreshOwnCalendarMenu_(uid, channelName) {
  if (!validLineUid_(uid)) return {ok:true,code:'invalid',message:'無法識別 LINE 帳號。'};
  const {properties, channels, assets, page} = masaluCalendarMenuSettings_();
  if (!page) return {ok:true,code:'menu_waiting',message:'本週月曆圖片正在更新，請稍後再試。'};
  const channel = channels.find(item=>item.name===(channelName || 'primary'));
  if (!channel) return {ok:true,code:'menu_unavailable',message:'這個 LINE 帳號的月曆選單尚未啟用。'};
  if (channel.uids && !channel.uids.has(uid)) {
    // A user may have joined the second OA before its follow webhook was ready.
    // Verify the existing roster-linked UID with the SECOND channel itself.
    // Never trust the typed employee ID or merely the webhook's presence.
    if (!masaluVerifyAndRememberBackupFriend_(uid, channel.token))
      return {ok:true,code:'menu_unbound',message:'請先完成第二個 LINE 官方帳號的好友驗證。'};
    channel.uids.add(uid);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    masaluEnsureDefaultCalendarMenu_(page,assets,channel.token);
    const {status, entries, failureStage} = masaluSyncPersonalMenus_(page,assets,
      channel.token,properties,uid,channel.prefix,channel.uids);
    const messages = {
      unbound:'請先綁定團員帳號，再更新個人班表。',
      empty:'本週與下週沒有你的班次，已切回新版排班月曆。',
      failed:'圖文選單暫時無法套用（'+failureStage+'）。',
      linked:'已套用個人班表（'+page.start+' 至 '+page.end+'），請返回 LINE 聊天列表再開啟聊天室查看。'
    };
    const list = entries.map(({date,groups,confirmed})=>[
      date.slice(5).replace('-','/')+'｜'+(confirmed?'已確認':'待回覆'),
      ...groups.map(({location,people})=>location+'｜'+people.map(person=>person.name).join('、'))
    ].join('\n'));
    const details = list.join('\n\n');
    const text = details.length > 1500 ? details.slice(0,1500)+'…\n開啟排班月曆查看完整名單。' : details;
    // The calendar overview is not an actionable shift notification. Add real,
    // versioned selection postbacks, followed by a separate confirmation step.
    const pendingButtons = list.length ? masaluPendingCalendarReplyButtons_(uid,page) : [];
    const quickReplyItems = [...pendingButtons,
      ...(list.length ? [masaluCalendarActionButton_('我要請假','leave_pick',0)] : [])];
    const pendingHint = pendingButtons.length ? '\n\n請先選擇日期及地點，再選擇「同意」或「不同意」。' : '';
    return {ok:true,code:status,message:list.length ? ['個人班表',text,'',messages[status]+pendingHint].join('\n') : messages[status],
      ...(quickReplyItems.length ? {quickReply:{items:quickReplyItems}} : {})};
  } finally {
    lock.releaseLock();
  }
}

/** Personal calendar: first select an assignment, then confirm a decision. */
function masaluPendingCalendarRows_(uid,page) {
  const roster = employeeIndex_();
  const matches = Object.keys(roster).filter(id => roster[id].uid === uid);
  if (matches.length !== 1) return [];
  const employeeId = matches[0];
  const sheet = shiftSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2,1,lastRow-1,SHIFT_APP.SHIFT_HEADERS.length).getValues()
    .map(row=>({row,date:masaluDate_(row[3])}))
    .filter(({row,date})=>String(row[2]||'').trim()===employeeId &&
      date>=page.start && date<=page.end &&
      /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(String(row[1]||'')) &&
      Number.isSafeInteger(Number(row[11])) && Number(row[11])>0 &&
      ['', '待回覆'].includes(String(row[7]||'').trim()) &&
      String(row[6]||'').trim() && String(row[6]).trim()!=='休假')
    .sort((a,b)=>a.date.localeCompare(b.date)||String(a.row[6]).localeCompare(String(b.row[6])));
}

function masaluPendingCalendarReplyButtons_(uid,page,index) {
  const waiting=masaluPendingCalendarRows_(uid,page);
  const limit=10;
  const n=Math.min(Math.max(0,Math.floor(Number(index)||0)),Math.max(0,Math.ceil(waiting.length/limit)-1));
  const items=waiting.slice(n*limit,(n+1)*limit).map(({row,date})=>{
    const label=date.slice(5).replace('-','/')+' '+String(row[6]).trim();
    return masaluCalendarActionButton_(label,'reply_select',String(row[1])+':'+Number(row[11]),
      {displayText:label.slice(0,100)});
  });
  if(n>0) items.push(masaluCalendarActionButton_('上一頁','reply_pick',n-1));
  if((n+1)*limit<waiting.length) items.push(masaluCalendarActionButton_('下一頁','reply_pick',n+1));
  return items;
}

function masaluReplyCalendarPage_(today) {
  const sunday=new Date(today+'T00:00:00Z');
  sunday.setUTCDate(sunday.getUTCDate()-sunday.getUTCDay());
  const end=new Date(sunday);
  end.setUTCDate(end.getUTCDate()+13);
  return {start:sunday.toISOString().slice(0,10),end:end.toISOString().slice(0,10)};
}


function masaluEnsureDefaultCalendarMenu_(page,assets,token) {
  const id = masaluPrepareCalendarRichMenu_(page,assets,token);
  if (!masaluCalendarMenuLinked_('user/all/richmenu',id,token)) {
    masaluCalendarRichMenuRequest_('user/all/richmenu/'+id,token,'post');
    if (!masaluCalendarMenuLinked_('user/all/richmenu',id,token))
      throw new Error('LINE 尚未回報新版預設月曆已套用');
  }
}

function masaluPrepareCalendarRichMenu_(page,assets,token) {
  const aliasPath = 'richmenu/alias/'+page.alias;
  const existing = masaluCalendarRichMenuRequest_(aliasPath,token,'get',null,404);
  if (existing.getResponseCode() === 200) return JSON.parse(existing.getContentText()).richMenuId;

  const image = UrlFetchApp.fetch(assets+page.image);
  const created = masaluCalendarRichMenuRequest_('richmenu',token,'post',page.menu);
  const id = JSON.parse(created.getContentText()).richMenuId;
  try {
    masaluUploadCalendarRichMenu_(id,token,page.image.endsWith('.jpg')?'image/jpeg':'image/png',image.getContent());
    masaluCalendarRichMenuRequest_('richmenu/alias',token,'post',{richMenuAliasId:page.alias,richMenuId:id});
    return id;
  } catch (error) {
    // Keep the menu if the alias request succeeded but its response was lost.
    const alias = masaluCalendarRichMenuRequest_(aliasPath,token,'get',null,404);
    if (alias.getResponseCode() === 404) masaluCalendarRichMenuRequest_('richmenu/'+id,token,'delete');
    throw error;
  }
}

function masaluUploadCalendarRichMenu_(id,token,contentType,bytes) {
  const uploaded = UrlFetchApp.fetch('https://api-data.line.me/v2/bot/richmenu/'+id+'/content',{
    method:'post',contentType,payload:Utilities.newBlob(bytes,contentType),
    headers:{Authorization:'Bearer '+token},muteHttpExceptions:true
  });
  if (uploaded.getResponseCode() !== 200) throw new Error('月曆圖片上傳失敗，HTTP '+uploaded.getResponseCode());
}

function masaluCalendarRichMenuRequest_(path,token,method,payload,allowedStatus) {
  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/'+path,{
    method,headers:{Authorization:'Bearer '+token},muteHttpExceptions:true,
    ...(payload ? {contentType:'application/json',payload:JSON.stringify(payload)} : {})
  });
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    const allowed = Array.isArray(allowedStatus) ? allowedStatus : [allowedStatus];
    if (!allowed.includes(status)) throw new Error('LINE 圖文選單設定失敗，HTTP '+status+'（'+path+'）');
  }
  return response;
}
