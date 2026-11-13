/** Run from the bound spreadsheet to publish the calendar menu for the next two weeks. */
function masaluPublishCalendarRichMenu() {
  const page = masaluUpdateCalendarRichMenu_(true);
  const result = page.failures ? '，有 '+page.failures+' 個個人選單待重試。' : '。';
  SpreadsheetApp.getUi().alert(page.start+' 至 '+page.end+' 排班月曆圖文選單已發布'+result);
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
  const {properties, token, assets, page} = settings;
  if (!page) {
    if (manual) throw new Error('本週月曆圖片正在更新，稍後會自動重試。');
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (manual || properties.getProperty('MASALU_RICH_MENU_ALIAS') !== page.alias) {
      const id = masaluPrepareCalendarRichMenu_(page,assets,token);
      masaluCalendarRichMenuRequest_('user/all/richmenu/'+id,token,'post');
      properties.setProperty('MASALU_RICH_MENU_ALIAS',page.alias);
    }
    page.failures = masaluSyncPersonalMenus_(page,assets,token,properties);
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
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('缺少 LINE_CHANNEL_ACCESS_TOKEN');
  const assets = 'https://raw.githubusercontent.com/EnddenMi128/MaSaLu/main/rich-menu/';
  const manifest = UrlFetchApp.fetch(assets+'menu.json',{muteHttpExceptions:true});
  if (manifest.getResponseCode() !== 200) throw new Error('兩週月曆圖文選單尚未發布');
  const page = JSON.parse(manifest.getContentText());
  if (!page?.menu || !page.alias || !page.image || !page.personal) throw new Error('月曆選單資料不完整');
  return {start, properties, token, assets, page:page.start === start ? page : null};
}

function masaluRefreshOwnCalendarMenu_(uid) {
  if (!validLineUid_(uid)) return {ok:true,code:'invalid',message:'無法識別 LINE 帳號。'};
  const {properties, token, assets, page} = masaluCalendarMenuSettings_();
  if (!page) return {ok:true,code:'menu_waiting',message:'本週月曆圖片正在更新，請稍後再試。'};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const {status, entries, failureStage} = masaluSyncPersonalMenus_(page,assets,token,properties,uid);
    const messages = {
      unbound:'請先綁定團員帳號，再更新個人班表。',
      empty:'本週與下週沒有你的班次。',
      failed:'圖文選單暫時無法套用（'+failureStage+'）。',
      linked:'已套用個人班表（'+page.start+' 至 '+page.end+'），請返回 LINE 聊天列表再開啟聊天室查看。'
    };
    const list = entries.map(({date,groups,confirmed})=>[
      date.slice(5).replace('-','/')+'｜'+(confirmed?'已確認':'待回覆'),
      ...groups.map(({location,people})=>location+'｜'+people.map(person=>person.name).join('、'))
    ].join('\n'));
    const details = list.join('\n\n');
    const text = details.length > 1500 ? details.slice(0,1500)+'…\n開啟排班月曆查看完整名單。' : details;
    return {ok:true,code:status,message:list.length ? ['個人班表',text,'',messages[status]].join('\n') : messages[status]};
  } finally {
    lock.releaseLock();
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
    if (status !== allowedStatus) throw new Error('LINE 圖文選單設定失敗，HTTP '+status+'（'+path+'）');
  }
  return response;
}
