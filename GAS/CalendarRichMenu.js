/** Run from the bound spreadsheet to publish the calendar menu for the next two weeks. */
function masaluPublishCalendarRichMenu() {
  const page = masaluUpdateCalendarRichMenu_(true);
  const result = page.failures ? '，有 '+page.failures+' 個個人選單待重試。' : '。';
  SpreadsheetApp.getUi().alert(page.start+' 至 '+page.end+' 排班月曆圖文選單已發布'+result);
}

function masaluEnableWeeklyCalendarRichMenu() {
  const handler = 'masaluRefreshWeeklyCalendarRichMenu';
  if (!ScriptApp.getProjectTriggers().some(trigger=>trigger.getHandlerFunction()===handler)) {
    ScriptApp.newTrigger(handler).timeBased().everyMinutes(15).create();
  }
  masaluPublishCalendarRichMenu();
}

function masaluRefreshWeeklyCalendarRichMenu() {
  masaluUpdateCalendarRichMenu_(false);
}

function masaluUpdateCalendarRichMenu_(manual) {
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
  if (page.start !== start) {
    if (manual) throw new Error('本週月曆圖片正在更新，稍後會自動重試。');
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (manual || properties.getProperty('MASALU_RICH_MENU_WEEK') !== start) {
      const id = masaluPrepareCalendarRichMenu_(page,assets,token);
      masaluCalendarRichMenuRequest_('user/all/richmenu/'+id,token,'post');
      properties.setProperty('MASALU_RICH_MENU_WEEK',start);
    }
    page.failures = masaluSyncPersonalMenus_(page,assets,token,properties);
  } finally {
    lock.releaseLock();
  }
  return page;
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
