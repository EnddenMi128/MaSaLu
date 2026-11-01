/**
 * 團長排班 LINE 通知系統｜Google Apps Script（綁定於 Google 試算表）
 * 版本：1.0.0
 * LINE Webhook 必須先經由 Cloudflare Worker 驗證簽章，再轉交 doPost。
 * 請勿將 Access Token、Channel Secret 或 GAS_SHARED_SECRET 寫在此原始碼內。
 */

const SHIFT_APP = Object.freeze({
  EMPLOYEES: '員工名單',
  SHIFTS: '排班表',
  LOG: '操作紀錄',
  BATCH_LIMIT: 25,
  TEMP_DAYS: 7,
  EMP_HEADERS: ['員工編號', '姓名', '綁定碼', 'LINE_UID', '綁定時間'],
  SHIFT_HEADERS: [
    '發送', '排班ID', '員工編號', '排班日期', '開始時間', '結束時間', '班別',
    '回覆狀態', '回覆時間', '通知狀態', '發送時間', '通知版次',
    '臨時姓名', '臨時LINE_UID', '臨時有效期限', '最後錯誤', '備註'
  ],
  LOG_HEADERS: ['時間', '動作', '排班ID', '員工編號', '結果', '備註', '事件ID']
});

function onOpen() {
  let ui;
try {
  ui = SpreadsheetApp.getUi();
} catch (_) {
  console.log('請重新開啟綁定的試算表，以顯示排班選單。');
  return;
}
ui.createMenu('📨 團長排班通知')
    .addItem('① 初始化試算表（首次）', 'setupShiftSystem')
    .addItem('② 產生員工綁定碼', 'generateEmployeeBindCodes')
    .addSeparator()
    .addItem('檢查選取排班（不發送）', 'previewSelectedShift')
    .addItem('發送選取排班（指定 1 人）', 'sendSelectedShift')
    .addItem('發送已勾選排班（指定多人）', 'sendCheckedShifts')
    .addItem('重新發送選取排班', 'resendSelectedShift')
    .addSeparator()
    .addItem('建立臨時收件人排班（不發送）', 'createTemporaryShiftFromSelected')
    .addItem('清理過期臨時 UID', 'cleanupExpiredTempRecipients')
    .addItem('安裝每日臨時資料清理', 'installDailyCleanupTrigger')
    .addToUi();
}

/** 首次執行：僅建立缺少的表格，不覆蓋既有資料。 */
function setupShiftSystem() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('請從 Google 試算表的「擴充功能 → Apps Script」建立綁定專案。');
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', book.getId());

  ensureSheet_(book, SHIFT_APP.EMPLOYEES, SHIFT_APP.EMP_HEADERS, false);
  ensureSheet_(book, SHIFT_APP.SHIFTS, SHIFT_APP.SHIFT_HEADERS, true);
  ensureSheet_(book, SHIFT_APP.LOG, SHIFT_APP.LOG_HEADERS, false);
  const emp = book.getSheetByName(SHIFT_APP.EMPLOYEES);
  emp.getRange('A:A').setNumberFormat('@');
  emp.getRange('C:D').setNumberFormat('@');
  const shifts = book.getSheetByName(SHIFT_APP.SHIFTS);
  shifts.getRange('C:C').setNumberFormat('@');
  shifts.getRange('D:D').setNumberFormat('yyyy/MM/dd');
  shifts.getRange('E:F').setNumberFormat('HH:mm');
  shifts.getRange('I:I').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  shifts.getRange('K:K').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  shifts.getRange('O:O').setNumberFormat('yyyy/MM/dd HH:mm:ss');
  SpreadsheetApp.getUi().alert('初始化完成。請先填「員工名單」、產生綁定碼，再填「排班表」。');
}

function ensureSheet_(book, name, headers, addCheckboxes) {
  let sheet = book.getSheetByName(name);
  const isNew = !sheet;
  if (!sheet) sheet = book.insertSheet(name);
  const isEmpty = sheet.getLastRow() === 0;
  if (!isNew && !isEmpty) {
    const oldHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (oldHeaders.some((text, index) => text !== headers[index])) {
      throw new Error('「' + name + '」已存在但欄名不符。為避免改壞資料，請先手動確認表格。');
    }
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#F07B36').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1);
  if ((isNew || isEmpty) && addCheckboxes) {
    // 僅設定核取框驗證規則，不將 999 列預先寫入 FALSE，避免 getLastRow() 誤判。
    const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, 999, 1).setDataValidation(checkboxRule);
  }
  return sheet;
}

/** 產生一次性綁定碼；已綁定的員工不重新產生。 */
function generateEmployeeBindCodes() {
  const sheet = openBook_().getSheetByName(SHIFT_APP.EMPLOYEES);
  const last = sheet.getLastRow();
  if (last < 2) return show_('請先在「員工名單」填入員工編號與姓名。');
  const rows = sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
  const seenIds = new Set();
  const seenCodes = new Set(rows.map(row => String(row[2]).trim().toUpperCase()).filter(Boolean));
  let created = 0;
  rows.forEach((row, index) => {
    const id = String(row[0]).trim();
    const name = String(row[1]).trim();
    if (!id && !name) return;
    if (!id || !name) throw new Error('員工名單第 ' + (index + 2) + ' 列缺少編號或姓名。');
    if (seenIds.has(id)) throw new Error('員工編號重複：' + id);
    seenIds.add(id);
    if (row[3] || row[2]) return;
    let code;
    do { code = Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase(); }
    while (seenCodes.has(code));
    seenCodes.add(code);
    sheet.getRange(index + 2, 3).setValue(code);
    created++;
  });
  show_('已產生 ' + created + ' 組一次性綁定碼。請私下把每位員工自己的碼交給本人。');
}

/** 指定單人：選取排班表的一列，不會把同員工其他日期一起發送。 */
function sendSelectedShift() {
  const row = selectedShiftRow_();
  const result = sendShiftRow_(row, false);
  show_(result.message);
}

/** 重新發送會產生新版按鈕，舊通知按鈕隨即失效。 */
function resendSelectedShift() {
  const row = selectedShiftRow_();
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('重新發送', '將清除這筆排班的舊回覆、產生新版按鈕，確定嗎？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  show_(sendShiftRow_(row, true).message);
}

/** 指定多人：只送 A 欄勾選的列，每次最多 25 筆；已發送者會跳過。 */
function sendCheckedShifts() {
  const sheet = shiftSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return show_('排班表沒有資料。');
  const marked = sheet.getRange(2, 1, last - 1, 1).getValues();
  const rows = marked.map((x, i) => x[0] === true ? i + 2 : null).filter(Boolean);
  if (!rows.length) return show_('請先在排班表 A 欄勾選收件對象。');
  if (SpreadsheetApp.getUi().alert('確認發送', '本次將嘗試發送前 ' + Math.min(rows.length, SHIFT_APP.BATCH_LIMIT) + ' 筆勾選排班，確定嗎？', SpreadsheetApp.getUi().ButtonSet.YES_NO) !== SpreadsheetApp.getUi().Button.YES) return;
  const counts = {sent: 0, skipped: 0, failed: 0};
  rows.slice(0, SHIFT_APP.BATCH_LIMIT).forEach(row => {
    const result = sendShiftRow_(row, false);
    counts[result.code] = (counts[result.code] || 0) + 1;
    if (result.code === 'sent' || result.code === 'skipped') sheet.getRange(row, 1).setValue(false);
  });
  show_('完成：成功 ' + counts.sent + '、略過 ' + counts.skipped + '、失敗 ' + counts.failed + '。' + (rows.length > SHIFT_APP.BATCH_LIMIT ? '\n其餘勾選資料請再執行一次。' : ''));
}

/** 不發送訊息，只驗證所選排班和收件人。 */
function previewSelectedShift() {
  const rowNumber = selectedShiftRow_();
  const sheet = shiftSheet_();
  const values = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
  const display = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getDisplayValues()[0];
  try {
    const recipient = recipientForRow_(values, employeeIndex_());
    const info = shiftInfo_(display);
    show_('不會發送（預覽）\n收件人：' + recipient.name + '\nLINE UID：***' + recipient.uid.slice(-4) + '\n日期：' + info.date + '\n時間：' + info.time + '\n班別：' + info.shift + '\n通知狀態：' + (display[9] || '未發送'));
  } catch (err) {
    show_('檢查未通過：' + err.message);
  }
}

/**
 * 第三選項：依所選排班複製日期／時間／班別，建立「臨時」對象的新列。
 * 不寫入永久員工名單；LINE UID 僅保留在這筆臨時排班，回覆或到期後刪除。
 * 臨時對象可向 Bot 發送「我的UID」取得自己的 UID，再交給團長。
 */
function createTemporaryShiftFromSelected() {
  const sourceRow = selectedShiftRow_();
  const sheet = shiftSheet_();
  const source = sheet.getRange(sourceRow, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
  const ui = SpreadsheetApp.getUi();
  const namePrompt = ui.prompt('臨時收件人（不存入員工名單）', '請輸入這次的收件人姓名', ui.ButtonSet.OK_CANCEL);
  if (namePrompt.getSelectedButton() !== ui.Button.OK) return;
  const name = namePrompt.getResponseText().trim();
  if (!name || name.length > 40) return show_('姓名必填且最多 40 字。');
  const uidPrompt = ui.prompt('臨時 LINE UID', '請輸入對方傳給你的 U 開頭 UID（可先請對方傳「我的UID」給 Bot）', ui.ButtonSet.OK_CANCEL);
  if (uidPrompt.getSelectedButton() !== ui.Button.OK) return;
  const uid = uidPrompt.getResponseText().trim();
  if (!validLineUid_(uid)) return show_('LINE UID 格式不正確。');
  if (!source[3] || !source[6]) return show_('來源排班需填寫日期與班別。');
  const expires = new Date(Date.now() + SHIFT_APP.TEMP_DAYS * 86400000);
  const row = [true, Utilities.getUuid(), '', source[3], source[4], source[5], source[6], '待回覆', '', '未發送', '', 0, name, uid, expires, '', '臨時收件人（7 天內有效）'];
  sheet.appendRow(row);
  const targetRow = sheet.getLastRow();
  sheet.getRange(targetRow, 1).insertCheckboxes().setValue(true);
  show_('已新增第 ' + targetRow + ' 列臨時排班（尚未發送）。\n請檢查新列，再按「發送選取排班」或「發送已勾選排班」。\n這個 UID 不會進入「員工名單」。');
}

function sendShiftRow_(rowNumber, force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return {code: 'failed', message: '系統忙碌，請稍後再試。'};
  try {
    const sheet = shiftSheet_();
    const current = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    const display = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getDisplayValues()[0];
    if (typeof masaluCancelled_ === 'function' && masaluCancelled_(current)) return {code:'skipped',message:'此班次已取消，不可重新發送舊通知'};
    if (!force && current[9] === '已發送') return {code: 'skipped', message: '第 ' + rowNumber + ' 列已發送；如需再通知請使用「重新發送」。'};
    if (!force && ['同意', '不同意'].includes(String(current[7]))) return {code: 'skipped', message: '第 ' + rowNumber + ' 列已有回覆；如需重發請使用「重新發送」。'};
    const info = shiftInfo_(display);
    const recipient = recipientForRow_(current, employeeIndex_());
    const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token) throw new Error('尚未設定 Script Property：LINE_CHANNEL_ACCESS_TOKEN');
    let id = String(current[1] || '').trim();
    if (!id) { id = Utilities.getUuid(); sheet.getRange(rowNumber, 2).setValue(id); }
    if (!/^[a-f0-9-]{36}$/i.test(id)) throw new Error('排班ID 格式不正確：' + id);
    const version = (Number(current[11]) || 0) + 1;
    sheet.getRange(rowNumber, 8, 1, 5).setValues([['待回覆', '', '發送中', '', version]]);
    sheet.getRange(rowNumber, 16).clearContent();
    SpreadsheetApp.flush();
    const message = buildShiftFlex_(id, version, recipient.name, info);
    try {
      const http = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        contentType: 'application/json; charset=UTF-8',
        headers: {Authorization: 'Bearer ' + token},
        payload: JSON.stringify({to: recipient.uid, messages: [message]}),
        muteHttpExceptions: true
      });
      const status = http.getResponseCode();
      if (status < 200 || status >= 300) throw new Error('LINE HTTP ' + status + '：' + http.getContentText().slice(0, 250));
      sheet.getRange(rowNumber, 10, 1, 2).setValues([['已發送', new Date()]]);
      safeLog_('發送', id, String(current[2] || ''), '成功', recipient.name, '');
      return {code: 'sent', message: '第 ' + rowNumber + ' 列已成功向 LINE API 提交通知：' + recipient.name + '。（不代表對方已讀）'};
    } catch (err) {
      sheet.getRange(rowNumber, 10).setValue('發送失敗');
      sheet.getRange(rowNumber, 16).setValue(String(err.message).slice(0, 450));
      safeLog_('發送', id, String(current[2] || ''), '失敗', String(err.message).slice(0, 200), '');
      return {code: 'failed', message: '第 ' + rowNumber + ' 列發送失敗：' + err.message};
    }
  } catch (err) {
    return {code: 'failed', message: '第 ' + rowNumber + ' 列未發送：' + err.message};
  } finally {
    lock.releaseLock();
  }
}

function buildShiftFlex_(shiftId, version, name, info) {
  const action = (label, choice, color) => ({
    type: 'button', style: 'primary', color: color,
    action: {
      type: 'postback', label: label,
      data: 'action=shift_reply&id=' + encodeURIComponent(shiftId) + '&v=' + version + '&choice=' + choice,
      displayText: '我選擇：' + label
    }
  });
  return {
    type: 'flex', altText: '【排班通知】' + info.date + '，請確認是否同意',
    contents: {
      type: 'bubble',
      body: {type: 'box', layout: 'vertical', spacing: 'md', contents: [
        {type: 'text', text: '📅 排班確認通知', weight: 'bold', size: 'lg', color: '#D96B24'},
        {type: 'separator'},
        {type: 'text', text: '親愛的 ' + name.slice(0, 40) + ' 您好：', wrap: true, size: 'sm'},
        {type: 'text', text: '團長已完成排班，請確認以下安排。', wrap: true, size: 'sm'},
        {type: 'text', text: '日期：' + info.date, wrap: true, size: 'sm'},
        {type: 'text', text: '時間：' + info.time, wrap: true, size: 'sm'},
        {type: 'text', text: '班別：' + info.shift.slice(0, 100), wrap: true, size: 'sm'}
      ]},
      footer: {type: 'box', layout: 'horizontal', spacing: 'sm', contents: [
        action('同意', 'agree', '#249B59'), action('不同意', 'disagree', '#CF4545')
      ]}
    }
  };
}

function shiftInfo_(display) {
  const date = String(display[3] || '').trim();
  const start = String(display[4] || '').trim();
  const end = String(display[5] || '').trim();
  const shift = String(display[6] || '').trim();
  if (!date || !shift) throw new Error('排班日期與班別為必填。');
  if (start && !end || !start && end) throw new Error('開始、結束時間必須同時填寫，或同時留空。');
  return {date, time: start && end ? start + '–' + end : '依團長公告', shift};
}

function recipientForRow_(values, roster) {
  const employeeId = String(values[2] || '').trim();
  const tempName = String(values[12] || '').trim();
  const tempUid = String(values[13] || '').trim();
  if (employeeId && (tempName || tempUid)) throw new Error('不能同時填永久員工編號與臨時收件人。');
  if (employeeId) {
    const employee = roster[employeeId];
    if (!employee) throw new Error('員工編號 ' + employeeId + ' 不存在於員工名單。');
    if (!validLineUid_(employee.uid)) throw new Error('員工「' + employee.name + '」尚未綁定 LINE。');
    return {name: employee.name, uid: employee.uid, temporary: false};
  }
  if (!tempName || !validLineUid_(tempUid)) throw new Error('臨時對象必須有姓名與有效 LINE UID。');
  const expireAt = values[14] instanceof Date ? values[14].getTime() : new Date(values[14]).getTime();
  if (!Number.isFinite(expireAt) || expireAt <= Date.now()) throw new Error('臨時收件人已過期，請重新建立。');
  return {name: tempName, uid: tempUid, temporary: true};
}

function employeeIndex_() {
  const sheet = openBook_().getSheetByName(SHIFT_APP.EMPLOYEES);
  const last = sheet.getLastRow();
  const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
  const index = Object.create(null);
  rows.forEach((row, i) => {
    const id = String(row[0]).trim();
    if (!id) return;
    if (index[id]) throw new Error('員工名單的員工編號重複：' + id + '（第 ' + (i + 2) + ' 列）');
    index[id] = {name: String(row[1]).trim(), uid: String(row[3]).trim()};
  });
  return index;
}

/** Cloudflare Worker 經驗證後呼叫；一般公網請求沒有共享密鑰就不能修改。 */
function doPost(e) {
  let result;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('缺少 JSON Body');
    const input = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('GAS_SHARED_SECRET');
    if (!secret || !constantTimeEqual_(String(input.secret || ''), secret)) {
      result = {ok: false, code: 'unauthorized', message: '拒絕存取'};
    } else if (input.kind === 'shift_reply') {
      result = handleShiftReply_(input);
    } else if (input.kind === 'bind_or_message') {
      result = handleLineText_(input);
    } else if (input.kind === 'calendar_api') {
      // Worker has already verified GAS_SHARED_SECRET.  The calendar adapter
      // verifies LINE-bound A01-A03 (or legacy key), and expectedRevision.
      result = masaluApi(String(input.action || ''), input.payload || {}, input.adminKey, input.lineAccessToken);
      if (result.ok && ['saveDay','clearDay'].includes(input.action)) {
        try {
          result.cancellationNotifications = masaluNotifyCancelledForDate_(input.payload.date);
          result.notifications = input.action === 'saveDay' ? masaluNotifyUnsentForDate_(input.payload.date) : {sent:0,failed:0,skipped:0,details:[]};
        } catch (notifyError) {
          result.notifications = {sent:0,failed:1,skipped:0,
            details:[{status:'failed',message:'排班已儲存，但通知流程失敗：'+String(notifyError.message)}]};
        }
        // Sending changes notification status/version in the sheet. The
        // optimistic revision must describe the FINAL persisted state.
        const viewer=masaluViewer_(input.lineAccessToken || input.adminKey);
        const fresh = masaluBootstrap_(null,viewer || undefined);
        result.revision = fresh.revision;
        result.schedules = fresh.schedules.filter(row => row.date === input.payload.date);
      }
    } else {
      // Never acknowledge an unrecognized mutation as successful.
      result = {ok: false, code: 'unsupported', message: '不支援的操作'};
    }
  } catch (err) {
    console.error('Webhook 失敗：' + String(err.message).slice(0, 120));
    result = {ok: false, code: 'internal_error', message: '系統暫時無法處理，請稍後重試。'};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Send each newly saved, still-unsent assignment to its bound employee. */
function masaluNotifyUnsentForDate_(dateValue) {
  const date = masaluDate_(dateValue);
  if (!date) throw new Error('通知日期格式錯誤');
  const sheet = shiftSheet_();
  const last = sheet.getLastRow();
  const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, SHIFT_APP.SHIFT_HEADERS.length).getValues();
  const candidates = [];
  rows.forEach((row, index) => {
    if (masaluDate_(row[3]) !== date || !row[2] || !row[6]) return;
    const status = String(row[9] || '');
    const reply = String(row[7] || '');
    // Do not duplicate delivered or in-flight messages, or alter replied rows.
    if (!['', '未發送', '發送失敗'].includes(status) || ['同意', '不同意'].includes(reply)) return;
    candidates.push(index + 2);
  });
  const summary = {sent:0,failed:0,skipped:0,details:[]};
  candidates.slice(0, SHIFT_APP.BATCH_LIMIT).forEach(rowNumber => {
    const outcome = sendShiftRow_(rowNumber, false);
    if (outcome.code === 'sent') summary.sent++;
    else if (outcome.code === 'failed') summary.failed++;
    else summary.skipped++;
    summary.details.push({row:rowNumber,status:outcome.code,message:outcome.message});
  });
  if (candidates.length > SHIFT_APP.BATCH_LIMIT) {
    const remaining = candidates.length - SHIFT_APP.BATCH_LIMIT;
    summary.skipped += remaining;
    summary.details.push({status:'skipped',message:'其餘 '+remaining+' 筆超過單次通知上限，請分批處理'});
  }
  return summary;
}

/** Deliver cancellation only to the employee whose previously sent shift changed. */
function masaluNotifyCancelledForDate_(dateValue) {
  const date=masaluDate_(dateValue);
  if(!date) throw new Error('取消通知日期格式錯誤');
  const sheet=shiftSheet_();
  const last=sheet.getLastRow();
  const rows=last<2 ? [] : sheet.getRange(2,1,last-1,SHIFT_APP.SHIFT_HEADERS.length).getValues();
  const roster=employeeIndex_();
  const token=PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const summary={sent:0,failed:0,skipped:0,details:[]};
  rows.forEach((r,index)=>{
    if(masaluDate_(r[3])!==date || !['已取消','已取消通知失敗'].includes(String(r[9]||'')))return;
    if(summary.sent+summary.failed>=SHIFT_APP.BATCH_LIMIT){summary.skipped++;return;}
    const rowNumber=index+2;
    const staffId=String(r[2]||'');
    const person=roster[staffId];
    sheet.getRange(rowNumber,10).setValue('取消通知中');
    SpreadsheetApp.flush();
    try{
      if(!token || !person || !validLineUid_(person.uid))throw new Error('員工未綁定 LINE 或缺少發送設定');
      const message='【排班取消通知】\n'+person.name+' 您好：\n日期：'+date+'\n班別：'+String(r[6]||'')+'\n此班次已取消或調整，請以最新排班為準。';
      const http=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
        method:'post',contentType:'application/json; charset=UTF-8',
        headers:{Authorization:'Bearer '+token},
        payload:JSON.stringify({to:person.uid,messages:[{type:'text',text:message}]}),
        muteHttpExceptions:true
      });
      if(http.getResponseCode()<200 || http.getResponseCode()>=300)throw new Error('LINE HTTP '+http.getResponseCode());
      sheet.getRange(rowNumber,10).setValue('已取消已通知');
      safeLog_('取消通知',String(r[1]||''),staffId,'成功',date,'');
      summary.sent++;
      summary.details.push({row:rowNumber,status:'sent'});
    }catch(err){
      sheet.getRange(rowNumber,10).setValue('已取消通知失敗');
      sheet.getRange(rowNumber,16).setValue(String(err.message).slice(0,450));
      safeLog_('取消通知',String(r[1]||''),staffId,'失敗',String(err.message).slice(0,200),'');
      summary.failed++;
      summary.details.push({row:rowNumber,status:'failed',message:String(err.message)});
    }
  });
  return summary;
}

function handleShiftReply_(input) {
  const uid = String(input.lineUid || '');
  const id = String(input.id || '');
  const version = Number(input.version);
  const choice = String(input.choice || '');
  if (!validLineUid_(uid) || !/^[a-f0-9-]{36}$/i.test(id) || !Number.isSafeInteger(version) || version < 1 || !['agree', 'disagree'].includes(choice)) {
    return {ok: true, code: 'invalid', message: '這張通知格式錯誤，請聯絡團長重新發送。'};
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = shiftSheet_();
    const rowNumber = findShiftRow_(sheet, id);
    if (!rowNumber) return {ok: true, code: 'missing', message: '找不到這筆排班，請聯絡團長。'};
    const row = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    if (typeof masaluCancelled_ === 'function' && masaluCancelled_(row)) {
      return {ok:true,code:'cancelled',message:'此班次已取消，舊通知不能再回覆。'};
    }
    const employeeId = String(row[2] || '').trim();
    let expectedUid;
    if (employeeId) {
      const employee = employeeIndex_()[employeeId];
      expectedUid = employee && employee.uid;
    } else {
      expectedUid = String(row[13] || '');
    }
    if (!expectedUid || expectedUid !== uid) {
      return {ok: true, code: 'wrong_user', message: '此排班通知不是發給這個 LINE 帳號，無法代替別人回覆。'};
    }
    if (Number(row[11]) !== version) return {ok: true, code: 'expired', message: '此按鈕已失效，請以團長最新的排班通知為準。'};
    const currentStatus = String(row[7] || '');
    if (currentStatus === '同意' || currentStatus === '不同意') {
      return {ok: true, code: 'already', message: '這筆排班已回覆「' + currentStatus + '」，若要修改請聯絡團長。'};
    }
    const status = choice === 'agree' ? '同意' : '不同意';
    sheet.getRange(rowNumber, 8, 1, 2).setValues([[status, new Date()]]);
    if (!employeeId) sheet.getRange(rowNumber, 14).clearContent(); // 一次性 UID 不留存
    safeLog_('員工回覆', id, employeeId, status, employeeId ? '' : '臨時UID已清除', String(input.eventId || ''));
    return {ok: true, code: 'confirmed', message: '已收到您的排班回覆：【' + status + '】，謝謝！'};
  } finally {
    lock.releaseLock();
  }
}

function handleLineText_(input) {
  const uid = String(input.lineUid || '');
  const text = String(input.text || '').trim();
  if (!validLineUid_(uid)) return {ok: true, code: 'invalid', message: '無法識別 LINE 帳號。'};
  if (!/^綁定[\s:：]+[A-Za-z0-9]{12}$/u.test(text)) {
    return {ok: true, code: 'invalid', message: '綁定格式：綁定 你的12碼綁定碼'};
  }
  const code = text.replace(/^綁定[\s:：]+/u, '').toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = openBook_().getSheetByName(SHIFT_APP.EMPLOYEES);
    const last = sheet.getLastRow();
    const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
    if (rows.some(row => String(row[3]).trim() === uid)) {
      return {ok: true, code: 'already_bound', message: '這個 LINE 帳號已完成綁定。如需更換員工身分，請找團長處理。'};
    }
    const index = rows.findIndex(row => String(row[2]).trim().toUpperCase() === code && !row[3]);
    if (index < 0) return {ok: true, code: 'invalid_code', message: '綁定碼無效或已使用，請向團長索取新的綁定碼。'};
    const rowNumber = index + 2;
    const name = String(rows[index][1]);
    sheet.getRange(rowNumber, 3, 1, 3).setValues([['', uid, new Date()]]);
    safeLog_('LINE綁定', '', String(rows[index][0]), '成功', name, String(input.eventId || ''));
    return {ok: true, code: 'bound', message: name + ' 您好，LINE 帳號綁定成功！之後可直接在這裡確認排班。'};
  } finally {
    lock.releaseLock();
  }
}

function cleanupExpiredTempRecipients() {
  const sheet = shiftSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return;
  const values = sheet.getRange(2, 1, last - 1, SHIFT_APP.SHIFT_HEADERS.length).getValues();
  values.forEach((row, i) => {
    if (row[2] || !row[13]) return;
    const expireAt = row[14] instanceof Date ? row[14].getTime() : new Date(row[14]).getTime();
    if (Number.isFinite(expireAt) && expireAt <= Date.now()) {
      sheet.getRange(i + 2, 14).clearContent();
      if (!row[7] || row[7] === '待回覆') sheet.getRange(i + 2, 8).setValue('已逾期');
    }
  });
}

function installDailyCleanupTrigger() {
  if (!ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === 'cleanupExpiredTempRecipients')) {
    ScriptApp.newTrigger('cleanupExpiredTempRecipients').timeBased().everyDays(1).create();
  }
  show_('已啟用每日清理；臨時 UID 在回覆或到期後會清除。');
}

function selectedShiftRow_() {
  const active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (active.getName() !== SHIFT_APP.SHIFTS) throw new Error('請先切換到「排班表」，再點選要處理的資料列。');
  const row = active.getActiveRange().getRow();
  if (row < 2) throw new Error('請選取標題列下方的某筆排班。');
  return row;
}

function findShiftRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const found = sheet.getRange(2, 2, last - 1, 1)
    .createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
  return found ? found.getRow() : 0;
}

function validLineUid_(uid) { return /^U[0-9a-f]{32}$/i.test(String(uid || '')); }

function constantTimeEqual_(given, expected) {
  if (!given || !expected) return false;
  let diff = given.length ^ expected.length;
  const n = Math.max(given.length, expected.length);
  for (let i = 0; i < n; i++) diff |= (given.charCodeAt(i) || 0) ^ (expected.charCodeAt(i) || 0);
  return diff === 0;
}

function openBook_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('尚未初始化：請先從表格選單執行「初始化試算表」。');
  return SpreadsheetApp.openById(id);
}

function shiftSheet_() {
  const sheet = openBook_().getSheetByName(SHIFT_APP.SHIFTS);
  if (!sheet) throw new Error('找不到「排班表」。');
  return sheet;
}

function safeLog_(action, shiftId, employeeId, result, note, eventId) {
  try {
    openBook_().getSheetByName(SHIFT_APP.LOG).appendRow([
      new Date(), action, shiftId, employeeId, result, String(note || '').slice(0, 300), eventId
    ]);
  } catch (err) { console.warn('操作紀錄寫入失敗：' + err.message); }
}

function show_(message) {
  try { SpreadsheetApp.getUi().alert(String(message)); }
  catch (_) { console.log(String(message)); } // 從 Apps Script 編輯器首次執行時可能尚無 UI
}
