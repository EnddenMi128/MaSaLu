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

function sendShiftRow_(rowNumber, force, shiftId, declinedVersion) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return {code: 'failed', message: '系統忙碌，請稍後再試。'};
  try {
    const sheet = shiftSheet_();
    if (shiftId) {
      rowNumber = findShiftRow_(sheet, shiftId);
      if (!rowNumber) return {code:'missing',message:'排班已不存在。'};
    }
    const current = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    if (declinedVersion !== undefined && (Number(current[11]) !== declinedVersion || current[7] !== '不同意')) {
      return {code:'expired',message:'這筆通知已有更新，請使用最新排班。'};
    }
    const declinedResend = force && current[7] === '不同意';
    const oldNote = String(current[16] || '');
    if (declinedResend) {
      const sourceMarker = '替補來源：'+String(current[1] || '');
      const hasReplacement = /\[替補ID:[a-f0-9-]{36}\]/i.test(oldNote) ||
        sheet.getRange(2,17,sheet.getLastRow()-1,1).getValues().some(row=>String(row[0] || '')===sourceMarker);
      if (hasReplacement) return {code:'replacement_exists',message:'這筆排班已安排替補，請先在月曆確認人員。'};
    }
    const display = sheet.getRange(rowNumber, 1, 1, SHIFT_APP.SHIFT_HEADERS.length).getDisplayValues()[0];
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
    const message = buildShiftFlex_(id,version,recipient.name,info);
    const retryKey = declinedResend ? masaluLineRetryKey_('shift-resend:'+id+':'+version+':'+recipient.uid) : '';
    sheet.getRange(rowNumber, 8, 1, 5).setValues([['待回覆', '', '發送中', '', version]]);
    sheet.getRange(rowNumber, 16).clearContent();
    SpreadsheetApp.flush();
    try {
      const http = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post',
        contentType: 'application/json; charset=UTF-8',
        headers: {Authorization:'Bearer '+token,...(retryKey?{'X-Line-Retry-Key':retryKey}:{})},
        payload: JSON.stringify({to: recipient.uid, messages: [message]}),
        muteHttpExceptions: true
      });
      const status = http.getResponseCode();
      if (!(status>=200 && status<300 || (retryKey && status===409 && masaluLineRetryAccepted_(http)))) {
        throw new Error('LINE HTTP '+status+'：'+http.getContentText().slice(0,250));
      }
      if (declinedResend && oldNote.includes('[管理已略過]')) {
        sheet.getRange(rowNumber,17).setValue(oldNote.replace(/(?:｜)?\[管理已略過\]/g,''));
      }
      sheet.getRange(rowNumber, 10, 1, 2).setValues([['已發送', new Date()]]);
      safeLog_('發送', id, String(current[2] || ''), '成功', recipient.name, '');
      return {code: 'sent', message: '第 ' + rowNumber + ' 列已成功向 LINE API 提交通知：' + recipient.name + '。（不代表對方已讀）'};
    } catch (err) {
      if (declinedResend) sheet.getRange(rowNumber,8,1,5).setValues([current.slice(7,12)]);
      else sheet.getRange(rowNumber,10).setValue('發送失敗');
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
    } else if (input.kind === 'shift_followup' || input.kind === 'contact_message') {
      result = masaluDeliverFollowup_(input,handleShiftFollowup_(input));
    } else if (input.kind === 'swap_response') {
      result = masaluDeliverFollowup_(input,masaluSwapResponse_(input));
    } else if (input.kind === 'manager_decision') {
      result = handleManagerDecision_(input);
    } else if (input.kind === 'bind_or_message') {
      result = handleLineText_(input);
    } else if (input.kind === 'calendar_api') {
      // Worker has already verified GAS_SHARED_SECRET. The calendar adapter
      // independently verifies the LINE access token and A01-A03 management role.
      result = masaluApi(String(input.action || ''), input.payload || {}, input.lineAccessToken);
      if (result.ok && (input.action === 'saveDay' || input.action === 'clearDay')) {
        try {
          result.notifications = masaluNotifyAddedForDate_(result.delta);
        } catch (notifyError) {
          result.notifications = {sent:0,failed:1,skipped:0,
            details:[{status:'failed',message:'排班已儲存，但通知流程失敗：'+String(notifyError.message)}]};
        }
        try {
          result.cancellationNotifications = masaluNotifyCancellations_(result.delta);
        } catch (notifyError) {
          result.cancellationNotifications = {sent:0,failed:1,skipped:0,
            details:[{status:'failed',message:'排班已更新，但取消通知失敗：'+String(notifyError.message)}]};
        }
        // Sending changes notification status/version in the sheet. The
        // optimistic revision must describe the FINAL persisted state.
        const fresh = masaluBootstrap_();
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

/** Send only NEW assignment IDs, never reissue an unchanged unanswered card. */
function masaluNotifyAddedForDate_(delta) {
  const date = masaluDate_(delta && delta.date);
  if (!date) throw new Error('通知日期格式錯誤');
  const addedIds = new Set((delta.addedIds || []).map(String));
  const sheet = shiftSheet_();
  const last = sheet.getLastRow();
  const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, SHIFT_APP.SHIFT_HEADERS.length).getValues();
  const candidates = [];
  rows.forEach((row, index) => {
    if (masaluDate_(row[3]) !== date || !row[2] || !row[6] || !addedIds.has(String(row[1]||''))) return;
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

/** A cleared date notifies all bound employees; normal edits notify only removed staff. */
function masaluNotifyCancellations_(delta) {
  const summary={sent:0,failed:0,skipped:0,details:[]};
  if (!delta || !delta.removed || !delta.removed.length) return summary;
  const roster=employeeIndex_();
  const token=PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('缺少 LINE_CHANNEL_ACCESS_TOKEN');
  const targets=new Map();
  if (delta.clear) {
    Object.keys(roster).forEach(id=>{
      const person=roster[id];
      if(validLineUid_(person.uid)) targets.set(person.uid,{
        uid:person.uid,id,name:person.name,all:true
      });
      else summary.skipped++;
    });
  } else {
    const sheet=shiftSheet_(),last=sheet.getLastRow();
    const added=new Set();
    if(last>=2) sheet.getRange(2,1,last-1,SHIFT_APP.SHIFT_HEADERS.length).getValues().forEach(r=>{
      if((delta.addedIds||[]).includes(String(r[1]||''))) added.add(String(r[2]||''));
    });
    delta.removed.forEach(old=>{
      if(added.has(old.staffId))return; // Changed shift: the replacement card communicates the update.
      const person=roster[old.staffId],uid=person?.uid||old.uid;
      if(validLineUid_(uid)) targets.set(uid,{
        uid,id:old.staffId,name:person?.name||old.name||'員工',shift:old.shift,all:false
      });
      else summary.skipped++;
    });
  }
  targets.forEach(person=>{
    const message=person.all
      ? '📢 排班異動通知\n'+delta.date+' 的排班已全部清空，目前當日沒有班次。請以最新月曆為準。'
      : '📢 排班異動通知\n'+person.name+' 您好，'+delta.date+' 的'+(person.shift?'「'+person.shift+'」':'')+'排班已取消。請以最新月曆為準。';
    try {
      const response=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
        method:'post',contentType:'application/json; charset=UTF-8',
        headers:{Authorization:'Bearer '+token},
        payload:JSON.stringify({to:person.uid,messages:[{type:'text',text:message}]}),
        muteHttpExceptions:true
      });
      if(response.getResponseCode()<200||response.getResponseCode()>=300)
        throw new Error('LINE HTTP '+response.getResponseCode());
      summary.sent++;
      safeLog_('取消通知','',person.id,'成功',delta.date,'');
    } catch(err) {
      summary.failed++;
      summary.details.push({status:'failed',id:person.id,message:String(err.message).slice(0,160)});
      safeLog_('取消通知','',person.id,'失敗',String(err.message).slice(0,160),'');
    }
  });
  return summary;
}

const MASALU_SERIALS = Object.freeze({agree:'A',disagree:'D',leave:'L'});
const MASALU_SERIAL_SHEET = 'MaSaLu流水號';
const MASALU_SERIAL_HEADERS = ['操作鍵','流水號','建立時間','排班ID','通知版次','員工編號','日期','班別','臨時驗證'];

function masaluSerial_(action, shift, create=true) {
  const book = openBook_();
  const sheet = book.getSheetByName(MASALU_SERIAL_SHEET) ||
    (create ? ensureSheet_(book,MASALU_SERIAL_SHEET,MASALU_SERIAL_HEADERS,false) : null);
  if (!sheet) return null;
  const data = masaluTable_(MASALU_SERIAL_SHEET,MASALU_SERIAL_HEADERS,book);
  const key = [shift.id,shift.version,action].join(':');
  const existing = data.rows.find(row=>row[0]===key);
  if (existing) return {serial:String(existing[1]),proof:String(existing[8] || '')};
  if (!create) return null;
  const prefix = MASALU_SERIALS[action]+'-';
  const last = data.rows.reduce((max,row)=>{
    const serial = String(row[1] || '');
    return serial.startsWith(prefix) ? Math.max(max,Number(serial.slice(prefix.length)) || 0) : max;
  },0);
  const serial = prefix+String(last+1).padStart(6,'0');
  sheet.appendRow([key,serial,new Date(),shift.id,shift.version,shift.employeeId,
    shift.date,shift.shift,shift.proof || '']);
  SpreadsheetApp.flush();
  return {serial,proof:shift.proof || ''};
}

function handleShiftReply_(input) {
  const uid = String(input.lineUid || '');
  const id = String(input.id || '');
  const version = Number(input.version);
  const choice = String(input.choice || '');
  if (!validLineUid_(uid) || !/^[a-f0-9-]{36}$/i.test(id) || !Number.isSafeInteger(version) || version < 1 || !['agree', 'disagree'].includes(choice)) {
    return {ok:true,code:'invalid',message:'這張通知格式錯誤，請聯絡團長重新發送。'};
  }
  const status = choice === 'agree' ? '同意' : '不同意';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let confirmed, repeated;
  try {
    const sheet = shiftSheet_();
    const rowNumber = findShiftRow_(sheet,id);
    if (!rowNumber) return {ok:true,code:'missing',message:'找不到這筆排班，請聯絡團長。'};
    const row = sheet.getRange(rowNumber,1,1,SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    const employeeId = String(row[2] || '').trim();
    const employee = employeeId ? employeeIndex_()[employeeId] : null;
    const currentStatus = String(row[7] || '');
    repeated = currentStatus === status;
    const expectedUid = employeeId ? employee?.uid : String(row[13] || '');
    let ownsReply = expectedUid === uid;
    if (!ownsReply && !employeeId && !expectedUid && repeated && Number(row[11]) === version) {
      const prior = masaluSerial_(choice,{id,version},false);
      ownsReply = !!prior?.proof && constantTimeEqual_(prior.proof,masaluTemporaryFollowupToken_(id,version,uid));
    }
    if (!ownsReply) return {ok:true,code:'wrong_user',message:'此排班通知不是發給這個 LINE 帳號，無法代替別人回覆。'};
    if (Number(row[11]) !== version) return {ok:true,code:'expired',message:'此按鈕已失效，請以團長最新的排班通知為準。'};
    if (!repeated && ['同意','不同意'].includes(currentStatus)) {
      return {ok:true,code:'already',message:'這筆排班已回覆「'+currentStatus+'」。如需更改，請聯絡團長重發。'};
    }
    const proof = employeeId ? '' : masaluTemporaryFollowupToken_(id,version,uid);
    confirmed = {id,version,status,employeeId,proof,date:masaluDate_(row[3]),
      shift:String(row[6] || ''),name:employee?.name || String(row[12] || '臨時人員')};
    confirmed.serial = masaluSerial_(choice,confirmed).serial;
    if (!repeated) {
      sheet.getRange(rowNumber,8,1,2).setValues([[status,new Date()]]);
      if (!employeeId) sheet.getRange(rowNumber,14).clearContent();
      SpreadsheetApp.flush();
      safeLog_('員工回覆',id,employeeId,status,confirmed.serial,String(input.eventId || ''));
    }
  } finally {
    lock.releaseLock();
  }

  const code = repeated ? 'already' : 'confirmed';
  const receipt = choice === 'agree'
    ? masaluAcceptedReply_(code,{...confirmed,token:confirmed.proof})
    : {ok:true,code,serial:confirmed.serial,message:['已回覆不同意排班',confirmed.date,confirmed.shift,
      '流水號：'+confirmed.serial].filter(Boolean).join('\n')};
  const delivered = masaluDeliverShiftReceipt_(uid,receipt,input.eventId);
  const result = delivered ? {ok:true,code,serial:confirmed.serial} : receipt;
  if (repeated) return result;

  let managerNotice = {sent:0,failed:0};
  try { managerNotice = masaluNotifyManagers_(confirmed,choice === 'agree' ? uid : ''); }
  catch (err) {
    console.error('團長通知失敗：'+String(err.message).slice(0,200));
    safeLog_('團長通知',id,confirmed.employeeId,'失敗',String(err.message).slice(0,200),'');
    managerNotice.failed++;
  }
  return {...result,managerNotice};
}

function masaluLineRetryKey_(identity) {
  const hex = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,identity,Utilities.Charset.UTF_8)
    .map(byte=>(byte & 255).toString(16).padStart(2,'0')).join('');
  return [hex.slice(0,8),hex.slice(8,12),'4'+hex.slice(13,16),'8'+hex.slice(17,20),hex.slice(20,32)].join('-');
}

function masaluLineRetryAccepted_(response) {
  return Object.keys(response.getAllHeaders()).some(name=>name.toLowerCase()==='x-line-accepted-request-id');
}

function masaluPushWorkflowMessage_(uid, message, identity) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return false;
  const headers = {Authorization:'Bearer '+token,'X-Line-Retry-Key':masaluLineRetryKey_(identity+':'+uid)};
  for (let attempt=0; attempt<2; attempt++) {
    try {
      const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
        method:'post',contentType:'application/json; charset=UTF-8',headers,
        payload:JSON.stringify({to:uid,messages:[message]}),muteHttpExceptions:true
      });
      const status=response.getResponseCode();
      if (status>=200 && status<300 || status===409 && masaluLineRetryAccepted_(response)) return true;
      console.error('LINE 推播失敗：HTTP '+status);
      if (status<500) return false;
    } catch (err) {
      console.error('LINE 轉送失敗：'+String(err.message).slice(0,120));
    }
  }
  return false;
}

function masaluDeliverFollowup_(input, result) {
  if (!result.message || !input.eventId || !validLineUid_(input.lineUid)) return result;
  const message = {type:'text',text:result.message,...(result.quickReply?{quickReply:result.quickReply}:{})};
  if (!masaluPushWorkflowMessage_(input.lineUid,message,'shift-followup:'+input.eventId)) return result;
  return {ok:true,code:result.code,delivered:true};
}

function masaluDeliverShiftReceipt_(uid, receipt, eventId) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) return false;
  try {
    const identity = 'shift-receipt:'+(eventId || receipt.serial);
    const text = {type:'text',text:receipt.message,...(receipt.quickReply?{quickReply:receipt.quickReply}:{})};
    const push = (message,key) => UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
      method:'post',contentType:'application/json; charset=UTF-8',
      headers:{Authorization:'Bearer '+token,'X-Line-Retry-Key':masaluLineRetryKey_(key)},
      payload:JSON.stringify({to:uid,messages:[message]}),muteHttpExceptions:true
    });
    let response = push(receipt.replyMessage || text,identity);
    if (response.getResponseCode()===400 && receipt.replyMessage) response = push(text,identity+':text');
    const status = response.getResponseCode();
    if (status>=200 && status<300 || status===409 && masaluLineRetryAccepted_(response)) return true;
    console.error('員工確認推播失敗：LINE HTTP '+status);
  } catch (err) {
    console.error('員工確認推播失敗：'+String(err.message).slice(0,120));
  }
  return false;
}

const MASALU_LEAVE_TYPES = Object.freeze({personal:'事假', sick:'病假', other:'其他'});
const MASALU_FOLLOWUP_HEADERS = [
  '請求識別碼','建立時間','排班ID','通知版次','員工編號','姓名','排班日期','班別',
  '請求類型','假別','留言','狀態','已通知團長'
];

function masaluTemporaryFollowupToken_(id, version, uid) {
  const secret = PropertiesService.getScriptProperties().getProperty('GAS_SHARED_SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(
    ['masalu-shift-followup', id, version, uid].join(':'), secret
  )).replace(/=+$/, '');
}

function masaluFollowupAction_(shift, choice, label) {
  const token = shift.token ? '&proof=' + encodeURIComponent(shift.token) : '';
  return {type:'postback',label,displayText:label,
    data:'action=shift_followup&id='+encodeURIComponent(shift.id)+'&v='+shift.version+'&choice='+choice+token};
}

function masaluAcceptedReply_(code, shift) {
  const title = '✅ '+shift.name+' 已同意排班';
  const details = [shift.date,shift.shift].filter(Boolean).join(' ・ ');
  const button = (label, choice) => ({
    type:'button',style:'secondary',height:'sm',action:masaluFollowupAction_(shift,choice,label)
  });
  const buttons = [button('聯絡團長','contact'),button('調班請求','swap'),button('我要請假','leave')];
  return {ok:true,code,serial:shift.serial,message:[title,details,shift.serial].filter(Boolean).join('\n'),
    quickReply:{items:buttons.map(item=>({type:'action',action:item.action}))},replyMessage:{
    type:'flex',altText:title+(details?'｜'+details:''),
    contents:{type:'bubble',
      body:{type:'box',layout:'vertical',spacing:'sm',contents:[
        {type:'box',layout:'horizontal',spacing:'sm',contents:[
          {type:'text',text:title,weight:'bold',size:'lg',wrap:true,flex:1},
          {type:'text',text:shift.serial,size:'xs',color:'#64748B',align:'end',flex:0}
        ]},
        ...(details ? [{type:'text',text:details,wrap:true,size:'sm'}] : [])
      ]},
      footer:{type:'box',layout:'vertical',spacing:'sm',contents:buttons}
    }
  }};
}

function handleShiftFollowup_(input) {
  const uid = String(input.lineUid || '');
  if (!validLineUid_(uid)) return {ok:true,code:'invalid',message:'無法識別 LINE 帳號。'};
  const chatText = input.kind === 'contact_message' ? String(input.text || '').trim() : '';
  const chatKey = 'masalu:contact:'+uid;
  let note = '';
  if (input.text && input.kind !== 'contact_message') {
    const match = /^排班留言\s+([A-F0-9]{8})(?:\s+([\s\S]*))?$/i.exec(String(input.text).trim());
    if (!match || !String(match[2] || '').trim() || match[2].trim().length > 300) {
      return {ok:true,code:'invalid_note',message:'請在留言代碼後輸入 1 到 300 字的內容，再送出。'};
    }
    const saved = CacheService.getScriptCache().get('masalu:shift-note:'+uid+':'+match[1].toUpperCase());
    if (!saved) return {ok:true,code:'expired_note',message:'留言連結已失效，請從原排班確認卡重新操作。'};
    input = JSON.parse(saved);
    note = match[2].trim();
  }

  if (input.kind === 'contact_message' && !CacheService.getScriptCache().get(chatKey)) return {ok:true,code:'ignored'};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const eventId = String(input.eventId || '');
    const eventTime = Number(input.timestamp) || Date.now();
    const cache = CacheService.getScriptCache();
    if (input.kind === 'contact_message' || input.choice === 'contact_end') {
      const chat = JSON.parse(cache.get(chatKey) || 'null');
      if (!chat || eventTime<chat.startedAt) return {ok:true,code:'ignored'};
      if (input.choice === 'contact_end' && (input.id!==chat.id || input.version!==chat.version))
        return {ok:true,code:'ignored'};
      if (chat.ended) return eventId && eventId===chat.endEventId
        ? {ok:true,code:'contact_ended',message:'已結束對話。'} : {ok:true,code:'ignored'};
      if (input.choice === 'contact_end' || chatText === '結束對話') {
        cache.put(chatKey,JSON.stringify({...chat,startedAt:eventTime,ended:true,endEventId:eventId}),1800);
        return {ok:true,code:'contact_ended',message:'已結束對話。'};
      }
      input = {...chat,choice:'contact_send'};
    }
    const {id,version,choice} = input;
    if (!/^[a-f0-9-]{36}$/i.test(id || '') || !Number.isSafeInteger(version) || version < 1 ||
        !/^(contact(?:_send)?|swap(?:_date|_shift|_target)?|leave(?:_(?:personal|sick|other)(?:_send)?)?)$/.test(choice || '')) {
      return {ok:true,code:'invalid',message:'請從原排班確認卡重新操作。'};
    }
    const sheet = shiftSheet_();
    const rowNumber = findShiftRow_(sheet,id);
    if (!rowNumber) return {ok:true,code:'missing',message:'這筆排班已取消或找不到。'};
    const row = sheet.getRange(rowNumber,1,1,SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    const employeeId = String(row[2] || '').trim();
    const roster = employeeIndex_();
    const employee = roster[employeeId];
    const owner = employeeId ? employee?.uid === uid :
      constantTimeEqual_(String(input.token || ''),masaluTemporaryFollowupToken_(id,version,uid));
    if (!owner) return {ok:true,code:'wrong_user',message:'僅限這筆排班的本人操作。'};
    if (Number(row[11]) !== version || row[7] !== '同意') {
      return {ok:true,code:'expired',message:'這筆已同意的排班有變更，請從最新通知操作。'};
    }

    const shift = {id,version,token:employeeId?'':input.token,date:masaluDate_(row[3]),
      shift:String(row[6] || ''),employeeId,name:employee?.name || String(row[12] || '臨時人員')};
    if (choice === 'contact') {
      const previous = JSON.parse(cache.get(chatKey) || 'null');
      if (previous && eventTime <= previous.startedAt) {
        if (previous.ended || !eventId || eventId!==previous.startEventId) return {ok:true,code:'ignored'};
      } else {
        cache.put(chatKey,JSON.stringify({id,version,token:shift.token,startedAt:eventTime,startEventId:eventId}),1800);
      }
      return masaluContactReply_(shift,'contact_open','請直接輸入文字，訊息會轉送給 A01～A03。');
    }
    if (choice === 'contact_send') {
      cache.put(chatKey,JSON.stringify({id,version,token:shift.token,startedAt:input.startedAt,startEventId:input.startEventId}),1800);
      if (!chatText || chatText.length>4000 || !eventId)
        return masaluContactReply_(shift,'contact_invalid','請輸入 1 到 4000 字的訊息後再送出。');
      const managerUids = ['A01','A02','A03'].map(id=>roster[id]?.uid);
      const recipients = [...new Set(managerUids.filter(validLineUid_))];
      const text = ['💬 '+shift.name+'｜聯絡團長',shift.date+' ・ '+shift.shift,'',chatText].join('\n');
      const sent = recipients.filter(to=>masaluPushWorkflowMessage_(to,{type:'text',text},'contact:'+eventId)).length;
      const message = sent===0 ? '訊息未能轉送給團長。' : sent<recipients.length || !managerUids.every(validLineUid_)
        ? '訊息已轉送給 '+sent+' 位團長，其餘團長暫時無法收到。' : '訊息已轉送給團長。';
      return masaluContactReply_(shift,'contact_sent',message);
    }
    if (choice.startsWith('swap')) {
      shift.shiftLabel = masaluSwapShiftText_(row);
      const result = masaluSwapSelection_(input,shift,roster,uid,sheet);
      return {...result,message:'🔄 調班申請\n原本的班｜'+shift.date+' ・ '+shift.shiftLabel+'\n\n'+result.message};
    }
    if (choice === 'leave') {
      return {ok:true,code:'leave_types',message:'請選擇假別',quickReply:{items:
        Object.entries(MASALU_LEAVE_TYPES).map(([type,label])=>({
          type:'action',action:masaluFollowupAction_(shift,'leave_'+type,label)
        }))
      }};
    }
    const kind = choice.split('_')[0];
    const leaveType = kind === 'leave' ? MASALU_LEAVE_TYPES[choice.split('_')[1]] : '';
    if (!choice.endsWith('_send')) {
      return masaluFollowupNotePrompt_(uid,shift,choice,leaveType);
    }
    return masaluSubmitFollowup_(shift,kind,leaveType,note,roster);
  } finally {
    lock.releaseLock();
  }
}

function masaluContactReply_(shift, code, message) {
  return {ok:true,code,message,quickReply:{items:[
    {type:'action',action:{type:'postback',label:'輸入訊息',data:'action=shift_note_keyboard',inputOption:'openKeyboard'}},
    {type:'action',action:masaluFollowupAction_(shift,'contact_end','結束對話')}
  ]}};
}

function masaluFollowupNotePrompt_(uid, shift, choice, leaveType) {
  const code = Utilities.getUuid().slice(0,8).toUpperCase();
  const submit = choice+'_send';
  CacheService.getScriptCache().put('masalu:shift-note:'+uid+':'+code,
    JSON.stringify({id:shift.id,version:shift.version,choice:submit,token:shift.token}),900);
  const prefix = '排班留言 '+code+' ';
  return {ok:true,code:'request_note',
    message:(leaveType?'請假｜'+leaveType:'調班請求')+'\n'+shift.date+' ・ '+shift.shift+
      '\n可加留言或直接送出，資料會交給團長處理。'+
      '\n如需手動輸入留言，請以「'+prefix.trim()+'」開頭，留一格再寫內容。',
    quickReply:{items:[
      {type:'action',action:{type:'postback',label:'加留言',data:'action=shift_note_keyboard',
        inputOption:'openKeyboard',fillInText:prefix}},
      {type:'action',action:masaluFollowupAction_(shift,submit,'直接送出')}
    ]}
  };
}

function masaluSubmitFollowup_(shift, kind, leaveType, note, roster, caseId) {
  const label = {swap:'調班請求',leave:'請假申請'}[kind];
  const book = openBook_();
  const sheetName = 'MaSaLu排班請求';
  const sheet = book.getSheetByName(sheetName) || ensureSheet_(book,sheetName,MASALU_FOLLOWUP_HEADERS,false);
  const data = masaluTable_(sheetName,MASALU_FOLLOWUP_HEADERS,book);
  const key = caseId || [shift.id,shift.version,kind].join(':');
  const serial = kind === 'leave' ? masaluSerial_(kind,shift).serial : '';
  const receipt = (code,message) => ({ok:true,code,message:message+(serial?'\n流水號：'+serial:''),...(serial?{serial}:{})});
  const index = data.rows.findIndex(row=>row[0]===key);
  const saved = index < 0 ? null : data.rows[index];
  const rowNumber = index < 0 ? sheet.getLastRow()+1 : index+2;
  const delivered = new Set(String(saved?.[12] || '').split(',').filter(Boolean));
  if (!saved) {
    const safeNote = /^[=+\-@]/.test(note) ? "'"+note : note;
    sheet.appendRow([key,new Date(),shift.id,shift.version,shift.employeeId,shift.name,
      shift.date,shift.shift,label,leaveType,safeNote,caseId?'待團長審核':'待處理','']);
    SpreadsheetApp.flush();
  }

  const managerIds = ['A01','A02','A03'].filter(id=>validLineUid_(roster[id]?.uid));
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  const storedType = saved ? saved[9] : leaveType;
  const storedNote = saved ? saved[10] : note;
  const text = [caseId?label+'｜待團長審核':label,caseId?'案件 '+caseId:'','人員：'+shift.name,'日期：'+shift.date,'班別：'+shift.shift,
    serial?'流水號：'+serial:'',storedType?'假別：'+storedType:'',storedNote?(caseId?'':'留言：')+storedNote:'',
    caseId?'請至試算表「MaSaLu排班請求」審核。':''].filter(Boolean).join('\n');
  let notified = false;
  if (token) managerIds.filter(id=>!delivered.has(id)).forEach(id=>{
    try {
      const recipient=roster[id].uid;
      const accepted=caseId
        ? masaluPushWorkflowMessage_(recipient,{type:'text',text},'swap-review:'+caseId)
        : masaluPushFollowupNotice_(recipient,text,token);
      if (accepted) {
        delivered.add(id);
        sheet.getRange(rowNumber,13).setValue([...delivered].join(','));
        SpreadsheetApp.flush();
        notified = true;
      }
    } catch (err) {
      console.error('排班請求通知失敗：'+String(err.message).slice(0,120));
    }
  });
  if (!managerIds.length || managerIds.some(id=>!delivered.has(id))) {
    return receipt('request_saved',label+'已登記。團長通知尚未全部送達，可再次點選重試。');
  }
  return receipt(saved && !notified?'request_already':'request_sent',label+'已送出，請等候團長回覆。');
}

function masaluPushFollowupNotice_(uid,text,token) {
  const response=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
    method:'post',contentType:'application/json; charset=UTF-8',headers:{Authorization:'Bearer '+token},
    payload:JSON.stringify({to:uid,messages:[{type:'text',text}]}),muteHttpExceptions:true
  });
  return response.getResponseCode()>=200 && response.getResponseCode()<300;
}

/** Notify A01–A03, skipping the individual update for a manager receiving their own acceptance card. */
function masaluNotifyManagers_(notice, acceptanceRecipientUid) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) throw new Error('缺少 LINE_CHANNEL_ACCESS_TOKEN');
  const managers = employeeIndex_();
  const uids = [...new Set(['A01', 'A02', 'A03']
    .map(id => managers[id] && managers[id].uid)
    .filter(validLineUid_))];
  if (!uids.length) throw new Error('A01–A03 尚未綁定可通知的 LINE 帳號');
  const summary = {sent:0,failed:0};
  const managerMessage = masaluManagerReplyMessage_(notice);
  const publication = notice.status === '同意' ? masaluReadyDaySummary_(notice.date) : '';
  uids.forEach(uid => {
    const messages = uid === acceptanceRecipientUid ? [] : [managerMessage];
    if (publication) messages.push({type:'text',text:publication});
    if (!messages.length) return;
    try {
      const result = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method:'post',contentType:'application/json; charset=UTF-8',
        headers:{Authorization:'Bearer '+token},
        payload:JSON.stringify({to:uid,messages}),muteHttpExceptions:true
      });
      if (result.getResponseCode() < 200 || result.getResponseCode() >= 300)
        throw new Error('LINE HTTP '+result.getResponseCode());
      summary.sent++;
    } catch (err) {
      summary.failed++;
      console.error('團長推播失敗：'+String(err.message).slice(0, 120));
    }
  });
  safeLog_('團長通知', notice.id, notice.employeeId,
    summary.failed ? '部分失敗' : '成功', '成功 '+summary.sent+'，失敗 '+summary.failed, '');
  if(publication && summary.sent) safeLog_('總表發布',notice.id,notice.employeeId,'完成',notice.date,'');
  return summary;
}

function masaluManagerReplyMessage_(notice) {
  const label = notice.name+' '+(notice.status==='同意'?'已同意':'無法出席')+'排班';
  if (notice.status === '同意') return {type:'text',text:['✅ '+label,'日期：'+notice.date,'班別：'+notice.shift,notice.serial].filter(Boolean).join('\n')};
  const serial = notice.serial ? '流水號：'+notice.serial : '';
  const action = (text, decision, color) => ({
    type:'button',style:'primary',color,
    action:{type:'postback',label:text,displayText:text+'：'+notice.name,
      data:'action=manager_shift&id='+encodeURIComponent(notice.id)+'&v='+notice.version+'&decision='+decision}
  });
  return {type:'flex',altText:'⚠️ '+label,
    contents:{type:'bubble',
      body:{type:'box',layout:'vertical',spacing:'md',contents:[
        {type:'text',text:'⚠️ '+label,weight:'bold',wrap:true,size:'lg'},
        {type:'text',text:'日期：'+notice.date,wrap:true,size:'sm'},
        {type:'text',text:'班別：'+notice.shift,wrap:true,size:'sm'},
        ...(serial?[{type:'text',text:serial,wrap:true,size:'sm'}]:[]),
        {type:'text',text:'原排班已從月曆隱藏。可略過、安排替補，或溝通後重發。',wrap:true,size:'sm'}
      ]},
      footer:{type:'box',layout:'vertical',spacing:'sm',contents:[
        action('略過','skip','#64748B'),action('替補','replace','#D96B24'),action('重發排班','resend','#2563EB')
      ]}
    }};
}

/** Publish an A01–A03-only day summary once all active assignments have agreed. */
function masaluReadyDaySummary_(date) {
  if (!date) return '';
  const sheet = shiftSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return '';
  const rows = sheet.getRange(2, 1, last-1, SHIFT_APP.SHIFT_HEADERS.length).getValues()
    .filter(r => masaluDate_(r[3]) === date && r[6] && (r[2] || r[12]));
  if (rows.some(r => String(r[7]||'') === '不同意' && !String(r[16]||'').includes('[管理已略過]'))) return '';
  const active = rows.filter(r => String(r[7]||'') !== '不同意');
  if (!active.length || active.some(r => String(r[7]||'') !== '同意')) return '';
  const roster = employeeIndex_();
  const details = active.map(r => {
    const employee = roster[String(r[2]||'').trim()];
    return '• '+(employee ? employee.name : String(r[12]||'臨時人員'))+'｜'+String(r[6]||'');
  }).join('\n');
  return ('📅 '+date+' 總排班已全部確認\n'+details).slice(0, 1800);
}

/** Actions on decline notifications are allowed only for bound A01–A03 managers. */
function handleManagerDecision_(input) {
  const result = masaluManagerDecisionLocked_(input);
  if (result.code === 'resend') {
    const sent = sendShiftRow_(null,true,input.id,Number(input.version));
    return {ok:true,code:sent.code,message:sent.code==='sent'?'已重新發送，請員工從新通知回覆。':sent.message};
  }
  if (!result.replacementId) return result;
  const sent = sendShiftRow_(null, false, result.replacementId);
  if (sent.code === 'missing') return {ok:true,code:'missing',message:sent.message};
  if (sent.code === 'failed') {
    const retry = masaluReplacementButton_('重試發送', input, 'select&staff='+encodeURIComponent(input.staffId), '重試發送替補通知');
    return {ok:true,code:'replaced',notification:sent.code,
      message:'替補已儲存，通知發送失敗。請點「重試發送」。',quickReply:{items:[retry]}};
  }
  return {ok:true,code:'replaced',notification:sent.code,
    message:'已安排 '+result.replacementName+' 替補，排班確認通知已送出。'};
}

function masaluReplacementCandidates_(dayRows, roster) {
  const unavailable = new Set(dayRows.map(r => String(r[2] || '').trim()));
  return Object.keys(roster).filter(id => !unavailable.has(id) && validLineUid_(roster[id].uid));
}

function masaluReplacementButton_(label, input, decision, displayText) {
  const data = 'action=manager_shift&id='+encodeURIComponent(input.id)+'&v='+Number(input.version)+'&decision='+decision;
  return {type:'action',action:{type:'postback',label:label.slice(0,20),data,displayText:displayText.slice(0,100)}};
}

function masaluReplacementPicker_(ids, row, roster, input) {
  if (!ids.length) return {ok:true,code:'no_candidates',message:'當日沒有可替補的尚未排班人員。'};
  const requestedPage = Number(input.page);
  const page = Number.isSafeInteger(requestedPage) && requestedPage > 0
    ? Math.min(requestedPage, Math.floor((ids.length-1)/10)) : 0;
  const items = ids.slice(page*10,page*10+10).map(id => {
    const name = roster[id].name || id;
    return masaluReplacementButton_(name,input,'select&staff='+encodeURIComponent(id),'選擇替補：'+name);
  });
  if (page > 0) items.push(masaluReplacementButton_('上一頁',input,'replace&page='+(page-1),'上一頁替補人員'));
  if ((page+1)*10 < ids.length) items.push(masaluReplacementButton_('下一頁',input,'replace&page='+(page+1),'下一頁替補人員'));
  items.push(masaluReplacementButton_('取消',input,'cancel','取消選擇替補'));
  return {ok:true,code:'replace',message:'請選擇替補對象\n日期：'+masaluDate_(row[3])+'\n班別：'+String(row[6])+'\n以下為當日尚未排班人員',quickReply:{items}};
}

function masaluManagerDecisionLocked_(input) {
  const uid=String(input.lineUid||''),id=String(input.id||'');
  const version=Number(input.version),decision=String(input.decision||'');
  if (!validLineUid_(uid) || !/^[a-f0-9-]{36}$/i.test(id) ||
      !Number.isSafeInteger(version) || version<1 || !['skip','replace','select','cancel','resend'].includes(decision))
    return {ok:true,code:'invalid',message:'管理操作格式不正確。'};
  const roster=employeeIndex_();
  if (!['A01','A02','A03'].some(key=>roster[key]&&roster[key].uid===uid))
    return {ok:true,code:'unauthorized',message:'只有 A01–A03 能處理替補。'};
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet=shiftSheet_();
    const number=findShiftRow_(sheet,id);
    if (!number) return {ok:true,code:'missing',message:'排班已不存在。'};
    const row=sheet.getRange(number,1,1,SHIFT_APP.SHIFT_HEADERS.length).getValues()[0];
    if (Number(row[11])!==version) return {ok:true,code:'expired',message:'通知已失效。'};
    if (String(row[7]||'')!=='不同意') return {ok:true,code:'expired',message:'這筆排班並非「不同意」狀態。'};
    if (decision==='resend') return {ok:true,code:'resend'};
    const date=masaluDate_(row[3]);
    if (decision==='cancel') return {ok:true,code:'cancelled',message:'已取消選擇替補。'};
    if (decision==='skip') {
      const note=String(row[16]||'');
      if (note.includes('[管理已略過]')) return {ok:true,code:'already',message:'這筆無法出席排班已略過。'};
      sheet.getRange(number,17).setValue((note ? note+'｜' : '')+'[管理已略過]');
      safeLog_('無法出席處理',id,String(row[2]||''),'略過',date,String(input.eventId||''));
      return {ok:true,code:'skipped',message:'已略過 '+date+' 的無法出席排班，原始紀錄保留。'};
    }
    if (!date || !row[6]) return {ok:true,code:'expired',message:'原排班的日期或班別已變更。'};
    const dayRows=sheet.getRange(2,1,sheet.getLastRow()-1,SHIFT_APP.SHIFT_HEADERS.length).getValues()
      .filter(r=>masaluDate_(r[3])===date && r[6]);
    const sourceMarker='替補來源：'+id;
    const replacement=String(row[16]||'').match(/\[替補ID:([a-f0-9-]{36})\]/i);
    const existing=dayRows.find(r=>String(r[16]||'')===sourceMarker || (replacement && r[1]===replacement[1]));
    if (existing || replacement) {
      if (existing && !replacement) {
        sheet.getRange(number,17).setValue(String(row[16]||'')+'[管理已略過][替補ID:'+existing[1]+']');
        SpreadsheetApp.flush();
      }
      if (decision==='select' && existing && String(existing[2])===String(input.staffId) &&
          ['未發送','發送失敗','發送中'].includes(String(existing[9]))) {
        return {ok:true,code:'replaced',replacementId:existing[1],replacementName:roster[existing[2]]?.name||String(existing[2])};
      }
      return {ok:true,code:'already',message:'這筆排班已安排替補，請查看最新月曆。'};
    }
    const candidates=masaluReplacementCandidates_(dayRows,roster);
    if (decision==='replace') return masaluReplacementPicker_(candidates,row,roster,input);
    const staffId=String(input.staffId||'');
    if (!candidates.includes(staffId)) {
      const picker=masaluReplacementPicker_(candidates,row,roster,input);
      picker.message='該人員目前無法替補。\n'+picker.message;
      return picker;
    }
    const replacementId=Utilities.getUuid();
    sheet.appendRow([false,replacementId,staffId,row[3],row[4],row[5],row[6],
      '待回覆','','未發送','',0,'','','','',sourceMarker]);
    sheet.getRange(number,17).setValue(String(row[16]||'')+'[管理已略過][替補ID:'+replacementId+']');
    SpreadsheetApp.flush();
    safeLog_('安排替補',id,staffId,'成功',date,String(input.eventId||''));
    return {ok:true,code:'replaced',replacementId,replacementName:roster[staffId].name||staffId};
  } finally { lock.releaseLock(); }
}

function handleLineText_(input) {
  const uid = String(input.lineUid || '');
  const text = String(input.text || '').trim();
  if (!validLineUid_(uid)) return {ok: true, code: 'invalid', message: '無法識別 LINE 帳號。'};
  const match = /^([^\r\n]+)([A-Za-z0-9]{12})$/u.exec(text);
  if (!match || /^綁定[\s:：]+[A-Za-z0-9]{12}$/u.test(text)) {
    return {ok: true, code: 'invalid', message: '請直接貼上員工名單的姓名和12碼綁定碼，例如：王小明 ABCDEF123456'};
  }
  const name = match[1].trim();
  const code = match[2].toUpperCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = openBook_().getSheetByName(SHIFT_APP.EMPLOYEES);
    const last = sheet.getLastRow();
    const rows = last < 2 ? [] : sheet.getRange(2, 1, last - 1, 5).getDisplayValues();
    if (rows.some(row => String(row[3]).trim() === uid)) {
      return {ok: true, code: 'already_bound', message: '這個 LINE 帳號已完成綁定。如需更換員工身分，請找團長處理。'};
    }
    const index = rows.findIndex(row =>
      String(row[1]).trim() === name && String(row[2]).trim().toUpperCase() === code && !row[3]);
    if (index < 0) return {ok: true, code: 'invalid_code', message: '姓名或綁定碼不符，或綁定碼已使用。請確認姓名和綁定碼來自同一列，或聯絡團長。'};
    const rowNumber = index + 2;
    sheet.getRange(rowNumber, 3, 1, 3).setValues([['', uid, new Date()]]);
    SpreadsheetApp.flush();
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


