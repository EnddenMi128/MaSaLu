/** Additive adapter for the supplied Code.gs. Keep its doPost and LINE functions. */
const MASALU = Object.freeze({
  SHIFTS: 'MaSaLu班別', DISPLAY: 'MaSaLu人員顯示',
  SHIFT_HEADERS: ['id','name','isOff'], DISPLAY_HEADERS: ['id','short','color'],
  COLORS: ['#7C3AED','#2563EB','#047857','#BE185D','#B45309','#0E7490','#4338CA','#B91C1C']
});

/** Run once manually. Creates only missing configuration tabs. Never clears data. */
function setupMaSaLuCalendar() {
  const book=openBook_();
  if(!book.getSheetByName(MASALU.SHIFTS)) {
    const sheet=book.insertSheet(MASALU.SHIFTS);
    sheet.appendRow(MASALU.SHIFT_HEADERS);
    const existing=shiftSheet_().getDataRange().getDisplayValues().slice(1);
    const names=[...new Set(existing.map(r=>String(r[6]||'').trim()).filter(Boolean))];
    (names.length ? names : ['玉門','田園','休假']).forEach(name=>sheet.appendRow([Utilities.getUuid(),name,name==='休假']));
  }
  if(!book.getSheetByName(MASALU.DISPLAY)) book.insertSheet(MASALU.DISPLAY).appendRow(MASALU.DISPLAY_HEADERS);
  show_('月曆設定分頁已就緒。已綁定 LINE 的 A01–A03 可透過 LINE 登入管理。');
}

/** Existing public JSONP read-only endpoint retained for legacy calendar consumers. */
function doGet(e) {
  const q=e && e.parameter || {};
  if(q.action || q.callback) {
    if(!/^[A-Za-z_$][A-Za-z0-9_$]{0,100}$/.test(q.callback||'')) return masaluJson_({ok:false,error:'Invalid callback'});
    let result;
    try { result=q.action==='bootstrap' ? masaluBootstrap_() : {ok:false,error:'此入口只提供讀取'}; }
    catch(err) { result={ok:false,error:String(err.message)}; }
    const json=JSON.stringify(result).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
    return ContentService.createTextOutput(q.callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return masaluJson_({ok:true,message:'MaSaLu 資料接口；請由 GitHub Pages 開啟排班網站'});
}

function masaluJson_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function masaluTable_(name,headers) {
  const sheet=openBook_().getSheetByName(name);
  if(!sheet) throw new Error('請先執行 setupMaSaLuCalendar：缺少 '+name);
  const data=sheet.getDataRange().getValues();
  if(headers.some((h,i)=>data[0][i]!==h)) throw new Error(name+' 欄位不符，已停止操作');
  return {sheet,rows:data.slice(1)};
}
function masaluDate_(value) {
  if(value instanceof Date && Number.isFinite(value.getTime())) return Utilities.formatDate(value,openBook_().getSpreadsheetTimeZone(),'yyyy-MM-dd');
  const text=String(value||'').trim();
  const match=/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text);
  if(!match) return '';
  const y=Number(match[1]),m=Number(match[2]),d=Number(match[3]);
  const test=new Date(Date.UTC(y,m-1,d));
  if(test.getUTCFullYear()!==y || test.getUTCMonth()!==m-1 || test.getUTCDate()!==d) return '';
  return match[1]+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0');
}
function masaluLoad_() {
  const employees=masaluTable_(SHIFT_APP.EMPLOYEES,SHIFT_APP.EMP_HEADERS);
  const shifts=masaluTable_(MASALU.SHIFTS,MASALU.SHIFT_HEADERS);
  const display=masaluTable_(MASALU.DISPLAY,MASALU.DISPLAY_HEADERS);
  const schedule=masaluTable_(SHIFT_APP.SHIFTS,SHIFT_APP.SHIFT_HEADERS);
  return {employees,shifts,display,schedule};
}
function masaluRevision_(data) {
  const json=JSON.stringify([data.employees.rows,data.shifts.rows,data.display.rows,data.schedule.rows]);
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,json).map(b=>(b&255).toString(16).padStart(2,'0')).join('');
}
function masaluBootstrap_(data) {
  data=data||masaluLoad_();
  const staff=data.employees.rows.filter(r=>r[0]&&r[1]).map((r,i)=>{
    const id=String(r[0]).trim(),name=String(r[1]).trim();
    const custom=data.display.rows.find(x=>String(x[0])===id)||[];
    return {id,name,short:String(custom[1]||name[0]),color:/^#[0-9a-f]{6}$/i.test(custom[2]||'')?custom[2]:MASALU.COLORS[i%MASALU.COLORS.length]};
  });
  const shifts=data.shifts.rows.filter(r=>r[0]&&r[1]).map(r=>({id:String(r[0]),name:String(r[1]),isOff:r[2]===true||String(r[2]).toLowerCase()==='true'}));
  // Existing shift names remain visible even before the manager adds configuration.
  data.schedule.rows.forEach(r=>{
    const name=String(r[6]||'').trim();
    if(name && !shifts.some(s=>s.name===name)) shifts.push({id:'legacy:'+name,name,isOff:name==='休假'});
  });
  const schedules=[];
  data.schedule.rows.forEach((r,i)=>{
    const date=masaluDate_(r[3]),shift=shifts.find(s=>s.name===String(r[6]||'').trim());
    if(!date||!shift)return;
    let staffId=String(r[2]||'').trim();
    if(!staffId && r[12]) {
      staffId='temporary:'+String(r[1]||i);
      if(!staff.some(p=>p.id===staffId)) staff.push({id:staffId,name:String(r[12]),short:String(r[12])[0],color:'#475569'});
    }
    if(staff.some(p=>p.id===staffId)) schedules.push({date,shiftId:shift.id,staffId});
  });
  return {ok:true,staff,shifts,schedules,settings:{app_name:'MASALU'},revision:masaluRevision_(data),serverTime:new Date().toISOString()};
}

/** Every calendar request uses a LINE token verified against the Login channel. */
function masaluApi(action,payload,lineAccessToken) {
  try {
    const viewer=masaluViewer_(lineAccessToken);
    if(!viewer) return {ok:false,code:'unauthorized',error:'找不到此 LINE 帳號的員工綁定，請確認使用同一個 LINE 帳號'};
    if(action==='bootstrap') return Object.assign(masaluBootstrap_(),{viewer});
    if(action==='managerSession') return viewer.manager ? {ok:true,manager:viewer} : {ok:false,code:'unauthorized',error:'只有 A01–A03 管理員可以編輯排班'};
    if(!viewer.manager) return {ok:false,code:'unauthorized',error:'只有 A01–A03 管理員可以編輯排班'};
    const lock=LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const data=masaluLoad_();
      if(!payload || payload.expectedRevision!==masaluRevision_(data)) throw new Error('資料已更新，請重新整理後再操作');
      if(action==='saveDay'||action==='clearDay') masaluSaveDay_(data,payload,action==='clearDay');
      else masaluSettings_(data,action,payload);
      SpreadsheetApp.flush();
      const result=masaluBootstrap_();
      if(action==='saveDay'||action==='clearDay') result.schedules=result.schedules.filter(r=>r.date===payload.date);
      result.viewer=viewer;
      return result;
    } finally { lock.releaseLock(); }
  } catch(err) { return {ok:false,error:String(err.message)}; }
}

function masaluProtected_(r) {
  return Boolean(r[9] && r[9]!=='未發送') || Boolean(r[7] && r[7]!=='待回覆') || Number(r[11])>0;
}
function masaluSaveDay_(data,payload,clear) {
  const date=masaluDate_(payload.date);
  if(!date || date!==payload.date) throw new Error('日期格式錯誤');
  const bootstrap=masaluBootstrap_(data),wanted=new Set();
  const assignments=clear ? {} : payload.assignments;
  if(!assignments || typeof assignments!=='object'||Array.isArray(assignments)) throw new Error('排班格式錯誤');
  Object.keys(assignments).forEach(id=>{
    const shift=bootstrap.shifts.find(s=>s.id===id);
    if(!shift||!Array.isArray(assignments[id])) throw new Error('班別已變更，請重新整理');
    assignments[id].forEach(staffId=>{
      if(!bootstrap.staff.some(p=>p.id===staffId)) throw new Error('人員不存在');
      wanted.add(JSON.stringify([staffId,shift.name]));
    });
  });
  const removals=[],existing=new Set();
  data.schedule.rows.forEach((r,i)=>{
    if(masaluDate_(r[3])!==date) return;
    const staffId=String(r[2]||'').trim() || (r[12]?'temporary:'+String(r[1]||i):'');
    const pair=JSON.stringify([staffId,String(r[6]||'').trim()]);
    if(wanted.has(pair)) {existing.add(pair);return;}
    if(masaluProtected_(r)||!staffId||!r[6]||!r[2]||!bootstrap.staff.some(p=>p.id===staffId)) throw new Error('這天包含已通知、已回覆、臨時或無法辨識的排班，請從試算表處理，原資料未變更');
    removals.push(i+2);
  });
  const additions=[];
  wanted.forEach(pair=>{
    if(existing.has(pair))return;
    const [staffId,name]=JSON.parse(pair);
    if(staffId.startsWith('temporary:')) throw new Error('臨時排班請從原試算表建立');
    additions.push([false,Utilities.getUuid(),staffId,date,'','',name,'待回覆','','未發送','',0,'','','','','月曆建立（尚未發送）']);
  });
  // Preflight above completes before the first write. Retained rows keep every notification field.
  additions.forEach(row=>data.schedule.sheet.appendRow(row));
  removals.forEach(row=>data.schedule.sheet.getRange(row,1,1,SHIFT_APP.SHIFT_HEADERS.length).clearContent());
  safeLog_('月曆儲存','', '', '成功',date+' 新增 '+additions.length+'，移除未通知 '+removals.length,'');
}

function masaluSettings_(data,action,p) {
  const name=String(p.name||'').trim();
  const shiftIndex=data.shifts.rows.findIndex(r=>String(r[0])===p.id);
  const employeeIndex=data.employees.rows.findIndex(r=>String(r[0])===p.id);
  if(action==='addShift'||action==='updateShift') {
    if(!name || name.length>60) throw new Error('班別名稱須為 1～60 字');
    if(data.shifts.rows.some(r=>r[1]===name && r[0]!==p.id)) throw new Error('班別名稱重複');
    if(action==='addShift') data.shifts.sheet.appendRow([Utilities.getUuid(),name,Boolean(p.isOff)]);
    else {
      if(shiftIndex<0) throw new Error('找不到班別');
      const oldName=data.shifts.rows[shiftIndex][1];
      if(data.schedule.rows.some(r=>r[6]===oldName)) throw new Error('此班別已有排班，為保留歷史資料請新增班別');
      data.shifts.sheet.getRange(shiftIndex+2,2).setValue(name);
    }
  } else if(action==='deleteShift') {
    if(shiftIndex<0) throw new Error('找不到班別');
    if(data.schedule.rows.some(r=>r[6]===data.shifts.rows[shiftIndex][1])) throw new Error('此班別已有排班，無法刪除');
    data.shifts.sheet.deleteRow(shiftIndex+2);
  } else if(action==='addStaff') {
    if(!name || name.length>40 || !/^#[0-9a-f]{6}$/i.test(p.color||'')) throw new Error('請確認姓名與顏色');
    const number=Math.max(0,...data.employees.rows.map(r=>/^A\d+$/.test(r[0])?Number(String(r[0]).slice(1)):0))+1;
    const id='A'+String(number).padStart(2,'0');
    data.employees.sheet.appendRow([id,name,Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase(),'','']);
    data.display.sheet.appendRow([id,String(p.short||name[0]).slice(0,1),p.color]);
  } else if(action==='deleteStaff') {
    if(employeeIndex<0) throw new Error('找不到人員');
    if(data.employees.rows[employeeIndex][3] || data.schedule.rows.some(r=>String(r[2])===p.id)) throw new Error('已綁定或已有排班的人員不能刪除');
    data.employees.sheet.deleteRow(employeeIndex+2);
    const displayIndex=data.display.rows.findIndex(r=>String(r[0])===p.id);
    if(displayIndex>=0) data.display.sheet.deleteRow(displayIndex+2);
  } else throw new Error('不支援的操作');
}

/** Verify LINE token, read the user profile, then resolve one bound employee. */
function masaluViewer_(accessToken) {
  const token=String(accessToken||'').trim();
  if(!token || token.length>4096) return null;
  const channel=PropertiesService.getScriptProperties().getProperty('MASALU_LINE_LOGIN_CHANNEL_ID') || '2010020470';
  try {
    const check=UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify?access_token='+encodeURIComponent(token),{muteHttpExceptions:true});
    if(check.getResponseCode()!==200) return null;
    const verified=JSON.parse(check.getContentText());
    if(String(verified.client_id)!==channel || Number(verified.expires_in)<=0) return null;
    const response=UrlFetchApp.fetch('https://api.line.me/v2/profile',{headers:{Authorization:'Bearer '+token},muteHttpExceptions:true});
    if(response.getResponseCode()!==200) return null;
    const profile=JSON.parse(response.getContentText());
    if(!validLineUid_(String(profile.userId||''))) return null;
    const matches=masaluTable_(SHIFT_APP.EMPLOYEES,SHIFT_APP.EMP_HEADERS).rows.filter(r=>String(r[3]||'').trim()===profile.userId);
    if(matches.length!==1) return null;
    const id=String(matches[0][0]||'').trim().toUpperCase();
    return {id:id,name:String(matches[0][1]||''),manager:['A01','A02','A03'].includes(id)};
  } catch(err) { return null; }
}

function masaluManager_(accessToken) {
  const viewer=masaluViewer_(accessToken);
  return viewer&&viewer.manager ? viewer : null;
}

