const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const crypto=require('node:crypto');
const baseRows=[
  [false,'11111111-1111-4111-8111-111111111111','A03','2026-11-01','','','玉門','不同意','', '已發送','',1,'','','','',''],
  [false,'22222222-2222-4222-8222-222222222222','A10','2026-11-01','','','玉門','同意','', '已發送','',1,'','','','',''],
];
function fixture(){
  const rows=baseRows.map(r=>r.slice());
  let uuid=0;const logs=[];
  const sheet={
    appendRow(row){rows.push(row.slice());},
    getRange(row,col){return {
      setValue(v){rows[row-2][col-1]=v;},
      setValues(v){v[0].forEach((x,i)=>rows[row-2][col+i-1]=x);}
    };}
  };
  const globals={
    SHIFT_APP:{EMPLOYEES:'員工名單',SHIFTS:'排班表',EMP_HEADERS:[],SHIFT_HEADERS:Array(17).fill('x')},
    Utilities:{getUuid(){return `00000000-0000-4000-8000-${String(++uuid).padStart(12,'0')}`;},
      DigestAlgorithm:{SHA_256:'sha256'},computeDigest(_alg,json){return [...crypto.createHash('sha256').update(json).digest()].map(x=>x>127?x-256:x);}},
    SpreadsheetApp:{flush(){}},
    safeLog_(...args){logs.push(args);},
    console,Date,JSON,Set,Object,Array,String,Number,Boolean,RegExp
  };
  const ctx=vm.createContext(globals);
  vm.runInContext(fs.readFileSync(__dirname+'/CalendarApi.gs','utf8'),ctx);
  const data={employees:{rows:[['A03','米馳浩'],['A10','米測浩'],['A05','勇翰']]},
    shifts:{rows:[['S1','玉門',false]]},display:{rows:[]},schedule:{rows,sheet}};
  return {ctx,data,rows,logs};
}
function save(f,staff){return f.ctx.masaluSaveDay_(f.data,{date:'2026-11-01',assignments:{S1:staff}},false);}
{
  const f=fixture();
  let view=f.ctx.masaluBootstrap_(f.data,{id:'A10',manager:false});
  assert.deepEqual(Array.from(view.schedules,r=>r.staffId),['A10']);
  assert.deepEqual(Array.from(view.staff,r=>r.id),['A10']);
  assert.equal(view.schedules[0].replyStatus,'同意');
  const changes=save(f,['A10','A05']);
  assert.equal(changes.added,1);assert.equal(changes.cancelled,0);
  assert.equal(f.rows[0][7],'不同意');
  assert.equal(f.rows[1][7],'同意');
  assert.equal(f.rows[2][2],'A05');
  view=f.ctx.masaluBootstrap_(f.data,{id:'A03',manager:false});
  assert.equal(view.schedules.length,0,'declined day is blank for A03');
}
{
  const f=fixture();
  const changes=save(f,['A05']);
  assert.equal(changes.added,1);assert.equal(changes.cancelled,1);
  assert.equal(f.rows[1][7],'同意','retain reply history');
  assert.equal(f.rows[1][9],'已取消');
  const view=f.ctx.masaluBootstrap_(f.data,{id:'A01',manager:true});
  assert.deepEqual(Array.from(view.schedules,r=>r.staffId),['A05']);
  assert.equal(save(f,['A05']).cancelled,0,'do not repeat cancel');
  assert.equal(f.rows.filter(r=>r[2]==='A05').length,1,'do not duplicate new shift');
}
{
  const f=fixture();
  assert.throws(()=>save(f,['INVALID']),/人員不存在/);
  assert.equal(f.rows.length,2,'validation must not write');
}
{
  const html=fs.readFileSync(__dirname+'/index.html','utf8');
  assert(!html.includes("prompt('請輸入管理金鑰"));
  assert(html.includes('liff.getAccessToken()'));
  assert(html.includes('state.viewer?.manager'));
  assert(!html.includes("return new Promise((resolve,reject)=>{"));
}
{
  const f=fixture();
  const lineUid='U'+'a'.repeat(32);
  f.ctx.PropertiesService={getScriptProperties(){return {getProperty(){return null;}}}};
  f.ctx.validLineUid_=uid=>/^U[0-9a-f]{32}$/i.test(uid);
  f.ctx.UrlFetchApp={fetch(url){return {
    getResponseCode(){return 200;},
    getContentText(){return JSON.stringify(url.includes('/verify?') ? {client_id:'2010020470',expires_in:3600} : {userId:lineUid});}
  };}};
  f.ctx.masaluTable_=()=>({rows:[['A03','米馳浩','','U'+'a'.repeat(32)],['A05','勇翰','','U'+'b'.repeat(32)]]});
  const manager=f.ctx.masaluViewer_('line-token');
  assert.equal(manager.id,'A03');assert.equal(manager.manager,true);
  f.ctx.masaluTable_=()=>({rows:[['A05','勇翰','',lineUid]]});
  const staff=f.ctx.masaluViewer_('line-token');
  assert.equal(staff.id,'A05');assert.equal(staff.manager,false);
  f.ctx.UrlFetchApp.fetch=()=>({getResponseCode(){return 200;},getContentText(){return JSON.stringify({client_id:'wrong-channel',expires_in:3600});}});
  assert.equal(f.ctx.masaluViewer_('line-token'),null,'wrong channel denied');
}
(async()=>{
  const script=fs.readFileSync(__dirname+'/worker-calendar-api.js','utf8');
  const worker=(await import('data:text/javascript;base64,'+Buffer.from(script).toString('base64'))).default;
  const origin='https://enddenmi128.github.io';
  const endpoint='https://example.workers.dev/calendar-api';
  const env={GAS_SHARED_SECRET:'server-only-secret',GAS_WEB_APP_URL:'https://script.google.com/macros/s/test/exec'};
  let forwarded;
  const originalFetch=global.fetch;
  global.fetch=async(_url,options)=>{forwarded=JSON.parse(options.body);return new Response(JSON.stringify({ok:true,viewer:{id:'A03',manager:true}}));};
  try{
    const make=(action,token,site=origin)=>new Request(endpoint,{method:'POST',headers:{origin:site,authorization:token?'Bearer '+token:'','content-type':'application/json'},body:JSON.stringify({action,payload:{}})});
    assert.equal((await worker.fetch(make('bootstrap',null),env)).status,401);
    assert.equal((await worker.fetch(make('bootstrap','x','https://attacker.invalid'),env)).status,403);
    assert.equal((await worker.fetch(make('deleteAll','x'),env)).status,400);
    const out=await worker.fetch(make('bootstrap','valid-line-token'),env);
    assert.equal(out.status,200);
    assert.equal(forwarded.lineAccessToken,'valid-line-token');
    assert.equal(forwarded.kind,'calendar_api');
    assert.equal(forwarded.secret,env.GAS_SHARED_SECRET);
  }finally{global.fetch=originalFetch;}
  console.log('PASS: schedule retention, personal privacy, LINE identity/roles, Worker origin and token forwarding');
})().catch(err=>{console.error(err);process.exitCode=1;});
