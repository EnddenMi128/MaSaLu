/**
 * MaSaLu calendar-only Cloudflare Worker.
 * Deploy as a SEPARATE Worker (do not replace the existing LINE webhook Worker).
 * Secrets: GAS_SHARED_SECRET. Variable: GAS_WEB_APP_URL (existing deployed /exec URL).
 * The LINE access token is validated and mapped to A01-A03 by GAS, never by this browser.
 */
const ORIGIN='https://enddenmi128.github.io';
const ACTIONS=new Set(['bootstrap','managerSession','saveDay','clearDay',
  'addShift','updateShift','deleteShift','addStaff','deleteStaff']);

function cors(origin){
  return {'access-control-allow-origin':origin,'access-control-allow-methods':'POST, OPTIONS',
    'access-control-allow-headers':'Authorization, Content-Type',
    'access-control-max-age':'600','vary':'Origin','cache-control':'no-store'};
}
function reply(value,status,origin){
  return new Response(JSON.stringify(value),{status,headers:{'content-type':'application/json; charset=utf-8',...cors(origin)}});
}

export default {
  async fetch(request,env){
    const url=new URL(request.url);
    const origin=request.headers.get('origin')||'';
    if(url.pathname!=='/calendar-api')return new Response('Not found',{status:404});
    if(origin!==ORIGIN)return new Response('Forbidden origin',{status:403});
    if(request.method==='OPTIONS')return new Response(null,{status:204,headers:cors(origin)});
    if(request.method!=='POST')return reply({ok:false,error:'僅支援 POST'},405,origin);
    const match=/^Bearer\s+(.+)$/i.exec(request.headers.get('authorization')||'');
    if(!match||match[1].length>4096)return reply({ok:false,code:'unauthorized',error:'請從 LINE 開啟排班月曆'},401,origin);
    const secret=env.GAS_SHARED_SECRET;
    const gasUrl=env.GAS_WEB_APP_URL;
    if(!secret||!gasUrl)return reply({ok:false,error:'Worker 尚未完成 GAS 環境設定'},503,origin);
    let body;
    try{
      const raw=await request.text();
      if(raw.length>30000)throw new Error('請求過大');
      body=JSON.parse(raw);
    }catch(_){return reply({ok:false,error:'請求格式錯誤'},400,origin);}
    const action=String(body?.action||'');
    if(!ACTIONS.has(action))return reply({ok:false,error:'不支援的操作'},400,origin);
    try{
      const destination=new URL(gasUrl);
      if(destination.protocol!=='https:'||destination.hostname!=='script.google.com'||!destination.pathname.endsWith('/exec'))throw Error('GAS URL 不正確');
      const upstream=await fetch(destination.toString(),{
        method:'POST',redirect:'follow',
        headers:{'content-type':'application/json'},
        body:JSON.stringify({kind:'calendar_api',secret,action,
          payload:body.payload&&typeof body.payload==='object'?body.payload:{},
          lineAccessToken:match[1]})
      });
      if(!upstream.ok)return reply({ok:false,error:'GAS 服務暫時無法連線'},502,origin);
      const result=await upstream.json();
      return reply(result,200,origin);
    }catch(_){return reply({ok:false,error:'月曆服務連線失敗，請稍後重試'},502,origin);}
  }
};

