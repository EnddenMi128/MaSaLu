import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';

const menu=await readFile(new URL('../GAS/CalendarRichMenu.js',import.meta.url),'utf8');
const actions=await readFile(new URL('../GAS/CalendarMenuActions.js',import.meta.url),'utf8');
const worker=await readFile(new URL('../cloudflare/workers.js',import.meta.url),'utf8');
const gas=await readFile(new URL('../GAS/程式碼.js',import.meta.url),'utf8');
const uid='U'+'a'.repeat(32), another='U'+'b'.repeat(32);
const first='550e8400-e29b-41d4-a716-446655440000';
const second='550e8400-e29b-41d4-a716-446655440001';
const page={start:'2026-11-22',end:'2026-12-05'};
const makeRow=({id=first,who='A10',date='2026-11-26',status='待回覆',version=1,shift='玉門'}={})=>{
  const r=Array(17).fill('');
  r[1]=id;r[2]=who;r[3]=date;r[6]=shift;r[7]=status;r[11]=version;return r;
};
function setup(rows,roster={A10:{uid},A11:{uid:another}}){
  const ctx={
    employeeIndex_:()=>roster,
    shiftSheet_:()=>({getLastRow:()=>rows.length+1,getRange:()=>({getValues:()=>rows})}),
    SHIFT_APP:{SHIFT_HEADERS:Array(17).fill('')},
    masaluDate_:date=>date,
    validLineUid_:v=>/^U[a-f0-9]{32}$/.test(v),
    openBook_:()=>({getSpreadsheetTimeZone:()=> 'Asia/Taipei'}),
    Utilities:{formatDate:()=> '2026-11-27'},
  };
  vm.createContext(ctx);vm.runInContext(menu+'\n'+actions,ctx);
  return ctx;
}
test('first step selects a date and location only; second step offers versioned decisions',()=>{
  const ctx=setup([makeRow(),makeRow({id:second,date:'2026-11-27',shift:'田園',version:3})]);
  const buttons=ctx.masaluPendingCalendarReplyButtons_(uid,page);
  assert.deepEqual(Array.from(buttons,b=>b.action.label),['11/26 玉門','11/27 田園']);
  const qs=new URLSearchParams(buttons[1].action.data);
  assert.equal(qs.get('action'),'calendar_menu');
  assert.equal(qs.get('choice'),'reply_select');
  assert.equal(qs.get('value'),second+':3');
  const result=ctx.masaluCalendarMenuAction_({lineUid:uid,choice:'reply_select',value:qs.get('value')});
  assert.match(result.message,/11\/27/);
  assert.match(result.message,/田園/);
  const items=result.quickReply.items;
  assert.deepEqual(Array.from(items,b=>b.action.label),['同意','不同意','返回']);
  for(let i=0;i<2;i++){
    const q=new URLSearchParams(items[i].action.data);
    assert.equal(q.get('action'),'shift_reply');
    assert.equal(q.get('id'),second);
    assert.equal(q.get('v'),'3');
    assert.equal(q.get('choice'),i===0?'agree':'disagree');
  }
  const back=ctx.masaluCalendarMenuAction_({lineUid:uid,choice:'reply_pick',value:0});
  assert.equal(back.quickReply.items.length,3);
});
test('other users and completed or unversioned assignments cannot be selected',()=>{
  const ctx=setup([
    makeRow({who:'A11',id:second}),makeRow({id:second,status:'同意'}),
    makeRow({id:second,status:'不同意'}),makeRow({id:second,status:'已請假'}),
    makeRow({id:second,version:0}),makeRow({id:'bad'}),
    makeRow({id:second,date:'2026-12-06'}),makeRow({id:first})
  ]);
  assert.equal(ctx.masaluPendingCalendarReplyButtons_(uid,page).length,1);
  for(const value of [second+':1',first+':2','invalid']){
    const result=ctx.masaluCalendarMenuAction_({lineUid:uid,choice:'reply_select',value});
    assert.equal(result.code,'expired');
    assert.equal(result.quickReply.items.length,1);
  }
});
test('pagination preserves room for leave and reaches every pending assignment',()=>{
  const rows=Array.from({length:24},(_,i)=>makeRow({id:'550e8400-e29b-41d4-a716-'+String(446655440010+i),date:'2026-11-26'}));
  const ctx=setup(rows);
  const all=[];
  for(let i=0;i<3;i++){
    const buttons=ctx.masaluCalendarMenuAction_({lineUid:uid,choice:'reply_pick',value:i}).quickReply.items;
    assert.ok(buttons.length<=13);
    for(const b of buttons){
      const qs=new URLSearchParams(b.action.data);
      if(qs.get('choice')==='reply_select')all.push(qs.get('value'));
    }
  }
  assert.equal(new Set(all).size,24);
});
test('ambiguous roster link prevents generating identity-specific actions',()=>{
  const ctx=setup([makeRow()],{A10:{uid},A12:{uid}});
  assert.equal(ctx.masaluPendingCalendarReplyButtons_(uid,page).length,0);
});
test('quick reply postbacks have an existing Worker handler and secure GAS route',()=>{
  assert.match(worker,/function parsePostbackData\(data\)/);
  assert.match(worker,/qs\.get\("action"\) !== "shift_reply"/);
  assert.match(worker,/kind = parsed \? "shift_reply"/);
  assert.match(gas,/input\.kind === 'shift_reply'/);
  assert.match(gas,/expectedUid === uid/);
});
