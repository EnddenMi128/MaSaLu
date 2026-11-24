import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFile} from 'node:fs/promises';

const menu=await readFile(new URL('../GAS/CalendarRichMenu.js',import.meta.url),'utf8');
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
  };
  vm.createContext(ctx);vm.runInContext(menu,ctx);
  return ctx.masaluPendingCalendarReplyButtons_(uid,page);
}
test('pending shifts render two versioned choices per assignment',()=>{
  const buttons=setup([makeRow(),makeRow({id:second,date:'2026-11-27',shift:'田園',version:3})]);
  assert.equal(buttons.length,4);
  assert.deepEqual(Array.from(buttons,b=>b.action.type),['postback','postback','postback','postback']);
  assert.deepEqual(Array.from(buttons,b=>b.action.label),['11/26 玉門 同意','11/26 玉門 不同意','11/27 田園 同意','11/27 田園 不同意']);
  for(const b of buttons){
    const qs=new URLSearchParams(b.action.data);
    assert.equal(qs.get('action'),'shift_reply');
    assert.ok([first,second].includes(qs.get('id')));
    assert.ok(['agree','disagree'].includes(qs.get('choice')));
    assert.ok([1,3].includes(Number(qs.get('v'))));
  }
});
test('other users and completed, expired or unversioned assignments cannot be replied to',()=>{
  const buttons=setup([
    makeRow({who:'A11'}),makeRow({id:second,status:'同意'}),
    makeRow({id:second,status:'不同意'}),
    makeRow({id:second,status:'已請假'}),
    makeRow({id:second,version:0}),
    makeRow({id:'bad'}),makeRow({id:second,date:'2026-12-06'}),
    makeRow({id:first})
  ]);
  assert.equal(buttons.length,2);
});
test('reply buttons are capped to 12 to leave a 13th slot for leave',()=>{
  const rows=Array.from({length:9},(_,i)=>makeRow({id:'550e8400-e29b-41d4-a716-'+String(446655440010+i),date:'2026-11-26'}));
  assert.equal(setup(rows).length,12);
});
test('ambiguous roster link prevents generating identity-specific actions',()=>{
  const b=setup([makeRow()],{A10:{uid},A12:{uid}});
  assert.equal(b.length,0);
});
test('quick reply postbacks have an existing Worker handler and secure GAS route',()=>{
  assert.match(worker,/function parsePostbackData\(data\)/);
  assert.match(worker,/qs\.get\("action"\) !== "shift_reply"/);
  assert.match(worker,/kind = parsed \? "shift_reply"/);
  assert.match(gas,/input\.kind === 'shift_reply'/);
  assert.match(gas,/expectedUid === uid/);
});
