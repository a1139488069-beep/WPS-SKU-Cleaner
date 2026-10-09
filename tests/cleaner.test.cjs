'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../js/SkuCleaner.js'), 'utf8');
function load() {
  const listeners = {}, alerts = [], timers = new Map(); let timerId = 0;
  const app = {EnableEvents:true,ScreenUpdating:true,WorksheetFunction:{IsError:c=>!!c.isError},
    ApiEvent:{AddApiEventListener:(name, fn)=>{(listeners[name] ||= []).push(fn)}}};
  const ctx = vm.createContext({Application:app,alert:x=>alerts.push(x),
    setTimeout:fn=>{timers.set(++timerId,fn);return timerId},clearTimeout:id=>timers.delete(id)});
  vm.runInContext(source, ctx);
  function book(name='test.xlsx') {
    const b = {Name:name,FullName:'C:/test/'+name,sheets:[]};
    b.Worksheets = {get Count(){return b.sheets.length},Item:i=>b.sheets[i-1]};
    return b;
  }
  function sheet(b, name='Data', startRow=1, startCol=1) {
    const cells = new Map();
    let lastRow=startRow,lastCol=startCol;
    const sh = {Name:name,Parent:b,ProtectContents:false,Cells:{Item:(r,c)=>{
      const key=r+','+c;
      if(!cells.has(key)) cells.set(key,{Value2:null,NumberFormat:'General',HasFormula:false,MergeCells:false,
        ClearContents(){this.Value2=null}});
      return cells.get(key);
    }}};
    sh.put = (r,c,v,props={})=>{
      lastRow=Math.max(lastRow,r);lastCol=Math.max(lastCol,c);
      Object.assign(sh.Cells.Item(r,c),{Value2:v},props);return sh.Cells.Item(r,c);
    };
    Object.defineProperty(sh,'UsedRange',{get:()=>({Row:startRow,Column:startCol,
      Rows:{Count:lastRow-startRow+1},Columns:{Count:lastCol-startCol+1}})});
    b.sheets.push(sh);app.ActiveWorkbook=b;app.ActiveSheet=sh;return sh;
  }
  function target(...boxes) {
    return {Areas:{Count:boxes.length,Item:i=>{
      const [r,c,rows=1,cols=1]=boxes[i-1];return {Row:r,Column:c,Rows:{Count:rows},Columns:{Count:cols}};
    }}};
  }
  function flush(){while(timers.size){const [id,fn]=timers.entries().next().value;timers.delete(id);fn()}}
  return {ctx,app,listeners,alerts,book,sheet,target,flush};
}
const cases = [
  ["'2714230032211",'convert',2714230032211],
  [' \t2714\u00a0230032211\u200b\ufeff\n','convert',2714230032211],
  ['２７１４２３００３２２１１','convert',2714230032211],
  ['000123','convert',123],['0000000000000000123','convert',123],['000','convert',0],
  ['999999999999999','convert',999999999999999],['1000000000000000','long'],
  ['ABC-123','text','ABC-123'],['SKU001','text','SKU001'],['12.0','text','12.0'],['1e12','text','1e12'],
  ['1,234','text','1,234'],['-123','text','-123'],['+123','text','+123'],['１２Ａ','text','12Ａ'],
  [true,'invalid'],[-1,'invalid'],[1.5,'invalid'],[Infinity,'invalid'],[NaN,'invalid'],
  [0,'numeric',0],[2714230032211,'numeric',2714230032211],[1000000000000000,'long'],
  [null,'blank'],['','blank'],[' \t\u200b','blankText'],['’00123','convert',123],
];
for(const [input,kind,value] of cases) test('classification '+String(input)+' => '+kind,()=>{
  const {ctx}=load();const out=ctx.skuClassify(input);assert.equal(out.kind,kind);
  if(value!==undefined) assert.equal(out.value,value);
});
test('mixed case/contained headers; other column and formulas preserved; format is exactly 0',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);
  s.put(1,1,'商品 sKu');s.put(1,2,'名称');s.put(1,3,'SKU');
  const good=s.put(2,1,"'2714230032211",{NumberFormat:'@'}),other=s.put(2,2,'000123');
  const formula=s.put(3,1,456,{HasFormula:true,Formula:'=456',NumberFormat:'@'});
  const bad=s.put(4,1,'ABC-123',{NumberFormat:'@'}),long=s.put(5,1,'1000000000000001',{NumberFormat:'@'});
  const merged=s.put(6,1,'123',{MergeCells:true}),err=s.put(7,1,2042,{isError:true});
  s.put(2,3,'００１２３');const whitespace=s.put(8,1,'  \u200b');
  const stats=x.ctx.skuExecute(b,s,null);
  assert.equal(stats.columns,2);assert.equal(stats.converted,2);assert.equal(good.Value2,2714230032211);
  assert.equal(good.NumberFormat,'0');assert.equal(s.Cells.Item(2,3).Value2,123);
  assert.equal(other.Value2,'000123');assert.equal(other.NumberFormat,'General');
  assert.equal(formula.Formula,'=456');assert.equal(formula.NumberFormat,'@');
  assert.equal(bad.Value2,'ABC-123');assert.equal(bad.NumberFormat,'0');assert.equal(stats.textFormatted,1);assert.equal(long.Value2,'1000000000000001');
  assert.equal(merged.Value2,'123');assert.equal(err.Value2,2042);assert.equal(err.NumberFormat,'General');
  assert.equal(whitespace.Value2,null);assert.equal(whitespace.NumberFormat,'0');
});
test('headers offset from A1, configurable header row and exact match',()=>{
  const x=load(),b=x.book(),s=x.sheet(b,'Data',3,4);s.put(4,5,'SKU');s.put(5,5,'0005');
  x.ctx.skuExecute(b,s,null);assert.equal(s.Cells.Item(5,5).Value2,5);
  const far=x.sheet(b,'Far',25,4);far.put(25,4,'SKU');far.put(26,4,'0006');
  x.ctx.skuExecute(b,far,null);assert.equal(far.Cells.Item(26,4).Value2,'0006');
  x.ctx.SKU_CONFIG.headerRow=25;x.ctx.skuExecute(b,far,null);assert.equal(far.Cells.Item(26,4).Value2,6);
  x.ctx.SKU_CONFIG.headerContains=false;assert.equal(x.ctx.skuIsHeader('商品SKU'),false);
  assert.equal(x.ctx.skuIsHeader(' sku '),true);
});
test('process whole workbook, skip protected sheets and sheets without SKU',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');s.put(2,1,'12');
  const p=x.sheet(b,'Protected');p.put(1,1,'sku');p.put(2,1,'13');p.ProtectContents=true;
  const n=x.sheet(b,'Other');n.put(1,1,'ID');n.put(2,1,'14');
  const stats=x.ctx.skuExecute(b,null,null);assert.equal(s.Cells.Item(2,1).Value2,12);
  assert.equal(p.Cells.Item(2,1).Value2,'13');assert.equal(n.Cells.Item(2,1).Value2,'14');assert.equal(stats.protectedSheets,1);
});
test('paste handler only touches changed SKU rows; duplicate target areas process once',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'sku');s.put(2,1,'001');s.put(3,1,'002');s.put(4,1,'003');
  const stats=x.ctx.skuExecute(b,s,x.target([3,1,1,2],[3,1]));
  assert.equal(stats.converted,1);assert.equal(s.Cells.Item(2,1).Value2,'001');
  assert.equal(s.Cells.Item(3,1).Value2,2);assert.equal(s.Cells.Item(4,1).Value2,'003');
  assert.equal(x.ctx.skuExecute(b,s,x.target([2,2])).columns,0);
});
test('pasting only a header avoids a full-sheet rescan',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');s.put(2,1,'001');s.put(3,1,'002');
  x.ctx.skuExecute(b,s,x.target([1,1]));assert.equal(s.Cells.Item(3,1).Value2,'002');
});
test('auto enable/disable, one registration, bound workbook, rename and close',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');s.put(2,1,'001');
  x.ctx.SKU_EnableAuto();x.ctx.SKU_EnableAuto();assert.equal(x.listeners.SheetChange.length,1);
  s.put(3,1,'002');x.listeners.SheetChange[0](s,x.target([3,1]));x.flush();assert.equal(s.Cells.Item(3,1).Value2,2);
  const other=x.book('other.xlsx'),o=x.sheet(other);o.put(1,1,'sku');o.put(2,1,'003');
  x.listeners.SheetChange[0](o,x.target([2,1]));assert.equal(o.Cells.Item(2,1).Value2,'003');
  b.FullName='C:/saved/newname.xlsx';s.put(4,1,'004');x.listeners.SheetChange[0](s,x.target([4,1]));
  x.flush();assert.equal(s.Cells.Item(4,1).Value2,4);
  x.ctx.SKU_DisableAuto();s.put(5,1,'005');x.listeners.SheetChange[0](s,x.target([5,1]));assert.equal(s.Cells.Item(5,1).Value2,'005');
  x.listeners.WorkbookBeforeClose[0](b);assert.equal(x.ctx.SKU_RUN.book,null);
});
test('global state survives macro re-evaluation',()=>{
  const x=load();x.ctx.SKU_STATE.enabled=true;vm.runInContext(source,x.ctx);assert.equal(x.ctx.SKU_STATE.enabled,true);
});
test('restore original event/screen settings even after sheet API failure',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);Object.defineProperty(s,'ProtectContents',{get(){throw Error('API failed')}});
  x.app.EnableEvents=false;x.app.ScreenUpdating=false;
  assert.throws(()=>x.ctx.skuExecute(b,s,null),/API failed/);
  assert.equal(x.app.EnableEvents,false);assert.equal(x.app.ScreenUpdating,false);assert.equal(x.ctx.SKU_STATE.busy,false);
});
test('write failure reported while other cells continue and recursion is guarded',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');const failed=s.put(2,1,'123');s.put(3,1,'456');
  Object.defineProperty(failed,'NumberFormat',{set(){throw Error('locked cell')}});
  const stats=x.ctx.skuExecute(b,s,null);assert.equal(stats.failed,1);assert.equal(s.Cells.Item(3,1).Value2,456);
  assert.equal(x.app.EnableEvents,true);assert.equal(x.app.ScreenUpdating,true);
  x.ctx.SKU_STATE.busy=true;assert.equal(x.ctx.skuExecute(b,s,null),null);
});
test('large paste bounded by used range rather than full worksheet rows',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');s.put(2,1,'123');s.put(3,1,'456');
  const stats=x.ctx.skuExecute(b,s,x.target([1,1,1048576]));assert.equal(stats.converted,2);assert.equal(stats.blank,0);
});
test('pure letters and mixed SKU clean whitespace/apostrophes but preserve case, symbols and text type',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');
  const samples=[[" 'Ab C\t\u200b",'AbC'],["' AB12-CD０３ ",'AB12-CD03'],[' 001AB ','001AB'],['ABC','ABC']];
  samples.forEach(([v],i)=>s.put(i+2,1,v,{NumberFormat:'@'}));
  const stats=x.ctx.skuExecute(b,s,null);assert.equal(stats.textFormatted,samples.length);
  samples.forEach(([,expected],i)=>{const c=s.Cells.Item(i+2,1);assert.equal(c.Value2,expected);assert.equal(c.NumberFormat,'0');assert.equal(typeof c.Value2,'string');});
  x.ctx.skuExecute(b,s,null);
  samples.forEach(([,expected],i)=>assert.equal(s.Cells.Item(i+2,1).Value2,expected));
});
test('text-like numbers and formula-like identifiers are written under text format before setting 0',()=>{
  const x=load(),b=x.book(),s=x.sheet(b);s.put(1,1,'SKU');
  const inputs=['1E12','12.0','=1+2','1/2','+123','-123'];
  inputs.forEach((input,i)=>{
    const cell=s.put(i+2,1,input);let value=input;
    Object.defineProperty(cell,'Value2',{get:()=>value,set:v=>{
      if(cell.NumberFormat!=='@') throw Error('text write would be interpreted');value=v;
    }});
  });
  const stats=x.ctx.skuExecute(b,s,null);assert.equal(stats.failed,0);assert.equal(stats.textFormatted,inputs.length);
  inputs.forEach((input,i)=>{const c=s.Cells.Item(i+2,1);assert.equal(c.Value2,input);assert.equal(c.NumberFormat,'0');assert.equal(c.HasFormula,false);});
});
