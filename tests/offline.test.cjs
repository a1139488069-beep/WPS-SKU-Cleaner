'use strict';
const test=require('node:test'),assert=require('node:assert/strict');
const fs=require('node:fs'),path=require('node:path');
const {host,root,pkg}=require('./fixtures/host.cjs');
function sample(x){const b=x.makeBook(),s=b.sh;s.Cells.Item(1,3).Value2='公司sku';s.Cells.Item(1,4).Value2='平台sku';s.Cells.Item(1,5).Value2='主状态';s.Cells.Item(2,3).Value2="'2714230032211";s.Cells.Item(2,4).Value2='2610230099311';s.Cells.Item(2,5).Value2='000123';return b}
test('offline entrypoint and generated bundle match',()=>{
  const dir=path.join(root,'offline','WPS-SKU-Cleaner_'+pkg.version);
  assert.ok(fs.readFileSync(path.join(dir,'index.html'),'utf8').includes('src="./main.js"'));
  assert.equal(fs.readFileSync(path.join(dir,'main.js'),'utf8'),fs.readFileSync(path.join(root,'addin/main.js'),'utf8'));
  assert.equal(fs.readFileSync(path.join(dir,'ribbon.xml'),'utf8'),fs.readFileSync(path.join(root,'addin/ribbon.xml'),'utf8'));
});
test('loading is idle even with stored auto=true and million-row formatted UsedRange',()=>{
  const x=host(),b=sample(x);b.sh.inflate(1048576,16384);x.metrics.reads=0;x.metrics.writes=0;
  x.ctx.SkuAddinLoad({});x.ctx.SkuAddinLoad({});
  assert.equal(x.ctx.SKU_RUN.enabled,false);assert.equal(x.metrics.reads,0);assert.equal(x.metrics.writes,0);assert.equal(x.timers.size,0);
  assert.deepEqual(Object.keys(x.events).sort(),['SheetChange','WorkbookBeforeClose']);assert.equal(x.registrations.SheetChange,1);
});
test('manual processing is deferred, recognizes both SKU columns and preserves other columns',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SKU_CleanCurrentSheet();
  assert.equal(b.sh.Cells.Item(2,3).Value2,"'2714230032211");x.flush();
  for(const c of [3,4]){assert.equal(typeof b.sh.Cells.Item(2,c).Value2,'number');assert.equal(b.sh.Cells.Item(2,c).NumberFormat,'0')}
  assert.equal(b.sh.Cells.Item(2,5).Value2,'000123');assert.ok(x.alerts.at(-1).startsWith('SKU处理完成'));
});
test('enable does not scan; event callback only queues, and only the bound book is handled',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();
  assert.equal(x.timers.size,0);assert.equal(b.sh.Cells.Item(2,3).Value2,"'2714230032211");
  const other=sample(x);x.events.SheetChange(other.sh,x.target(2,3));assert.equal(x.timers.size,0);
  b.FullName='C:/renamed.xlsx';x.metrics.reads=0;x.metrics.writes=0;
  x.events.SheetChange(b.sh,x.target(2,3));assert.equal(x.metrics.reads,0);assert.equal(x.metrics.writes,0);
  x.flush();assert.equal(b.sh.Cells.Item(2,3).Value2,2714230032211);assert.equal(b.sh.Cells.Item(2,4).Value2,'2610230099311');
});
test('pure letters/mixed/formula-looking strings retain text and get format 0',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';
  [' AB C０３ ','1E12',' =1+2 ','+123','12.0'].forEach((v,i)=>s.Cells.Item(i+2,1).Value2=v);
  x.ctx.SKU_CleanCurrentSheet();x.flush();
  ['ABC03','1E12','=1+2','+123','12.0'].forEach((v,i)=>{const c=s.Cells.Item(i+2,1);assert.equal(c.Value2,v);assert.equal(c.NumberFormat,'0');assert.equal(c.HasFormula,false)});
});
test('reprocessing normalized data performs zero writes, including delayed own events',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();x.flush();
  x.metrics.writes=0;for(let i=0;i<100;i++)x.events.SheetChange(b.sh,x.target(2,3));
  assert.equal(x.ctx.SKU_RUN.pending.length,1);assert.equal(x.ctx.SKU_RUN.pending[0].bounds.length,1);x.flush();
  assert.equal(x.metrics.writes,0);assert.equal(x.ctx.SKU_RUN.enabled,true);
});
test('synchronous write feedback is guarded and queued tasks terminate',()=>{
  let x;x=host({onWrite:(s,r,c)=>{if(x && x.events.SheetChange)x.events.SheetChange(s,x.target(r,c))}});
  const b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();x.flush();
  assert.equal(x.ctx.SKU_RUN.pending.length,0);assert.equal(x.timers.size,0);assert.equal(x.ctx.SKU_RUN.enabled,true);
});
test('whole-column paste ignores formatted empty tail and never formats empty cells',()=>{
  const x=host(),b=sample(x);b.sh.inflate(1048576,5);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();
  x.events.SheetChange(b.sh,x.target(1,3,1048576));x.flush();
  assert.equal(x.ctx.SKU_STATE.last.visited,1);assert.equal(b.sh.Cells.Item(3,3).NumberFormat,'@');assert.equal(x.metrics.finds,1);
});
test('header-only changes do not scan existing data',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.events.SheetChange(b.sh,x.target(1,3));x.flush();
  assert.equal(b.sh.Cells.Item(2,3).Value2,"'2714230032211");assert.equal(x.ctx.SKU_STATE.last.visited,0);
});
test('100000 rows run in bounded slices and cancellation restores host flags',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';
  for(let r=2;r<=100001;r++)s.Cells.Item(r,1).Value2='ABC'+r;
  x.ctx.Date={now:()=>0};
  x.metrics.writes=0;x.ctx.SKU_CleanCurrentSheet();x.flush();
  assert.equal(x.ctx.SKU_STATE.last.visited,100000);assert.equal(x.metrics.writes,100);
  assert.ok(x.metrics.maxVisited<=1000);assert.ok(x.metrics.ticks<=110);
  assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);
  x.metrics.writes=0;x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.metrics.writes,0);
  x.ctx.SKU_CleanCurrentSheet();x.tick();x.ctx.SkuAddinStop();assert.equal(x.timers.size,0);assert.equal(x.ctx.SKU_RUN.current,null);
});
test('200000-cell limit rejects oversized plan before sheet writes',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';s.Cells.Item(200002,1).Value2='001';x.metrics.writes=0;
  x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.metrics.writes,0);assert.match(x.ctx.SKU_RUN.error,/200000/);assert.equal(x.ctx.SKU_RUN.enabled,false);
});
test('small selected region can be processed far down a large sheet',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';s.Cells.Item(900000,1).Value2='001';
  x.app.Selection=x.target(900000,1);x.ctx.SKU_CleanSelection();x.flush();assert.equal(s.Cells.Item(900000,1).Value2,1);assert.equal(x.ctx.SKU_STATE.last.visited,1);
});
test('unsupported Find never falls back to a huge scan',()=>{
  const x=host({noFind:true}),b=sample(x);b.sh.inflate(1048576,5);x.metrics.writes=0;x.ctx.SKU_CleanCurrentSheet();x.flush();
  assert.equal(x.metrics.writes,0);assert.match(x.ctx.SKU_RUN.error,/20000/);
  x.app.Selection=x.target(2,3);x.ctx.SKU_CleanSelection();x.flush();assert.equal(b.sh.Cells.Item(2,3).Value2,2714230032211);
});
test('closing bound/running workbook cancels all work before further reads/writes',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();
  x.events.WorkbookBeforeClose(b);x.metrics.reads=0;x.metrics.writes=0;x.flush();assert.equal(x.metrics.reads,0);assert.equal(x.metrics.writes,0);assert.equal(x.ctx.SKU_RUN.book,null);
});
test('reads latest user edit instead of writing a stale snapshot',()=>{
  const x=host(),b=sample(x);x.ctx.SKU_CleanCurrentSheet();b.sh.Cells.Item(2,3).Value2='NEW ABC';x.flush();assert.equal(b.sh.Cells.Item(2,3).Value2,'NEWABC');
});
test('manual double-click creates one task; disable cancels outstanding changes',()=>{
  const x=host(),b=sample(x);x.ctx.SkuAddinLoad({});x.ctx.SKU_CleanCurrentSheet();x.ctx.SKU_CleanCurrentSheet();assert.equal(x.ctx.SKU_RUN.pending.length,1);
  x.ctx.SkuAddinDisableAuto();x.flush();assert.equal(b.sh.Cells.Item(2,3).Value2,"'2714230032211");
});
test('missing timer API fails closed without synchronous fallback',()=>{
  const x=host({noTimers:true}),b=sample(x);x.metrics.writes=0;x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();
  assert.equal(x.ctx.SKU_RUN.enabled,false);assert.equal(x.metrics.writes,0);assert.ok(x.alerts.some(v=>v.includes('分批调度')));
});
test('10 write failures stop processing and pause auto, flags restored',()=>{
  const x=host({noArrays:true}),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';
  for(let r=2;r<=30;r++){const c=s.Cells.Item(r,1);c.Value2='001';Object.defineProperty(c,'NumberFormat',{get:()=> '@',set:()=>{throw Error('host failure')}})}
  x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.ctx.SKU_STATE.last.failed,10);assert.equal(x.ctx.SKU_RUN.enabled,false);assert.match(x.ctx.SKU_RUN.error,/10/);
});
test('in-WPS self-test is isolated and restores mode/settings',()=>{
  const x=host();sample(x);x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();const bound=x.ctx.SKU_RUN.book;
  x.ctx.SkuAddinSelfTest();assert.ok(x.alerts.at(-1).startsWith('WPS批量接口自测通过'));assert.equal(x.app.EnableEvents,true);assert.equal(x.ctx.SKU_RUN.enabled,true);assert.equal(x.ctx.SKU_RUN.book,bound);
});
test('elapsed-time budget yields early on a slow host',()=>{
  const x=host(),b=sample(x);let clock=0;const original=b.sh.Cells.Item;
  b.sh.Cells.Item=(...args)=>{clock+=10;return original(...args)};x.ctx.Date={now:()=>clock};
  x.ctx.SKU_CleanCurrentSheet();x.flush();assert.ok(x.metrics.maxVisited<=2);assert.equal(b.sh.Cells.Item(2,3).Value2,2714230032211);
});
test('selection without a cell range does not accidentally scan a whole sheet',()=>{
  const x=host(),b=sample(x);x.app.Selection=null;x.ctx.SKU_CleanSelection();assert.equal(x.timers.size,0);assert.equal(b.sh.Cells.Item(2,3).Value2,"'2714230032211");
});
test('changing protection between slices stops writes to that sheet',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';for(let r=2;r<=2000;r++)s.Cells.Item(r,1).Value2='001';
  x.ctx.SKU_CleanCurrentSheet();x.tick();s.ProtectContents=true;x.metrics.writes=0;x.flush();assert.equal(x.metrics.writes,0);assert.equal(x.ctx.SKU_STATE.last.protectedSheets,1);
});
test('changing SKU header during an active job aborts before more data writes',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';for(let r=2;r<=2000;r++)s.Cells.Item(r,1).Value2='001';
  x.ctx.SKU_CleanCurrentSheet();x.tick();s.Cells.Item(1,1).Value2='Other';x.metrics.writes=0;x.flush();assert.equal(x.metrics.writes,0);assert.match(x.ctx.SKU_RUN.error,/表头发生变化/);
});
test('failed text write restores the original format and reports the failure',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';const c=s.Cells.Item(2,1);c.Value2='A B';c.NumberFormat='General';
  Object.defineProperty(c,'Value2',{get:()=> 'A B',set:()=>{throw Error('write rejected')}});
  x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(c.NumberFormat,'General');assert.equal(c.Value2,'A B');assert.equal(x.ctx.SKU_STATE.last.failed,1);
});
test('large mixed numeric/alphanumeric paste preserves formulas, errors and precision boundaries',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';
  for(let r=2;r<=10001;r++)s.Cells.Item(r,1).Value2=r%2 ? " ' AB０３ ":'000123';
  s.Cells.Item(10002,1).Value2='1000000000000001';s.Cells.Item(10003,1).Formula='=1+2';s.Cells.Item(10004,1).Value2=2042;s.Cells.Item(10004,1).isError=true;s.Cells.Item(10005,1).Value2='123';s.Cells.Item(10005,1).MergeCells=true;
  x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.events.SheetChange(s,x.target(2,1,10004));x.flush();
  assert.equal(x.ctx.SKU_STATE.last.converted,5000);assert.equal(x.ctx.SKU_STATE.last.textFormatted,5000);
  assert.equal(x.ctx.SKU_STATE.last.long,1);assert.equal(x.ctx.SKU_STATE.last.formula,1);assert.equal(x.ctx.SKU_STATE.last.invalid,1);assert.equal(x.ctx.SKU_STATE.last.merged,1);
  assert.equal(s.Cells.Item(2,1).Value2,123);assert.equal(s.Cells.Item(3,1).Value2,'AB03');assert.equal(s.Cells.Item(10002,1).Value2,'1000000000000001');
});

test('timer scheduling error clears queued work and pauses mode',()=>{
  const x=host();sample(x);x.ctx.setTimeout=()=>{throw Error('timer unavailable')};x.ctx.SKU_CleanCurrentSheet();
  assert.equal(x.ctx.SKU_RUN.pending.length,0);assert.equal(x.ctx.SKU_RUN.current,null);assert.equal(x.ctx.SKU_RUN.enabled,false);assert.match(x.ctx.SKU_RUN.error,/timer unavailable/);
});
