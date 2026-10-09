'use strict';
const test=require('node:test'),assert=require('node:assert/strict');
const {host}=require('./fixtures/host.cjs');
function sheet(x,values){const b=x.makeBook(),s=b.sh;s.Cells.Item(1,1).Value2='SKU';values.forEach((v,i)=>s.Cells.Item(i+2,1).Value2=v);return s}
test('20000 rows/two columns use at most 80 write calls instead of per-cell calls',()=>{
  const x=host(),b=x.makeBook(),s=b.sh;x.ctx.Date={now:()=>0};
  s.Cells.Item(1,1).Value2='公司sku';s.Cells.Item(1,2).Value2='平台sku';
  for(let r=2;r<=20001;r++){s.Cells.Item(r,1).Value2='000123';s.Cells.Item(r,2).Value2=" 'AB C-１２３ "}
  x.metrics.reads=0;x.metrics.writes=0;x.ctx.SKU_CleanCurrentSheet();x.flush();
  assert.equal(x.ctx.SKU_STATE.last.visited,40000);assert.equal(x.ctx.SKU_STATE.last.converted,20000);assert.equal(x.ctx.SKU_STATE.last.textFormatted,20000);
  assert.equal(x.ctx.SKU_STATE.last.batches,40);assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);
  assert.equal(x.metrics.bulkWrites,40);assert.equal(x.metrics.writes,80);assert.ok(x.metrics.reads<100);assert.ok(x.metrics.ticks<=45);
  assert.equal(s.Cells.Item(20001,1).Value2,123);assert.equal(s.Cells.Item(20001,2).Value2,'ABC-123');
});
test('bulk mask excludes formulas, constant errors, booleans, blank, long IDs and invalid numbers',()=>{
  const x=host(),s=sheet(x,['0001',' A B ',999,'1000000000000001',2042,true,null,-1,1.2,'\u200b ','0009']);
  s.Cells.Item(4,1).HasFormula=true;s.Cells.Item(4,1).Formula='=999';s.Cells.Item(6,1).isError=true;
  const protectedRows=[4,5,6,7,8,9,10];x.ctx.SKU_CleanCurrentSheet();x.flush();
  for(const r of protectedRows)assert.equal(s.Cells.Item(r,1).NumberFormat,'@');
  assert.equal(s.Cells.Item(2,1).Value2,1);assert.equal(s.Cells.Item(3,1).Value2,'AB');assert.equal(s.Cells.Item(11,1).Value2,null);assert.equal(s.Cells.Item(11,1).NumberFormat,'0');assert.equal(s.Cells.Item(12,1).Value2,9);
  assert.equal(x.ctx.SKU_STATE.last.formula,1);assert.equal(x.ctx.SKU_STATE.last.long,1);assert.equal(x.ctx.SKU_STATE.last.invalid,4);assert.ok(x.ctx.SKU_STATE.last.batches>0);assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);
});
test('mixed text and numeric arrays preserve formula/date/scientific looking strings',()=>{
  const x=host(),s=sheet(x,['000123',' =1+2 ',' 1E12 ',' 12.0 ',' 1/2 ',' 001AB ','000456']);
  x.ctx.SKU_CleanCurrentSheet();x.flush();
  [123,'=1+2','1E12','12.0','1/2','001AB',456].forEach((v,i)=>{const c=s.Cells.Item(i+2,1);assert.equal(c.Value2,v);assert.equal(c.HasFormula,false);assert.equal(c.NumberFormat,'0')});assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);
});
test('all-formula blocks skip without writing evaluated values',()=>{
  const x=host(),s=sheet(x,Array(1000).fill(123));for(let r=2;r<=1001;r++)s.Cells.Item(r,1).Formula='=123';
  x.metrics.writes=0;x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.ctx.SKU_STATE.last.formula,1000);assert.equal(x.metrics.writes,0);assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);
});
test('merged exceptions are isolated, then bulk processing resumes',()=>{
  const x=host(),s=sheet(x,Array(2500).fill('001'));x.ctx.Date={now:()=>0};s.Cells.Item(1020,1).MergeCells=true;
  x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(s.Cells.Item(1020,1).Value2,'001');assert.equal(s.Cells.Item(1020,1).NumberFormat,'@');assert.equal(x.ctx.SKU_STATE.last.merged,1);assert.equal(x.ctx.SKU_STATE.last.converted,2499);assert.ok(x.ctx.SKU_STATE.last.fallbackCells<=25);assert.ok(x.ctx.SKU_STATE.last.batches>=2);
});
test('mixed formats are normalized but skipped cells retain their own formats',()=>{
  const x=host(),s=sheet(x,['ABC','001AB','1000000000000001',' DEF ','12.0']);s.Cells.Item(2,1).NumberFormat='General';s.Cells.Item(3,1).NumberFormat='0';s.Cells.Item(5,1).NumberFormat='0.00';
  x.ctx.SKU_CleanCurrentSheet();x.flush();for(const r of [2,3,5,6])assert.equal(s.Cells.Item(r,1).NumberFormat,'0');assert.equal(s.Cells.Item(4,1).NumberFormat,'@');assert.equal(s.Cells.Item(5,1).Value2,'DEF');
});
test('missing bulk APIs report fallback counts instead of claiming bulk success',()=>{
  const x=host({noArrays:true}),s=sheet(x,Array(60).fill('001'));x.ctx.Date={now:()=>0};x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.ctx.SKU_STATE.last.fallbackCells,60);assert.equal(x.ctx.SKU_STATE.last.batches,0);assert.equal(s.Cells.Item(61,1).Value2,1);
});
test('unexpected out-of-scope SpecialCells results cause safe fallback only',()=>{
  const x=host(),s=sheet(x,['001','002','003']),original=s.Range;
  s.Range=address=>{const r=original(address);r.SpecialCells=()=>({Areas:{Count:1,Item:()=>({Row:1,Column:99,Rows:{Count:3},Columns:{Count:1}})}});return r};
  x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.ctx.SKU_STATE.last.fallbackCells,3);assert.equal(s.Cells.Item(1,99).Value2,null);assert.equal(s.Cells.Item(2,1).Value2,1);
});
test('bulk write error stops rather than retrying partial writes cell by cell',()=>{
  const x=host({failBulkWrite:true}),s=sheet(x,Array(1000).fill(' A B '));x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();x.ctx.SKU_CleanCurrentSheet();x.flush();assert.equal(x.ctx.SKU_STATE.last.failed,1);assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);assert.equal(x.ctx.SKU_RUN.enabled,false);assert.match(x.ctx.SKU_RUN.error,/批量写入失败/);assert.equal(s.Cells.Item(2,1).Value2,' A B ');
});
test('2万行测速 only creates its own book and verifies results on completion',()=>{
  const x=host(),old=sheet(x,['0009']);x.ctx.Date={now:()=>0};x.ctx.SkuAddinLoad({});x.ctx.SkuAddinEnableAuto();const bound=x.ctx.SKU_RUN.book;
  x.ctx.SkuAddinBenchmark();assert.equal(x.books.length,2);assert.equal(x.ctx.SKU_RUN.book,bound);x.flush();
  assert.equal(old.Cells.Item(2,1).Value2,'0009');assert.equal(x.ctx.SKU_STATE.last.visited,40000);assert.equal(x.ctx.SKU_STATE.last.fallbackCells,0);assert.ok(x.alerts.at(-1).includes('测速通过'));
});
test('batch self-test fails clearly when only per-cell compatibility mode is available',()=>{
  const x=host({noArrays:true});x.ctx.SkuAddinSelfTest();assert.ok(x.alerts.at(-1).startsWith('WPS批量接口自测未通过'));
});