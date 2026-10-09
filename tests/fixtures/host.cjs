'use strict';
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm');
const root=path.join(__dirname,'../..'),pkg=require('../../package.json');
function host(options={}) {
  const books=[],events={},registrations={},alerts=[],timers=new Map();
  const metrics={reads:0,writes:0,finds:0,ticks:0,maxVisited:0,bulkReads:0,bulkWrites:0};let timerId=0, bulkMutation=false;
  function makeBook(){
    const book={Name:'test'+books.length+'.xlsx',FullName:'C:/test/'+books.length+'.xlsx'};
    const cells=new Map();let lastRow=1,lastCol=1;
    const sh={Name:'Data',Parent:book,ProtectContents:false,Cells:{Item:(r,c)=>{
      metrics.reads++;const key=r+','+c;if(!cells.has(key)){
        let value=null,formula='',format='@';const cell={Row:r,Column:c,HasFormula:false,MergeCells:false,ClearContents(){this.Value2=null}};
        Object.defineProperty(cell,'NumberFormat',{configurable:true,get:()=>format,set:v=>{if(!bulkMutation)metrics.writes++;format=v;}});
        Object.defineProperty(cell,'Value2',{configurable:true,get:()=>value,set:v=>{
          if(!bulkMutation)metrics.writes++;value=v;lastRow=Math.max(lastRow,r);lastCol=Math.max(lastCol,c);
          if(options.onWrite) options.onWrite(sh,r,c);
        }});
        Object.defineProperty(cell,'Formula',{get:()=>formula,set:v=>{formula=v;cell.HasFormula=true;lastRow=Math.max(lastRow,r);lastCol=Math.max(lastCol,c);}});
        cells.set(key,cell);
      }return cells.get(key);
    }}};
    Object.defineProperty(sh,'UsedRange',{get:()=>({Row:1,Column:1,Rows:{Count:lastRow},Columns:{Count:lastCol}})});
    sh.inflate=(r,c)=>{lastRow=r;lastCol=c};
    sh.Range=address=>{
      const m=address.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if(!m)throw Error('bad range');
      function column(name){let n=0;for(const ch of name)n=n*26+ch.charCodeAt(0)-64;return n}
      const left=column(m[1]),right=column(m[3]),top=+m[2],bottom=+m[4];
      function cell(r,c){return sh.Cells.Item(r,c)}
      function all(fn){const out=[];for(let r=top;r<=bottom;r++)for(let c=left;c<=right;c++)out.push(fn(cell(r,c)));return out}
      function uniform(fn){const values=all(fn);return values.every(v=>v===values[0])?values[0]:null}
      const range={Row:top,Column:left,Rows:{Count:bottom-top+1},Columns:{Count:right-left+1},Count:(bottom-top+1)*(right-left+1)};
      // Internal host work is not a JavaScript-to-host cell access.
      function internal(fn){const before=metrics.reads;try{return fn()}finally{metrics.reads=before}}
      Object.defineProperty(range,'MergeCells',{get:()=>internal(()=>uniform(c=>c.MergeCells))});
      Object.defineProperty(range,'HasFormula',{get:()=>internal(()=>uniform(c=>c.HasFormula))});
      Object.defineProperty(range,'NumberFormat',{get:()=>internal(()=>uniform(c=>c.NumberFormat)),set:v=>{
        metrics.writes++;bulkMutation=true;try{internal(()=>all(c=>{c.NumberFormat=v}))}finally{bulkMutation=false}
      }});
      Object.defineProperty(range,'Value2',{get:()=>{
        if(options.noArrays)throw Error('array API unavailable');metrics.bulkReads++;
        return internal(()=>{const rows=[];for(let r=top;r<=bottom;r++){const row=[];for(let c=left;c<=right;c++)row.push(cell(r,c).Value2);rows.push(row)}return rows});
      },set:values=>{
        if(options.failBulkWrite)throw Error('native array write failed');
        if(options.noArrays)throw Error('array API unavailable');
        metrics.writes++;metrics.bulkWrites++;bulkMutation=true;
        try{internal(()=>{for(let r=top;r<=bottom;r++)for(let c=left;c<=right;c++){
          const value=Array.isArray(values)?values[r-top][c-left]:values;
          if(typeof value==='string' && cell(r,c).NumberFormat!=='@')throw Error('unsafe text array write');
          cell(r,c).Value2=value;
        }})}finally{bulkMutation=false}
      }});
      range.SpecialCells=(type,flags)=>{
        if(options.noSpecialCells)throw Error('special API unavailable');
        return internal(()=>{
          const spans=[];let current=null;
          for(let r=top;r<=bottom;r++){
            const c=cell(r,left),v=c.Value2;
            const match=type===-4123 ? c.HasFormula : type===2 && !c.HasFormula && !c.isError &&
              ((typeof v==='number' && ((flags||23)&1)) || (typeof v==='string' && v!=='' && ((flags||23)&2)));
            if(match){if(!current){current={Row:r,Column:left,Rows:{Count:0},Columns:{Count:1}};spans.push(current)}current.Rows.Count++}else current=null;
          }
          if(!spans.length)throw Error('no cells found');
          return {Count:spans.reduce((n,a)=>n+a.Rows.Count,0),Areas:{Count:spans.length,Item:i=>spans[i-1]}};
        });
      };
      range.Find=(...args)=>{
        if(options.noFind)throw Error('unsupported');metrics.finds++;
        let best=null;for(const c of cells.values()){
          if(c.Column===left && c.Row>=top && c.Row<=bottom && (c.HasFormula || (c.Value2!==null && c.Value2!==''))){if(!best || c.Row>best.Row)best=c}
        }return best;
      };
      return range;
    };
    book.Worksheets={Count:1,Item:()=>sh};book.sh=sh;books.push(book);app.ActiveWorkbook=book;app.ActiveSheet=sh;return book;
  }
  const app={EnableEvents:true,ScreenUpdating:true,WorksheetFunction:{IsError:c=>!!c.isError},
    Workbooks:{get Count(){return books.length},Item:i=>books[i-1],Add:makeBook},
    PluginStorage:{getItem:()=>true,setItem:()=>{throw Error('must not persist auto mode')}},
    ApiEvent:{AddApiEventListener:(k,fn)=>{events[k]=fn;registrations[k]=(registrations[k]||0)+1}}};
  const sandbox={Application:app,alert:x=>alerts.push(x)};
  if(!options.noTimers){sandbox.setTimeout=fn=>{timers.set(++timerId,fn);return timerId};sandbox.clearTimeout=id=>timers.delete(id)}
  const ctx=vm.createContext(sandbox);vm.runInContext('window=globalThis',ctx);
  vm.runInContext(fs.readFileSync(path.join(root,'offline','WPS-SKU-Cleaner_'+pkg.version,'main.js'),'utf8'),ctx);
  function tick(){
    const next=timers.entries().next().value;if(!next)return false;
    const before=ctx.SKU_RUN.current ? ctx.SKU_RUN.current.stats.visited : 0;
    timers.delete(next[0]);next[1]();metrics.ticks++;
    const after=ctx.SKU_STATE.last ? ctx.SKU_STATE.last.visited : 0;
    metrics.maxVisited=Math.max(metrics.maxVisited,after-before);
    if(app.EnableEvents!==true || app.ScreenUpdating!==true || ctx.SKU_STATE.busy)throw Error('host settings not restored between slices');
    return true;
  }
  function flush(limit=50000){let count=0;while(timers.size){if(++count>limit)throw Error('timer feedback loop');tick()}return count}
  function target(r,c,rows=1,cols=1){return {Areas:{Count:1,Item:()=>({Row:r,Column:c,Rows:{Count:rows},Columns:{Count:cols}})}}}
  return {ctx,app,events,registrations,alerts,makeBook,metrics,timers,tick,flush,target,books};
}
module.exports={host,root,pkg};