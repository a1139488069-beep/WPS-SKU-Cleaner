'use strict';
const test=require('node:test'), assert=require('node:assert/strict');
const fs=require('node:fs'), path=require('node:path'), vm=require('node:vm');
const read=p=>fs.readFileSync(path.join(__dirname,'..',p),'utf8');
test('ribbon callbacks and entry scripts resolve to supplied globals',()=>{
  const written=[];
  const app={};const window={Application:app};
  const ctx=vm.createContext({window,Application:app,document:{write:s=>written.push(s)},alert:()=>{}});
  vm.runInContext(read('addin/main.js'),ctx);
  assert.equal(written.length,0,'single bundle must not depend on document.write imports');
  assert.ok(read('addin/main.js').includes('function skuClassify'));
  vm.runInContext(read('js/SkuCleaner.js'),ctx);vm.runInContext(read('addin/adapter.js'),ctx);
  const xml=read('addin/ribbon.xml');
  for(const m of xml.matchAll(/(?:onLoad|onAction)="([^"]+)"/g)) assert.equal(typeof ctx[m[1]],'function',m[1]);
  assert.equal(ctx.SkuAddinLoad({}),true);
});
test('adapter supports legacy host exposing wps.EtApplication',()=>{
  const app={};const ctx=vm.createContext({window:{},wps:{EtApplication:()=>app},alert:()=>{}});
  vm.runInContext(read('addin/adapter.js'),ctx);assert.equal(ctx.SkuAddinLoad({}),true);
  assert.equal(ctx.window.Application,app);
});
test('installer displays stable SDK response fields, even with cyclic browser event',()=>{
  const html=read('install.html'),inline=html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  const status={textContent:''},target={response:'{"status":0}'};target.self=target;
  let received;
  const ctx=vm.createContext({window:{},location:{origin:'http://127.0.0.1:39871',protocol:'http:'},document:{getElementById:()=>status},
    WpsAddonMgr:{enable:(d,cb)=>{received=d;cb({status:0,res:{target}})}}});
  vm.runInContext(inline,ctx);ctx.manage('enable');assert.equal(received.addonType,'et');
  assert.equal(received.url,'http://127.0.0.1:39871/');assert.ok(status.textContent.includes('status'));
});
test('opening install.html directly gives launch instructions instead of a false SDK/WPS error',()=>{
  const html=read('install.html'),inline=html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  const status={textContent:''};let calls=0;
  const ctx=vm.createContext({window:{},location:{origin:'null',protocol:'file:'},document:{getElementById:()=>status},
    WpsAddonMgr:{enable:()=>calls++}});
  vm.runInContext(inline,ctx);assert.ok(status.textContent.includes('file://'));
  assert.ok(status.textContent.includes('Start-Local.cmd'));ctx.manage('enable');assert.equal(calls,0);
  assert.ok(html.includes('src="./vendor/wpsjsrpcsdk.js"'));
});
test('missing SDK on HTTP explains project loading failure and does not blame WPS startup',()=>{
  const inline=read('install.html').match(/<script>\s*([\s\S]*?)<\/script>/)[1];
  const status={textContent:''};const ctx=vm.createContext({window:{},location:{origin:'http://127.0.0.1:39871',protocol:'http:'},document:{getElementById:()=>status}});
  vm.runInContext(inline,ctx);ctx.manage('enable');assert.ok(status.textContent.includes('初始化失败'));
  assert.ok(status.textContent.includes('与WPS是否启动无关'));
});
test('actual vendored SDK initializes in a browser context and exposes the installation API',()=>{
  const ctx=vm.createContext({location:{protocol:'http:'},sessionStorage:{getItem:()=>null},
    XMLHttpRequest:function(){this.open=function(){};this.send=function(){};},
    navigator:{userAgent:'Mozilla/5.0 Chrome/130.0.0.0'},setInterval:()=>0,clearInterval:()=>{}});
  vm.runInContext('window=globalThis',ctx);
  vm.runInContext(read('vendor/wpsjsrpcsdk.js'),ctx);
  for(const key of ['enable','disable','verifyStatus']) assert.equal(typeof ctx.WpsAddonMgr[key],'function',key);
});
