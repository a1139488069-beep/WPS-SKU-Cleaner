'use strict';
const test=require('node:test');
const assert=require('node:assert/strict');
const http=require('node:http');
const {createServer}=require('../server.cjs');
function request(url,method='GET'){
  return new Promise((resolve,reject)=>{
    const req=http.request(url,{method,agent:false},res=>{
      let text='';res.setEncoding('utf8');res.on('data',chunk=>text+=chunk);
      res.on('end',()=>resolve({status:res.statusCode,text}));
    });req.on('error',reject);req.end();
  });
}
test('local delivery serves complete installation/add-in files and blocks unrelated paths',async()=>{
  const server=createServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));
  const root='http://127.0.0.1:'+server.address().port;
  try{
    for(const route of ['/install.html','/ribbon.xml','/index.html','/main.js','/core.js','/adapter.js','/vendor/wpsjsrpcsdk.js']){
      const res=await request(root+route);assert.equal(res.status,200,route);assert.ok(res.text.length>30,route);
      if(route==='/ribbon.xml') assert.ok(res.text.startsWith('<customUI'),'WPS SDK requires customUI as the first characters');
    }
    assert.equal(JSON.parse((await request(root+'/health')).text).app,'WPS-SKU-Cleaner');
    for(const route of ['/package.json','/server.cjs','/README.md','/%2e%2e/package.json']) assert.equal((await request(root+route)).status,404);
    assert.equal((await request(root+'/core.js','POST')).status,405);
  }finally{await new Promise(r=>server.close(r));}
});
