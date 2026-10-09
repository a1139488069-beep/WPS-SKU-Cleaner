'use strict';
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const routes = {
  '/': ['install.html', 'text/html; charset=utf-8'],
  '/install.html': ['install.html', 'text/html; charset=utf-8'],
  '/ribbon.xml': ['addin/ribbon.xml', 'application/xml; charset=utf-8'],
  '/main.js': ['addin/main.js', 'text/javascript; charset=utf-8'],
  '/core.js': ['js/SkuCleaner.js', 'text/javascript; charset=utf-8'],
  '/adapter.js': ['addin/adapter.js', 'text/javascript; charset=utf-8'],
  '/vendor/wpsjsrpcsdk.js': ['vendor/wpsjsrpcsdk.js', 'text/javascript; charset=utf-8']
};
function createServer() {
  return http.createServer((req, res) => {
    let pathname;
    try { pathname = new URL(req.url, 'http://127.0.0.1').pathname; }
    catch { res.writeHead(400); return res.end(); }
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405); return res.end(); }
    if (pathname === '/health') {
      res.setHeader('Content-Type', 'application/json');
      return res.end(req.method === 'HEAD' ? '' : JSON.stringify({app:'WPS-SKU-Cleaner',version:require('./package.json').version}));
    }
    // Generated delivery page; no index.html file in the native add-in source.
    if (pathname === '/index.html') {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.end(req.method === 'HEAD' ? '' : '<!doctype html><html><head><meta charset="utf-8"></head><body><script src="/main.js"></script></body></html>');
    }
    const route = routes[pathname];
    if (!route) { res.writeHead(404); return res.end('Not found'); }
    fs.readFile(path.join(__dirname, route[0]), (err, data) => {
      if (err) { res.writeHead(500); return res.end('Required project file missing'); }
      res.setHeader('Content-Type', route[1]);
      res.end(req.method === 'HEAD' ? undefined : data);
    });
  });
}
if (require.main === module) {
  const port = Number(process.env.SKU_PORT || 39871);
  const server = createServer();
  server.on('error', err => { console.error(err.message); process.exitCode = 1; });
  server.listen(port, '127.0.0.1', () => console.log('WPS-SKU-Cleaner http://127.0.0.1:' + port + '/install.html'));
}
module.exports = {createServer};
