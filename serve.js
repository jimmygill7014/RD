const http = require('http');
const fs = require('fs');
const path = require('path');
const dir = __dirname;
const mime = { '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript', '.json': 'application/json' };
http.createServer((req, res) => {
  let p = req.url === '/' ? '/index.html' : req.url.split('?')[0];
  const fp = path.join(dir, p);
  fs.readFile(fp, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, { 'Content-Type': mime[path.extname(fp)] || 'text/plain', 'Cache-Control': 'no-store' });
    res.end(data);
  });
}).listen(process.env.PORT || 8080, () => console.log('Serving on ' + (process.env.PORT || 8080)));
