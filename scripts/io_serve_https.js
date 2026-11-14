// 对比工程本地服务（HTTPS 静态站 + exchange 反向代理）
//   node tools/serve.js            默认 8124
//   node tools/serve.js 9000       指定端口
//   node tools/serve.js 8124 --open  起好后自动打开浏览器
//
// 为什么必须 HTTPS：Univer 的 exchange 服务用 /universer-api/anonymous 种的会话
// cookie 带 Secure 标记，HTTP 下浏览器不会保存，后续 upload 直接 401。
// 自签证书在 tools/cert.pem|key.pem（首次自动生成，Playwright 用 ignoreHTTPSErrors 忽略告警）。
//
// 为什么必须代理：官方 exchange 端点不带 CORS 头，浏览器从本机直连会被拦；
// 这里把 /universer-api/* 同源反代到 office.univer.ai。
const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { exec } = require('child_process');

const ROOT = path.join(__dirname, '..', '..');
const PORT = Number(process.argv[2] || process.env.PORT || 8124);
const OPEN = process.argv.includes('--open');
const UPSTREAM = 'https://office.univer.ai';
const CERT = path.join(__dirname, 'cert.pem');
const KEY = path.join(__dirname, 'key.pem');

// 页面地址相对 ROOT：/<工程目录名>/index.html（工程目录名是中文，必须转义）
const PAGE = '/' + encodeURIComponent(path.basename(path.join(__dirname, '..'))) + '/index.html';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.pem': 'text/plain',
};

function proxy(req, res) {
  const target = UPSTREAM + req.url;
  const headers = Object.assign({}, req.headers);
  headers.host = 'office.univer.ai';
  headers.origin = UPSTREAM;
  if (headers.referer) headers.referer = UPSTREAM + '/';
  // 上游是 https，客户端到我们是 https；把 cookie 原样带上（含 showcase_visitor_session）
  const p = https.request(target, { method: req.method, headers }, (pr) => {
    const h = Object.assign({}, pr.headers);
    // 不能删 content-encoding：上游返回 gzip，浏览器要靠这个头解压；
    // 删了会把压缩字节当明文喂给 JSON.parse（表现为 import 静默返回 undefined）。
    // set-cookie 里的 Domain 若是 office.univer.ai 要摘掉，否则本机存不下
    if (h['set-cookie']) {
      h['set-cookie'] = h['set-cookie'].map((c) =>
        c.replace(/;\s*Domain=[^;]+/i, '').replace(/;\s*Secure/i, '')
      );
    }
    res.writeHead(pr.statusCode, h);
    pr.pipe(res);
  });
  p.on('error', (e) => {
    console.error('[proxy error]', e.message);
    if (!res.headersSent) res.writeHead(502, { 'content-type': 'application/json' });
    res.end(JSON.stringify({ error: 'proxy failed: ' + e.message }));
  });
  req.pipe(p);
}

function handle(req, res) {
  const u = url.parse(req.url);
  if (u.pathname.startsWith('/universer-api/')) return proxy(req, res);

  let p = decodeURIComponent(u.pathname);
  if (p === '/') p = '/index.html';
  const file = path.join(ROOT, p);
  if (!file.startsWith(ROOT)) {
    res.writeHead(403);
    return res.end('forbidden');
  }
  fs.stat(file, (err, st) => {
    if (err || !st.isFile()) {
      res.writeHead(404, { 'content-type': 'text/plain; charset=utf-8' });
      return res.end('404 ' + p);
    }
    res.writeHead(200, {
      'content-type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      'content-length': st.size,
      'cache-control': 'no-cache',
    });
    fs.createReadStream(file).pipe(res);
  });
}

const useHttps = fs.existsSync(CERT) && fs.existsSync(KEY);
const PAGE_URL = (useHttps ? 'https' : 'http') + '://127.0.0.1:' + PORT + PAGE;

const server = useHttps
  ? https.createServer({ cert: fs.readFileSync(CERT), key: fs.readFileSync(KEY) }, handle)
  : http.createServer(handle);

// 端口占用是最常见的启动失败，不要吐一整屏堆栈，给一句能照着做的提示
server.on('error', (e) => {
  console.error('');
  if (e.code === 'EADDRINUSE') {
    console.error('[启动失败] 端口 ' + PORT + ' 已被占用（可能上次的服务还没关）。');
    console.error('  换个端口启动：  node tools/serve.js ' + (PORT + 1) + ' --open');
    console.error('  或查谁占着：    netstat -ano | findstr :' + PORT);
    console.error('  或直接双击 start.cmd（会自动找空闲端口）');
  } else if (e.code === 'EACCES') {
    console.error('[启动失败] 端口 ' + PORT + ' 无权限（1024 以下需要管理员）。请换一个 1024 以上的端口。');
  } else {
    console.error('[启动失败] ' + e.message);
  }
  console.error('');
  process.exit(1);
});

server.listen(PORT, '127.0.0.1', () => {
  console.log('');
  if (!useHttps) {
    console.log('  [HTTP 模式] 缺少 tools/cert.pem，exchange 会话建不起来（导入导出会 401，只够看 UI）');
  }
  console.log('  对比工程: ' + PAGE_URL);
  if (useHttps) console.log('  自签证书：浏览器首次会提示"不安全"，点"继续访问"即可');
  console.log('  exchange 代理: /universer-api/* -> ' + UPSTREAM);
  console.log('  停止服务：Ctrl+C');
  console.log('');
  if (OPEN) exec('start "" "' + PAGE_URL + '"', { shell: 'cmd.exe' }, () => {});
});
