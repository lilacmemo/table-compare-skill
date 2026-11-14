// 对照台布局诊断：找"高度被谁吃掉了"
//   node tools/layout_diag.js                 默认宽度 1500×950，只打功能视图
//   node tools/layout_diag.js 3              三个视图都打（func / perf / api）+ 截图
// 用途：面板里出现莫名的横条、或画布明显比容器矮时，第一时间跑这个——
//   逐层打印 #uniHost 子树的实际高度，第一个"比父元素矮"的层就是元凶。
// 实战：元凶是 Univer 自己渲染的 <header>/<footer> 被宿主页裸元素选择器改了高度。
// 依赖：playwright（NODE_PATH 指向装着它的 node_modules）；PW_CHANNEL 默认 msedge。

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8473);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const ALL_VIEWS = process.argv.includes('3') || process.argv.includes('--all');

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end('not found'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

// 逐层高度（在页面内执行，不能用外部函数）
function dumpTree() {
  const out = [];
  const walk = (el, depth) => {
    const r = el.getBoundingClientRect();
    const st = getComputedStyle(el);
    const cls = (el.className && typeof el.className === 'string')
      ? el.className.replace(/dark:!?\S+/g, '').slice(0, 70) : el.tagName;
    out.push({ d: depth, h: Math.round(r.height), w: Math.round(r.width), cls, flex: st.flex, disp: st.display });
    if (depth < 12) [...el.children].forEach((c) => walk(c, depth + 1));
  };
  walk(document.getElementById('uniHost'), 0);
  return out;
}

function sizes() {
  const rect = (id) => { const r = document.getElementById(id).getBoundingClientRect(); return Math.round(r.width) + 'x' + Math.round(r.height); };
  const canvases = [...document.querySelectorAll('#uniHost canvas')]
    .map((c) => { const r = c.getBoundingClientRect(); return Math.round(r.width) + 'x' + Math.round(r.height); });
  return { uniHost: rect('uniHost'), sjsHost: rect('sjsHost'), footer: rect('footer'), canvases };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page.on('pageerror', (e) => console.log('PAGEERR', e.message));

  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__spread && window.__wb && window.univerAPI, null, { timeout: 60000 });
  await page.waitForTimeout(1500);

  const views = ALL_VIEWS ? ['func', 'perf', 'api'] : ['func'];
  for (const v of views) {
    if (v !== 'func') { await page.click(`.tab[data-view="${v}"]`); await page.waitForTimeout(900); }
    const s = await page.evaluate(sizes);
    console.log(`\n=== ${v} ===  uniHost=${s.uniHost}  sjsHost=${s.sjsHost}  footer=${s.footer}  canvas=[${s.canvases.join(', ')}]`);

    const rows = await page.evaluate(dumpTree);
    // 只打"承接高度"的层：canvas/HEADER/FOOTER 与高度发生变化的层
    const interesting = rows.filter((r, i) => {
      const prev = rows[i - 1];
      const isTag = /HEADER|FOOTER|CANVAS/.test(r.cls);
      const shrunk = prev && prev.d < r.d && r.h < prev.h - 1;
      return r.d === 0 || isTag || shrunk;
    });
    console.log(interesting.map((r) => '  ' + '  '.repeat(r.d) + `d${r.d} ${String(r.h).padStart(4)}h ${String(r.w).padStart(4)}w flex=${r.flex} | ${r.cls}`).join('\n'));

    const canvasH = parseInt((s.canvases.filter((c) => !c.startsWith('0x'))[0] || '0x0').split('x')[1], 10);
    const hostH = parseInt(s.uniHost.split('x')[1], 10);
    if (canvasH && hostH - canvasH > 2) {
      console.log(`  ⚠ 画布比容器矮 ${hostH - canvasH}px —— 看上面哪个 HEADER/FOOTER 在吃高度；` +
        '多半是宿主页裸元素选择器（header/footer）命中了 Univer 内部标签。');
    }
    await page.screenshot({ path: path.join(OUT, `layout-${v}.png`) });
  }

  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
