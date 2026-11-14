// Univer 侧打印链路探测：print() 到底做了什么？能不能拿到可落盘的 PDF/图片？
//   NODE_PATH=../poc-compare/node_modules node tools/probe_univer_print.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out', 'print');
const PORT = Number(process.env.PORT || 8487);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.ttf': 'font/ttf',
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

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 } });
  const page = await ctx.newPage();
  page.on('pageerror', (e) => console.log('PAGEERR', e.message));
  const logs = [];
  page.on('console', (m) => { const t = m.text(); if (/print|licen|watermark/i.test(t)) logs.push(`[${m.type()}] ${t.slice(0, 200)}`); });
  page.on('popup', (p) => console.log('POPUP', p.url()));

  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.univerAPI && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1200);

  const step1 = await page.evaluate(async () => {
    const u = window.univerAPI;
    const wb = window.__wb;
    const ws = wb.getActiveSheet();
    const R = {};

    // 造一页有内容 + 中文 + 合并 + 条件格式的表
    const vals = [
      ['区域', 'Q1', 'Q2', '合计'],
      ['华东 1', 451, 274, 725],
      ['华南 1', 169, 194, 363],
      ['华北 1', 216, 494, 710],
    ];
    for (let r = 0; r < vals.length; r++) {
      for (let c = 0; c < vals[r].length; c++) ws.getRange(r, c).setValue(vals[r][c]);
    }
    // 打印配置
    try {
      wb.updatePrintConfig({
        area: u.Enum.PrintArea.CurrentSheet,
        paperSize: u.Enum.PrintPaperSize.A4,
        direction: u.Enum.PrintDirection.Landscape,
        scale: u.Enum.PrintScale.FitWidth,
        margin: u.Enum.PrintPaperMargin.Normal,
      });
      R.setLayout = 'ok';
    } catch (e) { R.setLayout = 'ERR ' + e.message; }
    try {
      wb.updatePrintRenderConfig({
        gridlines: true,
        hAlign: u.Enum.PrintAlign.Middle,
        headerFooter: [u.Enum.PrintHeaderFooter.WorksheetTitle],
      });
      R.setRender = 'ok';
    } catch (e) { R.setRender = 'ERR ' + e.message; }

    // 区域截图（另一条"导出图像"的公开路径）
    try {
      const shot = await ws.getRange('A1:D4').getScreenshot({ includeHeaders: false });
      R.rangeShot = typeof shot === 'string' ? ('base64 ' + shot.length) : String(shot);
      if (typeof shot === 'string') R.rangeShotData = shot.slice(0, 80);
    } catch (e) { R.rangeShot = 'ERR ' + e.message; }

    // 打印前的 DOM 快照
    R.domBefore = document.body.children.length;
    R.printFn = typeof wb.print;
    return R;
  });

  console.log('=== 配置与截图 ===');
  console.log(JSON.stringify(step1, null, 1));

  // 调 print()，观察 DOM / 新窗口
  const step2 = await page.evaluate(async () => {
    const wb = window.__wb;
    const R = { errors: [] };
    // headless Chromium 里 window.print() 不阻塞，但 Univer 可能先渲染打印视图
    try {
      const ret = wb.print();
      R.ret = String(ret && ret.constructor && ret.constructor.name);
      if (ret && ret.then) await ret;
    } catch (e) { R.errors.push('print: ' + e.message); }
    await new Promise((r) => setTimeout(r, 1500));
    R.domAfter = document.body.children.length;
    R.printishNodes = Array.from(document.querySelectorAll('*'))
      .filter((el) => /print/i.test(el.className && el.className.toString ? el.className.toString() : ''))
      .slice(0, 8).map((el) => el.tagName + '.' + el.className);
    R.bodyTop = Array.from(document.body.children).map((el) => el.tagName + '.' + (el.className || '')).slice(0, 12);
    return R;
  });
  console.log('=== print() 行为 ===');
  console.log(JSON.stringify(step2, null, 1));
  console.log('=== console ===');
  console.log(logs.slice(0, 10).join('\n') || '(无)');

  // 用 Chromium 自己的 printToPDF 兜底：如果 Univer 用 @media print，这会带出来
  try {
    const pdf = await page.pdf({ format: 'A4', landscape: true, printBackground: true });
    fs.writeFileSync(path.join(OUT, 'univer-page.pdf'), pdf);
    console.log('page.pdf ->', pdf.length, 'bytes（仅当 Univer 走 @media print 时才可信）');
  } catch (e) { console.log('page.pdf ERR', e.message); }

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
