// 用 SpreadJS 渲染同一份样本（作为 Excel 基线参照），逐表截图。
// 用法：NODE_PATH=... node tools/render_samples.js [sheetName...]
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = 8478;
const FILE = path.join(ROOT, 'samples', '格式与条件格式-全量用例-A至K.xlsx');
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end('nf'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

(async () => {
  const targets = process.argv.slice(2);
  const sheets = targets.length ? targets : ['G条件格式规则', 'H条件格式参数', 'I条件格式管理'];
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => console.log('PAGEERR', String(e.message).slice(0, 160)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__spread && window.GC, null, { timeout: 60000 });
  await page.waitForTimeout(1200);

  const buf = fs.readFileSync(FILE);
  // v19 的 spread.import 返回 Promise（没有回调），传回调会报 "e is not a function"
  await page.evaluate(async (arr) => {
    const blob = new Blob([new Uint8Array(arr)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const r = window.__spread.import(blob, { fileType: window.GC.Spread.Sheets.FileType.excel });
    if (r && typeof r.then === 'function') await r;
  }, Array.from(buf));

  await page.waitForTimeout(2500);
  const names = await page.evaluate(() => {
    const b = window.__spread; const out = [];
    for (let i = 0; i < b.getSheetCount(); i++) out.push(b.getSheet(i).name());
    return out;
  });
  console.log('导入后工作表:', names.join(' / '));

  for (const nm of sheets) {
    const idx = names.indexOf(nm);
    if (idx < 0) { console.log('未找到表', nm); continue; }
    const info = await page.evaluate((i) => {
      const b = window.__spread;
      if (typeof b.setActiveSheetIndex === 'function') b.setActiveSheetIndex(i);
      if (typeof b.setActiveSheet === 'function') b.setActiveSheet(i);
      const sh = b.getActiveSheet();
      // 逐行统计条件格式规则数（C 列代表格）
      const rules = [];
      for (let r = 2; r < 18; r++) {
        const rs = sh.conditionalFormats.getRules(r, 2);
        rules.push({ row: r + 1, n: rs ? rs.length : 0 });
      }
      return { name: sh.name(), rowCount: sh.getRowCount(), rules };
    }, idx);
    console.log('==', info.name, '规则数:', JSON.stringify(info.rules.filter((x) => x.n)));
    await page.waitForTimeout(900);
    const safe = nm.replace(/[^\u4e00-\u9fa5A-Za-z0-9]/g, '');
    await page.screenshot({ path: path.join(OUT, `sjs-${safe}.png`) });
  }
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
