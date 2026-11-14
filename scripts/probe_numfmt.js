// 数字格式「引擎能力」单侧探针：不经过导入，直接用 API 写值 + 写格式，看显示结果。
//   目的：把「导入丢格式」与「Univer numfmt 引擎本身不支持」分开。
// 用法：NODE_PATH=... node tools/probe_numfmt.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = 8477;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
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

async function pageFn() {
  const api = window.univerAPI;
  const wb = window.__wb;
  const sh = wb.getActiveSheet();

  const cases = [
    { id: 'percent-2', value: 0.256, fmt: '0.00%' },
    { id: 'percent-0', value: 0.256, fmt: '0%' },
    { id: 'date-cn-weekday', value: '2025-01-21', fmt: 'yyyy"年"m"月"d"日" aaaa' },
    { id: 'date-cn-weekday-numeric', value: '2025-01-21', fmt: 'yyyy"年"m"月"d"日" aaaa', alt: 'aaa' },
    { id: 'date-iso', value: '2025-01-21', fmt: 'yyyy-mm-dd' },
    { id: 'dbnum1', value: 12345, fmt: '[DBNum1][$-804]General' },
    { id: 'dbnum2', value: 12345, fmt: '[DBNum2][$-804]General' },
    { id: 'pad6', value: 123, fmt: '000000' },
    { id: 'currency-cn', value: 1234.5, fmt: '¥#,##0.00' },
    { id: 'text-format', value: '00123', fmt: '@' },
    { id: 'scientific', value: 123456789, fmt: '0.00E+00' },
    { id: 'thousand', value: 1234567.891, fmt: '#,##0.00' },
  ];

  const results = [];
  for (let i = 0; i < cases.length; i++) {
    const c = cases[i];
    const row = i + 1;
    const rec = { id: c.id, value: c.value, fmt: c.fmt };
    try {
      sh.getRange(row, 0).setValue(c.value);
      await sh.getRange(row, 0).setNumberFormat(c.fmt);
      await new Promise((r) => setTimeout(r, 30));
      const rg = sh.getRange(row, 0);
      rec.rawValue = rg.getValue();
      rec.display = typeof rg.getDisplayValue === 'function' ? rg.getDisplayValue() : '(无 getDisplayValue)';
      rec.readFormat = typeof rg.getNumberFormat === 'function' ? rg.getNumberFormat() : '(无 getNumberFormat)';
    } catch (e) {
      rec.error = String(e && e.message || e);
    }
    results.push(rec);
  }

  // 百分比单元格的编辑态：Univer 进入编辑时取的是 value 还是 display
  let editState = null;
  try {
    const rg = sh.getRange(1, 0);
    editState = { value: rg.getValue(), display: rg.getDisplayValue() };
  } catch (e) { editState = { error: String(e && e.message || e) }; }

  // 关键追问：套了百分比格式后，这一格还能不能参与计算？
  let calcAfterFormat = null;
  try {
    sh.getRange(20, 3).setFormula('=A2*2');
    await new Promise((r) => setTimeout(r, 60));
    calcAfterFormat = {
      a2Value: sh.getRange(1, 0).getValue(),
      a2Type: typeof sh.getRange(1, 0).getValue(),
      b2Formula: sh.getRange(20, 3).getFormula && sh.getRange(20, 3).getFormula(),
      b2Value: sh.getRange(20, 3).getValue(),
      b2Display: sh.getRange(20, 3).getDisplayValue(),
    };
  } catch (e) { calcAfterFormat = { error: String(e && e.message || e) }; }

  return { at: new Date().toISOString(), results, editState, calcAfterFormat };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1200, height: 800 } });
  page.on('pageerror', (e) => console.log('PAGEERR', String(e.message).slice(0, 160)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__wb && window.univerAPI, null, { timeout: 60000 });
  await page.waitForTimeout(1500);
  const res = await page.evaluate(pageFn);
  fs.writeFileSync(path.join(OUT, 'numfmt-engine.json'), JSON.stringify(res, null, 2), 'utf8');
  for (const r of res.results) {
    console.log([r.id, '值=' + JSON.stringify(r.rawValue), '显示=' + JSON.stringify(r.display), '格式=' + JSON.stringify(r.readFormat), r.error ? 'ERR:' + r.error : ''].join(' | '));
  }
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
