// 打印 / PDF 能力面探针
//   目的：在造用例之前先确认运行时到底有没有入口、能不能调、无 license 是什么表现。
//   符号进了 bundle ≠ facade 挂上了 ≠ 调用不报错，三者分开验。
//   NODE_PATH=../poc-compare/node_modules node tools/probe_print.js
// 输出：tools/out/print-surface.json + 控制台摘要

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8483);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';

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

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page.on('pageerror', (e) => console.log('PAGEERR', e.message));
  const logs = [];
  page.on('console', (m) => {
    const t = `[${m.type()}] ${m.text()}`;
    if (/licen|水印|watermark|print|pdf/i.test(t)) logs.push(t.slice(0, 300));
  });

  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__spread && window.univerAPI && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1500);

  const report = await page.evaluate(async () => {
    const R = { univer: {}, sjs: {}, enums: {} };
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const u = window.univerAPI;

    // ---------- Univer ----------
    const wbf = u.getActiveWorkbook();
    const want = ['updatePrintConfig', 'updatePrintRenderConfig', 'getPrintConfig',
      'getPrintRenderConfig', 'print', 'openPrintDialog', 'closePrintDialog',
      'saveScreenshotToClipboard'];
    R.univer.workbook = {};
    want.forEach((k) => { R.univer.workbook[k] = typeof (wbf && wbf[k]); });
    R.univer.printishKeys = wbf ? Object.keys(wbf).filter((k) => /print/i.test(k)) : [];

    const ws = wbf && wbf.getActiveSheet();
    R.univer.sheet = ws ? Object.keys(ws).filter((k) => /print|screenshot/i.test(k)) : [];
    try {
      const rg = ws.getRange('A1:D6');
      R.univer.rangeGetScreenshot = typeof rg.getScreenshot;
    } catch (e) { R.univer.rangeGetScreenshot = 'ERR ' + e.message; }

    ['PrintArea', 'PrintDirection', 'PrintScale', 'PrintPaperSize', 'PrintPaperMargin',
      'PrintFreeze', 'PrintHeaderFooter', 'PrintAlign'].forEach((k) => {
      R.enums[k] = u.Enum && u.Enum[k] ? Object.keys(u.Enum[k]) : null;
    });

    // 真设一次，看是否抛错
    try {
      wbf.updatePrintConfig({
        area: u.Enum.PrintArea.CurrentSheet,
        paperSize: u.Enum.PrintPaperSize.A4,
        direction: u.Enum.PrintDirection.Landscape,
        scale: u.Enum.PrintScale.FitWidth,
      });
      R.univer.setConfig = 'ok';
    } catch (e) { R.univer.setConfig = 'ERR ' + e.message; }

    try {
      wbf.updatePrintRenderConfig({ gridlines: false, hAlign: u.Enum.PrintAlign.Middle });
      R.univer.setRenderConfig = 'ok';
    } catch (e) { R.univer.setRenderConfig = 'ERR ' + e.message; }

    // 有没有 getter 能读回（判定 PASS 的关键）
    try { R.univer.readBack = wbf.getPrintConfig ? JSON.stringify(wbf.getPrintConfig()).slice(0, 300) : 'no getter'; }
    catch (e) { R.univer.readBack = 'ERR ' + e.message; }

    try { R.univer.readBackRender = wbf.getPrintRenderConfig ? JSON.stringify(wbf.getPrintRenderConfig()).slice(0, 300) : 'no getter'; }
    catch (e) { R.univer.readBackRender = 'ERR ' + e.message; }

    // ---------- SpreadJS ----------
    const sp = window.__spread;
    const sheet = sp.getActiveSheet();
    R.sjs.savePDF = typeof sp.savePDF;
    R.sjs.print = typeof sp.print;
    const pi = sheet.printInfo();
    const methods = [];
    for (const k in pi) if (typeof pi[k] === 'function') methods.push(k);
    R.sjs.printInfoMethods = methods.sort();
    R.sjs.printNamespace = window.GC && GC.Spread && GC.Spread.Sheets.Print
      ? Object.keys(GC.Spread.Sheets.Print) : null;

    // 真设一次 + 读回
    try {
      pi.orientation(GC.Spread.Sheets.Print.PrintPageOrientation.landscape);
      pi.margin({ top: 60, bottom: 60, left: 30, right: 30, header: 10, footer: 15 });
      R.sjs.setConfig = 'ok';
      R.sjs.readBack = JSON.stringify({ orientation: pi.orientation(), margin: pi.margin() });
    } catch (e) { R.sjs.setConfig = 'ERR ' + e.message; }

    await wait(300);
    return R;
  });

  report.console = logs.slice(0, 20);
  fs.writeFileSync(path.join(OUT, 'print-surface.json'), JSON.stringify(report, null, 2));

  console.log('=== Univer workbook 方法 ===');
  console.log(report.univer.workbook);
  console.log('print 相关键:', report.univer.printishKeys);
  console.log('sheet print 键:', report.univer.sheet);
  console.log('range.getScreenshot:', report.univer.rangeGetScreenshot);
  console.log('设置结果:', report.univer.setConfig, '| render:', report.univer.setRenderConfig);
  console.log('读回:', report.univer.readBack);
  console.log('读回 render:', report.univer.readBackRender);
  console.log('=== 枚举 ===');
  console.log(report.enums);
  console.log('=== SpreadJS ===');
  console.log('savePDF:', report.sjs.savePDF, '| print:', report.sjs.print);
  console.log('printInfo 方法数:', report.sjs.printInfoMethods.length);
  console.log(report.sjs.printInfoMethods.join(', '));
  console.log('Print 命名空间:', report.sjs.printNamespace);
  console.log('设置/读回:', report.sjs.setConfig, report.sjs.readBack);
  console.log('=== 控制台（license/print 相关） ===');
  console.log(report.console.join('\n') || '(无)');

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
