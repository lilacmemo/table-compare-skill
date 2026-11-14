// 打印预览 UI 并排取证：SpreadJS Designer 文件→打印 vs Univer openPrintDialog
//   NODE_PATH=<poc-compare/node_modules> node tools/capture_print_preview.js
// 产物：out/shots/print-ui-sjs.png / print-ui-uni.png（单侧元素截图）
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(ROOT, 'out', 'shots');
const SAMPLE = 'samples/打印-PDF-多元素样本.xlsx';

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 2600, height: 1500 }, deviceScaleFactor: 1.5 });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => console.log('PAGEERROR', e.message.slice(0, 140)));

  // ?autoload=0 时页面不自动灌样本，但 boot 完才能调 APP.importOne
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });

  const t = await p.evaluate(async (f) => {
    const r = await APP.loadSample(f, '打印-PDF-多元素样本.xlsx');
    return r;
  }, SAMPLE);
  console.log('IMPORT', JSON.stringify(t));
  await p.waitForTimeout(1500);

  // 表名对齐：两边都停在同一张表（打印样本首表即正文，跳过 Evaluation Version）
  const align = await p.evaluate(() => {
    const isEval = (n) => /evaluation/i.test(String(n || ''));
    const sp = APP.sjs;
    let si = -1;
    for (let i = 0; i < sp.getSheetCount(); i++) if (!isEval(sp.getSheet(i).name())) { si = i; break; }
    if (si >= 0) { sp.setActiveSheetIndex(si); sp.getSheet(si).showCell(0, 0); }
    const wb = APP.uni.getActiveWorkbook();
    const ws = wb.getActiveSheet();
    try { ws.scrollToCell(0, 0); } catch (e) {}
    return {
      sjsSheets: Array.from({ length: sp.getSheetCount() }, (_, i) => sp.getSheet(i).name()),
      uniSheets: wb.getSheets().map((s) => s.getSheetName()),
      uniActive: ws.getSheetName(),
    };
  });
  console.log('ALIGN', JSON.stringify(align));

  // ---- SJS：文件 → 打印（backstage，右列是打印预览）----
  await p.click('.gc-designer-fileMenuButton');
  await p.waitForTimeout(900);
  const hitPrint = await p.evaluate(() => {
    const items = Array.from(document.querySelectorAll('.gc-list-control-item'));
    const it = items.find((x) => (x.textContent || '').trim() === '打印');
    if (!it) return { ok: false, all: items.map((x) => (x.textContent || '').trim()).slice(0, 20) };
    it.click();
    return { ok: true };
  });
  console.log('SJS_PRINT_ITEM', JSON.stringify(hitPrint));
  // 预览是异步分页渲染的，等预览区出现
  await p.waitForTimeout(3500);
  await p.screenshot({ path: path.join(SHOTS, 'print-ui-sjs-full.png') });
  const paneSjs = await p.$('#paneSjs');
  await paneSjs.screenshot({ path: path.join(SHOTS, 'print-ui-sjs.png') });

  // ---- Univer：openPrintDialog ----
  const uniOpen = await p.evaluate(async () => {
    const wb = APP.uni.getActiveWorkbook();
    const out = { has: typeof wb.openPrintDialog };
    try { wb.openPrintDialog(); out.called = true; } catch (e) { out.err = e.message.slice(0, 120); }
    return out;
  });
  console.log('UNI_OPEN', JSON.stringify(uniOpen));
  await p.waitForTimeout(3500);
  await p.screenshot({ path: path.join(SHOTS, 'print-ui-uni-full.png') });
  const paneUni = await p.$('#paneUni');
  await paneUni.screenshot({ path: path.join(SHOTS, 'print-ui-uni.png') });

  await b.close();
  console.log('OK', fs.existsSync(path.join(SHOTS, 'print-ui-sjs.png')), fs.existsSync(path.join(SHOTS, 'print-ui-uni.png')));
})();
