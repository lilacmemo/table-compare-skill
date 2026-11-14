// 线上 Pro 打印 · 配置生效验证（输出侧）
// 只判「同一份数据，改一个字段，打印出来的页数/画布尺寸/内容是否跟着变」
//   paperSize  A4 vs A3        → 画布尺寸
//   direction  Portrait/Landscape → 画布宽高互换
//   maxRowsEachPage 0 vs 10    → 页数
//   scale      Origin vs FitPage → 页数
//   freeze     [Row]           → 第 2 页是否有表头（已知 FAIL，复验）
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const OUT = path.join(__dirname, 'out', 'pro-print');
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e).slice(0, 160)));

  const go = async () => {
    await page.goto(URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 60; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  if (!await go()) { console.log('ENV NOT READY'); await browser.close(); process.exit(1); }
  await page.waitForTimeout(2500);

  // 建数据：60 行 × 4 列
  await page.evaluate(async () => {
    const u = window.univerAPI;
    const wb = u.getActiveWorkbook();
    let sh = wb.getActiveSheet();
    const rows = [['月份', '销售额', '成本', '利润']];
    for (let i = 1; i <= 30; i++) { const v = 1000 + i * 37; rows.push(['M' + i, v, Math.round(v * 0.58), Math.round(v * 0.42)]); }
    sh.getRange(`A1:D${rows.length}`).setValues(rows);
    window.__sh = sh;
  });
  await page.waitForTimeout(500);

  const CASES = [
    { id: 'base', label: 'A4 纵向（基线）', cfg: (E) => ({ area: E.PrintArea.CurrentSheet, paperSize: E.PrintPaperSize.A4, direction: E.PrintDirection.Portrait, scale: E.PrintScale.Origin, maxRowsEachPage: 0 }) },
    { id: 'a3', label: 'A3 纵向', cfg: (E) => ({ paperSize: E.PrintPaperSize.A3 }) },
    { id: 'land', label: 'A4 横向', cfg: (E) => ({ paperSize: E.PrintPaperSize.A4, direction: E.PrintDirection.Landscape }) },
    { id: 'rows10', label: 'A4 纵向 + 每页最多 10 行', cfg: (E) => ({ direction: E.PrintDirection.Portrait, maxRowsEachPage: 10 }) },
    { id: 'fit', label: 'A4 纵向 + 整页缩放', cfg: (E) => ({ maxRowsEachPage: 0, scale: E.PrintScale.FitPage }) },
    { id: 'freeze', label: 'A4 纵向 + 重复标题行 [Row]', cfg: (E) => ({ scale: E.PrintScale.Origin, freeze: [E.PrintFreeze.Row] }) },
  ];

  const results = [];
  for (const c of CASES) {
    const r = await page.evaluate(async (cfgStr) => {
      const u = window.univerAPI, E = u.Enum, wb = u.getActiveWorkbook();
      const wait = (ms) => new Promise((r) => setTimeout(r, ms));
      const cfg = (new Function('E', 'return ' + cfgStr))(E);
      // 清掉上一轮遗留的打印画布，避免页数累加
      document.querySelectorAll('.printing-canvas-container, .printing-canvas').forEach((n) => {
        const box = n.closest('.printing-canvas-container') || n;
        if (box && box.parentNode && box !== document.body) box.parentNode.removeChild(box);
      });
      await wait(300);
      wb.updatePrintConfig(cfg);
      await wait(300);
      wb.print();
      await wait(4500);
      const cs = Array.from(document.querySelectorAll('.printing-canvas canvas'));
      const seen = new Set();
      const pages = [];
      cs.forEach((x) => {
        const d = x.toDataURL('image/png');
        if (seen.has(d)) return;   // 同一页会被挂两次，按像素去重
        seen.add(d);
        pages.push({ w: x.width, h: x.height, d });
      });
      return {
        pages: cs.length,
        raw: cs.length,
        sizes: pages.map((x) => x.w + 'x' + x.h),
        firstPage: pages[0] ? pages[0].d : null,
      };
    }, c.cfg.toString().replace(/^\(E\)\s*=>\s*/, ''));

    if (r.firstPage) fs.writeFileSync(path.join(OUT, `eff-${c.id}-p1.png`), Buffer.from(r.firstPage.split(',')[1], 'base64'));
    results.push({ id: c.id, label: c.label, pages: r.pages, raw: r.raw, sizes: r.sizes });
    console.log(c.label.padEnd(30), '去重页数=' + r.pages + '(原始' + r.raw + ')', '画布=' + r.sizes.join(', '));
    // 关掉打印预览再继续
    await page.evaluate(() => { try { window.univerAPI.getActiveWorkbook().closePrintDialog(); } catch (e) {} });
    await page.waitForTimeout(600);
  }

  fs.writeFileSync(path.join(OUT, 'print-effect-pro.json'), JSON.stringify({ results, errs }, null, 2));
  console.log('pageerrors', errs.length, errs.slice(0, 3));
  await browser.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
