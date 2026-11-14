// 诊断 8：exportImage 返回 undefined 的触发条件
//   node _diag_pro_chart8.js
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const go = async () => {
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  console.log('ready', await go());
  await page.waitForTimeout(2500);

  const out = await page.evaluate(async () => {
    const u = window.univerAPI, E = u.Enum, T = E.ChartTypeString;
    const wb = u.getActiveWorkbook();
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const r = {};
    const vals = [['Cat', 'S1', 'S2', 'S3']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5, 30 - i]);

    const trial = async (sheet, label) => {
      for (const c of sheet.getCharts()) { try { await c.remove(); } catch (e) {} }
      sheet.getRange('A1:D7').setValues(vals);
      await wait(200);
      const b = sheet.newChart(T.Column).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320);
      const c = await sheet.insertChart(b.build());
      const seq = [];
      for (let i = 0; i < 6; i++) {
        await wait(400);
        const res = await c.exportImage();
        seq.push(typeof res === 'string' ? 'str:' + res.length : String(res));
      }
      return { sheet: sheet.getSheetName(), isActive: wb.getActiveSheet().getSheetName() === sheet.getSheetName(), seq };
    };

    // A) 原始活动表
    r.orig = await trial(wb.getActiveSheet(), 'orig');

    // B) 新插入的表（probe 用法）
    let s = wb.getSheetByName('__chart8') || wb.insertSheet('__chart8');
    wb.setActiveSheet(s);
    r.inserted = await trial(s, 'inserted');

    // C) 新插入表 + 强制激活 + render
    try { wb.setActiveSheet(s); await wait(300); r.insertedAgain = await trial(s, 'inserted2'); } catch (e) {}
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 4000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
