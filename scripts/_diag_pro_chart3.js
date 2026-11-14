// 诊断 3：验证 probe 的 __chart 新表面上数据是否真的写进去了
//   node _diag_pro_chart3.js
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
    const vals = [['Cat', 'S1', 'S2']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5]);

    // 方式 1：executeCommand（probe 现在用的）
    let s1 = wb.getSheetByName('__d3a') || wb.insertSheet('__d3a');
    wb.setActiveSheet(s1);
    await u.executeCommand('sheet.command.set-range-values', { value: vals, range: { startRow: 0, startColumn: 0, endRow: 6, endColumn: 2 } });
    await wait(100);
    r.cmdActiveSheet = { name: s1.getSheetName(), a1: s1.getRange('A1').getValue(), a2: s1.getRange('B2').getValue() };

    // 方式 2：FRange.setValues（带 unitId）
    let s2 = wb.getSheetByName('__d3b') || wb.insertSheet('__d3b');
    try {
      s2.getRange('A1:C7').setValues(vals);
      await wait(100);
      r.rangeSetValues = { a1: s2.getRange('A1').getValue(), b2: s2.getRange('B2').getValue() };
    } catch (e) { r.rangeSetValues = 'ERR ' + String(e).slice(0, 140); }

    // 两种方式各建一张图，看 config.series 是否有 3 个
    const mk = async (s, label) => {
      for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
      const b = s.newChart(T.Column).setSource({ range: 'A1:C7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320)
        .setSeries(1, { color: '#FFFF00' });
      const c = await s.insertChart(b.build());
      await wait(120);
      const cfg = c.getInfo().config;
      r[label] = { series: cfg.series || null, mapping: cfg.mapping, a1: s.getRange('A1').getValue() };
    };
    await mk(s1, 'chartOnCmdSheet');
    await mk(s2, 'chartOnRangeSheet');

    // 顺带：饼图 setPieLabel 是否可用、3D setThreeD 是否可用
    r.pieBuilderHasSetSeries = typeof s2.newChart(T.Pie).setSeries;
    r.threeDBuilderHasSetThreeD = typeof s2.newChart(T.Column3D).setThreeD;
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 4000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
