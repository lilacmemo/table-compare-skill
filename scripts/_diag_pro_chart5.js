// 诊断 5：饼图标签落库路径（单系列）+ 扇区分离 + 命名/选中
//   node _diag_pro_chart5.js
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1600, height: 1000 } })).newPage();
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
    const sh = wb.getActiveSheet();
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const r = {};
    const vals = [['Cat', 'S1']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3]);
    sh.getRange('A1:B7').setValues(vals);
    await wait(120);

    const mk = async (type, build, range) => {
      for (const c of sh.getCharts()) { try { await c.remove(); } catch (e) {} }
      const b = sh.newChart(type).setSource({ range: range || 'A1:B7' }).setPosition({ row: 1, column: 5 }).setSize(480, 320);
      if (build) build(b);
      const c = await sh.insertChart(b.build());
      await wait(150);
      return c;
    };
    const pieCfg = (c) => { const k = JSON.parse(JSON.stringify(c.getInfo().config)); return { pie: k.pie || null, series: k.series || null }; };

    r.pie = {};
    r.pie.flat = pieCfg(await mk(T.Pie, (b) => b.setPieLabel({ visible: true, contentType: 16, position: 'outside' })));
    r.pie.onlyContentType = pieCfg(await mk(T.Pie, (b) => b.setPieLabel({ visible: true, contentType: 16 })));
    r.pie.onlyPosition = pieCfg(await mk(T.Pie, (b) => b.setPieLabel({ position: 'outside' })));
    r.pie.baseline = pieCfg(await mk(T.Pie));
    try {
      r.pie.fontAndColor = pieCfg(await mk(T.Pie, (b) => b.setPieLabel({ visible: true, fontSize: 13, color: '#008000', family: 'Arial' })));
    } catch (e) { r.pie.fontAndColor = 'ERR ' + String(e).slice(0, 120); }
    // 扇区分离
    try {
      r.pie.explode = pieCfg(await mk(T.Pie, (b) => b.setSeries(1, { dataPoints: { 0: { explosion: 0.2 } } })));
    } catch (e) { r.pie.explode = 'ERR ' + String(e).slice(0, 120); }
    // 枚举值
    r.enums = { PieLabelPosition: E.ChartPieLabelPosition || null, LabelContentType: E.LabelContentType || null };

    // 命名 / 选中 / 悬停 / 动画
    const c = await mk(T.Column, null, 'A1:B7');
    r.chart = {
      hasSetName: typeof c.setName,
      hasGetName: typeof c.getName,
      hasSelect: typeof c.select,
      hasIsSelected: typeof c.isSelected,
      hasHover: typeof c.setHoverStyle,
      hasAnimation: typeof c.setAnimation,
      nameFields: Object.keys(c).slice(0, 12),
    };
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 6000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
