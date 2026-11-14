// 诊断 7：图例读回路径 / 饼图 setExplosion 读回
//   node _diag_pro_chart7.js
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
    const vals = [['Cat', 'S1', 'S2', 'S3']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5, 30 - i]);
    sh.getRange('A1:D7').setValues(vals);
    await wait(150);

    const mk = async (type, build, range) => {
      for (const c of sh.getCharts()) { try { await c.remove(); } catch (e) {} }
      const b = sh.newChart(type).setSource({ range: range || 'A1:D7' }).setPosition({ row: 1, column: 5 }).setSize(480, 320);
      if (build) build(b);
      const c = await sh.insertChart(b.build());
      await wait(200);
      return c;
    };
    const cfg = (c) => JSON.parse(JSON.stringify(c.getInfo().config));

    r.legend = {};
    const cases = {
      none: null,
      visTrue: { visible: true },
      visFalse: { visible: false },
      posLeft: { position: 'left' },
      visFalsePosLeft: { visible: false, position: 'left' },
      visTruePosLeft: { visible: true, position: 'left' },
      visTrueFont: { visible: true, family: 'Arial', fontSize: 15 },
    };
    for (const [k, v] of Object.entries(cases)) {
      const c = await mk(T.Column, v ? (b) => b.setLegend(v) : null);
      r.legend[k] = cfg(c).legend === undefined ? 'undefined' : cfg(c).legend;
    }
    // 二次修改：visible:false 通过 toBuilder().clearLegend()
    try {
      const c = await mk(T.Column, (b) => b.setLegend({ visible: true, position: 'left' }));
      const b2 = c.toBuilder();
      if (typeof b2.clearLegend === 'function') b2.clearLegend();
      await c.update(b2.build());
      await wait(200);
      r.legend.clearLegend = cfg(c).legend === undefined ? 'undefined' : cfg(c).legend;
    } catch (e) { r.legend.clearLegend = 'ERR ' + String(e).slice(0, 140); }

    // 饼图爆炸
    r.explode = {};
    try {
      const c = await mk(T.Pie, (b) => b.setExplosion(0.2), 'A1:B7');
      const k = cfg(c);
      r.explode.cfg = { pie: k.pie || null, series: k.series || null };
    } catch (e) { r.explode.err = String(e).slice(0, 160); }
    try {
      const pb = sh.newChart(T.Pie);
      r.explode.sig = String(pb.setExplosion).slice(0, 120);
    } catch (e) {}
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 5000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
