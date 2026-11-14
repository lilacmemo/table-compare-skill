// 诊断 6：图例显隐读回 / 系列标签 family 字段名 / 饼图 setSeries|setExplosion 是否存在
//   node _diag_pro_chart6.js
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
      await wait(150);
      return c;
    };
    const cfg = (c) => JSON.parse(JSON.stringify(c.getInfo().config));

    // 1) 图例：true / false / 不设置
    r.legend = {};
    r.legend.baseline = cfg(await mk(T.Column)).legend || null;
    r.legend.trueOnly = cfg(await mk(T.Column, (b) => b.setLegend({ visible: true }))).legend || null;
    r.legend.falseOnly = cfg(await mk(T.Column, (b) => b.setLegend({ visible: false }))).legend || null;
    r.legend.falseWithPos = cfg(await mk(T.Column, (b) => b.setLegend({ visible: false, position: 'left' }))).legend || null;

    // 2) 系列标签 family vs fontFamily
    r.label = {};
    r.label.family = (cfg(await mk(T.Column, (b) => b.setSeries(1, { label: { visible: true, family: 'Arial', fontSize: 12, color: '#008000' } }))).series || []).find((x) => x.selector.index === 1);
    r.label.fontFamily = (cfg(await mk(T.Column, (b) => b.setSeries(1, { label: { visible: true, fontFamily: 'Arial' } }))).series || []).find((x) => x.selector.index === 1);
    // 轴标签 family 对照（C10 用的也是 family，已 PASS，说明轴的路径正常）
    r.label.axisFamily = (cfg(await mk(T.Column, (b) => b.setXAxis({ label: { family: 'Arial', fontSize: 13 } }))).axes || {}).x;

    // 3) 饼图 builder 到底有没有 setSeries / setExplosion
    const pb = sh.newChart(T.Pie);
    r.pie = { setSeries: typeof pb.setSeries, setExplosion: typeof pb.setExplosion, setPieLabel: typeof pb.setPieLabel };
    // 扇区分离候选：setPieLabel 无；试 setSeries→报错；试 slices 直接塞
    try {
      const c = await mk(T.Pie, (b) => { if (typeof b.setSeries === 'function') b.setSeries(1, { dataPoints: { 0: { explosion: 0.2 } } }); }, 'A1:B7');
      r.pie.afterSetSeries = cfg(c).pie || null;
    } catch (e) { r.pie.afterSetSeries = 'ERR ' + String(e).slice(0, 120); }
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 5000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
