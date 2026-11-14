// 诊断 4：剩余可疑项 —— exportImage 形态 / 3D view / 饼图 label / 趋势线字段 / chart 方法面
//   node _diag_pro_chart4.js
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
    const allMethods = (o) => {
      const s = new Set(); let p = Object.getPrototypeOf(o);
      while (p && p !== Object.prototype) { Object.getOwnPropertyNames(p).forEach((n) => s.add(n)); p = Object.getPrototypeOf(p); }
      return [...s].filter((n) => n[0] !== '_' && typeof o[n] === 'function').sort();
    };
    const r = {};
    const vals = [['Cat', 'S1', 'S2', 'S3']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5, 30 - i]);
    sh.getRange('A1:D7').setValues(vals);
    await wait(120);

    const mk = async (type, build) => {
      for (const c of sh.getCharts()) { try { await c.remove(); } catch (e) {} }
      const b = sh.newChart(type).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 5 }).setSize(480, 320);
      if (build) build(b);
      const c = await sh.insertChart(b.build());
      await wait(150);
      return c;
    };

    // 1) FChart 方法面
    const c0 = await mk(T.Column);
    r.chartMethods = allMethods(c0);
    r.exportImageSrc = typeof c0.exportImage === 'function' ? String(c0.exportImage).slice(0, 300) : typeof c0.exportImage;

    // 2) exportImage 各种调用形态
    r.export = {};
    const tries = [
      ['noArg', () => c0.exportImage()],
      ['emptyObj', () => c0.exportImage({})],
      ['type', () => c0.exportImage({ type: 'png' })],
      ['chartObj', () => c0.exportImage(c0.getInfo())],
    ];
    for (const [k, fn] of tries) {
      try {
        const res = await fn();
        r.export[k] = { t: typeof res, ctor: res && res.constructor && res.constructor.name, keys: res && typeof res === 'object' ? Object.keys(res).slice(0, 8) : null, len: typeof res === 'string' ? res.length : null };
      } catch (e) { r.export[k] = 'ERR ' + String(e).slice(0, 120); }
    }

    // 3) 3D：view 对象
    r.threeD = {};
    try {
      const c = await mk(T.Column3D, (b) => b.setThreeD({ view: { rotX: 30, rotY: 20, perspective: 30 }, grouping: 'clustered', gapDepth: 150 }));
      r.threeD.cfg = JSON.parse(JSON.stringify(c.getInfo().config)).threeD || null;
    } catch (e) { r.threeD.err = String(e).slice(0, 140); }

    // 4) 饼图 label（setPieLabel）
    r.pie = {};
    for (const [k, v] of [
      ['pct', { label: { visible: true, contentType: 16 } }],
      ['posIn', { label: { visible: true, position: 'inside' } }],
    ]) {
      try {
        const c = await mk(T.Pie, (b) => b.setPieLabel(v));
        r.pie[k] = JSON.parse(JSON.stringify(c.getInfo().config)).pie || null;
      } catch (e) { r.pie[k] = 'ERR ' + String(e).slice(0, 140); }
    }
    // 图例/标题等 Pie 是否支持
    r.pieLegendTitle = {};
    try {
      const c = await mk(T.Pie, (b) => b.setTitle({ text: 'P' }).setSource({ range: 'A1:B7' }));
      const cfg = JSON.parse(JSON.stringify(c.getInfo().config));
      r.pieLegendTitle = { title: cfg.title || null, legend: cfg.legend || null };
    } catch (e) { r.pieLegendTitle = 'ERR ' + String(e).slice(0, 140); }

    // 5) 趋势线全套字段
    r.trend = {};
    try {
      const c = await mk(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear', showFormula: true, showR2: true, forward: 2, backward: 1 }));
      r.trend.linear = JSON.parse(JSON.stringify(c.getInfo().config)).trendlines || null;
    } catch (e) { r.trend.err1 = String(e).slice(0, 160); }
    try {
      const c = await mk(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'exponential', order: 2, period: 3 }));
      r.trend.exp = JSON.parse(JSON.stringify(c.getInfo().config)).trendlines || null;
    } catch (e) { r.trend.err2 = String(e).slice(0, 160); }

    // 6) 误差线
    r.errBar = {};
    try {
      const b = sh.newChart(T.Line);
      r.errBar.hasSetErrorBars = typeof b.setErrorBars;
      const c = await mk(T.Line, (bb) => { if (typeof bb.setErrorBars === 'function') bb.setErrorBars({ seriesIndex: 1, direction: 'y', valueType: 'fixedVal', value: 30, type: 'both' }); });
      r.errBar.cfg = JSON.parse(JSON.stringify(c.getInfo().config)).errorBars || null;
    } catch (e) { r.errBar.err = String(e).slice(0, 160); }

    // 7) gapWidth / overlap 真字段
    r.bar = {};
    try {
      const c = await mk(T.Column, (b) => b.setBar({ gapWidth: 80, overlap: -20 }));
      const cfg = JSON.parse(JSON.stringify(c.getInfo().config));
      r.bar.cfg = cfg.bar || null;
    } catch (e) { r.bar.err = String(e).slice(0, 160); }

    // 8) doughnutHole / explosion
    r.pieExtra = {};
    try {
      const c = await mk(T.Donut, (b) => b.setDoughnutHole(0.4));
      r.pieExtra.donut = JSON.parse(JSON.stringify(c.getInfo().config)).pie || null;
    } catch (e) { r.pieExtra.donutErr = String(e).slice(0, 160); }
    try {
      const c = await mk(T.Pie, (b) => { if (typeof b.setRosePie === 'function') b.setRosePie(true); });
      r.pieExtra.rose = JSON.parse(JSON.stringify(c.getInfo().config)).pie || null;
    } catch (e) { r.pieExtra.roseErr = String(e).slice(0, 160); }

    // 9) 组合图：setSeriesType / setSeriesAxis
    r.combo = {};
    try {
      const c = await mk(T.Combination, (b) => b.setSeriesType(2, 'line').setSeriesAxis(2, 'right'));
      const cfg = JSON.parse(JSON.stringify(c.getInfo().config));
      r.combo.cfg = { series: cfg.series, rightY: cfg.axes && cfg.axes.rightY ? 'yes' : null };
    } catch (e) { r.combo.err = String(e).slice(0, 160); }

    // 10) 图表外观 IChartAppearanceSpec
    r.appearance = {};
    try {
      const b = sh.newChart(T.Column);
      r.appearance.hasSetAppearance = typeof b.setAppearance;
      const c = await mk(T.Column, (bb) => { if (typeof bb.setAppearance === 'function') bb.setAppearance({ background: { color: '#EEEEEE' }, border: { color: '#333333', width: 2 } }); });
      const cfg = JSON.parse(JSON.stringify(c.getInfo().config));
      r.appearance.cfg = cfg.appearance || cfg.background || null;
    } catch (e) { r.appearance.err = String(e).slice(0, 160); }
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 12000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
