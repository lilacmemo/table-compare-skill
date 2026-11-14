// 诊断 2：核查 Univer Pro 图表探针 47 条 FAIL 里哪些是「我写错」而非「产品没有」
//   node _diag_pro_chart2.js
// 覆盖 5 类可疑项：系列 patch / 位置更新 / exportImage / 饼图 label / 3D builder
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e).slice(0, 160)));

  const go = async () => {
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 60; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
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
      if (!o) return [];
      const s = new Set(); let p = Object.getPrototypeOf(o);
      while (p && p !== Object.prototype) { Object.getOwnPropertyNames(p).forEach((n) => s.add(n)); p = Object.getPrototypeOf(p); }
      return [...s].filter((n) => n[0] !== '_' && typeof o[n] === 'function').sort();
    };
    const r = {};

    // 造数据
    const vals = [['类别', 'S1', 'S2', 'S3']];
    for (let i = 0; i < 6; i++) vals.push(['C' + i, 10 + i, 20 + i * 2, 30 - i]);
    await u.executeCommand('sheet.command.set-range-values', { value: vals, range: { startRow: 0, startColumn: 0, endRow: 6, endColumn: 3 } });

    const clear = async () => { for (const c of sh.getCharts()) { try { await c.remove(); } catch (e) {} } };
    const mk = async (type, build) => {
      await clear();
      const b = sh.newChart(type).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320);
      if (build) build(b);
      const c = await sh.insertChart(b.build());
      await wait(80);
      return c;
    };
    const pick = (c, i) => ((c.getInfo().config.series || []).find((x) => x.selector && x.selector.index === i) || {});

    // ---------- 0. builder 方法面（按图型） ----------
    const seriesOf = async (type) => allMethods(sh.newChart(type)).filter((n) => /series|label|bar|pie|doughnut|three|axis|name|title|trend|error|dataTable|markLine/i.test(n));
    r.methods = {
      Column: await seriesOf(T.Column),
      Pie: await seriesOf(T.Pie),
      Column3D: await seriesOf(T.Column3D),
      Combination: await seriesOf(T.Combination),
      Line: await seriesOf(T.Line),
    };

    // ---------- 1. 系列 patch：三种 selector 写法 ----------
    r.series = {};
    try {
      let c = await mk(T.Column, (b) => b.setSeries(1, { color: '#FFFF00' }));
      r.series.numberSelector = pick(c, 1);
    } catch (e) { r.series.numberSelector = 'ERR ' + String(e).slice(0, 100); }
    try {
      let c = await mk(T.Column, (b) => b.setSeries({ index: 1 }, { color: '#FFFF00' }));
      r.series.objSelector = pick(c, 1);
    } catch (e) { r.series.objSelector = 'ERR ' + String(e).slice(0, 100); }
    try {
      let c = await mk(T.Column, (b) => b.setAllSeriesStyle({ color: '#FFFF00' }));
      const i = c.getInfo();
      r.series.allSeriesStyle = i.config.allSeriesStyle || i.config.series || null;
    } catch (e) { r.series.allSeriesStyle = 'ERR ' + String(e).slice(0, 100); }
    // 二次修改：update 后再读
    try {
      let c = await mk(T.Column);
      const b = c.toBuilder(); b.setSeries(1, { color: '#00FF00' });
      await c.update(b.build()); await wait(120);
      r.series.afterUpdate = pick(c, 1);
    } catch (e) { r.series.afterUpdate = 'ERR ' + String(e).slice(0, 140); }
    // 完整 config 快照（用带 patch 的图）
    try {
      let c = await mk(T.Column, (b) => b.setSeries(1, { color: '#FFFF00', label: { visible: true, contentType: 8 } }));
      r.series.full = JSON.parse(JSON.stringify(c.getInfo().config));
    } catch (e) { r.series.full = 'ERR ' + String(e).slice(0, 140); }

    // ---------- 2. 位置：直接 setPosition vs update ----------
    r.position = {};
    try {
      let c = await mk(T.Column);
      c.setPosition({ row: 3, column: 6 }); await wait(100);
      r.position.direct = { anchor: c.getInfo().anchor, position: c.getInfo().position };
    } catch (e) { r.position.direct = 'ERR ' + String(e).slice(0, 120); }
    try {
      let c = await mk(T.Column);
      const b = c.toBuilder(); b.setPosition({ row: 3, column: 6 });
      await c.update(b.build()); await wait(120);
      r.position.viaUpdate = { anchor: c.getInfo().anchor, position: c.getInfo().position };
    } catch (e) { r.position.viaUpdate = 'ERR ' + String(e).slice(0, 120); }
    try {
      let c = await mk(T.Column);
      const b = c.toBuilder(); b.setSize(500, 240);
      await c.update(b.build()); await wait(120);
      r.position.sizeViaUpdate = c.getInfo().size;
    } catch (e) { r.position.sizeViaUpdate = 'ERR ' + String(e).slice(0, 120); }

    // ---------- 3. exportImage ----------
    r.exportImage = {};
    try {
      let c = await mk(T.Column);
      const fn = c.exportImage;
      r.exportImage.typeof = typeof fn;
      if (typeof fn === 'function') {
        const res = await fn.call(c);
        r.exportImage.retType = typeof res;
        r.exportImage.retKeys = res && typeof res === 'object' ? Object.keys(res) : null;
        r.exportImage.retStr = typeof res === 'string' ? res.slice(0, 60) + ' len=' + res.length : String(res).slice(0, 80);
        r.exportImage.ctorName = res && res.constructor && res.constructor.name;
      }
    } catch (e) { r.exportImage.err = String(e).slice(0, 200); }
    try {
      let c = await mk(T.Column);
      if (typeof c.exportImage === 'function') {
        const res = await c.exportImage({ type: 'png' });
        r.exportImage.withOption = { retType: typeof res, len: typeof res === 'string' ? res.length : null, blob: !!(res && res.size) };
      }
    } catch (e) { r.exportImage.err2 = String(e).slice(0, 200); }

    // ---------- 4. 饼图/圆环 label ----------
    r.pie = {};
    const pieTry = async (label, key) => {
      try {
        let c = await mk(T.Pie, (b) => { if (typeof b.setSeries === 'function') b.setSeries(1, label); });
        r.pie[key + '_series'] = { cfg: JSON.parse(JSON.stringify(c.getInfo().config)).pie || null, s1: pick(c, 1) };
      } catch (e) { r.pie[key + '_series'] = 'ERR ' + String(e).slice(0, 120); }
    };
    await pieTry({ label: { visible: true, contentType: 16 } }, 'pct');
    for (const m of ['setPie', 'setPieLabel', 'setPieComposite']) {
      try {
        let c = await mk(T.Pie, (b) => { if (typeof b[m] === 'function') b[m]({ label: { visible: true, contentType: 16 }, position: 'outside' }); });
        r.pie[m] = { has: true, cfg: JSON.parse(JSON.stringify(c.getInfo().config)).pie || null };
      } catch (e) { r.pie[m] = { has: typeof sh.newChart(T.Pie)[m] === 'function', err: String(e).slice(0, 140) }; }
    }

    // ---------- 5. 3D ----------
    r.threeD = {};
    try {
      let c = await mk(T.Column3D, (b) => { if (typeof b.setThreeD === 'function') b.setThreeD({ perspective: 30 }); });
      r.threeD.setThreeD = { has: true, cfg: JSON.parse(JSON.stringify(c.getInfo().config)).threeD || null };
    } catch (e) { r.threeD.setThreeD = { has: typeof sh.newChart(T.Column3D).setThreeD === 'function', err: String(e).slice(0, 140) }; }

    // ---------- 6. 其它单点 ----------
    r.misc = {};
    try {
      let c = await mk(T.Column);
      r.misc.setName = typeof c.setName;
      if (typeof c.setName === 'function') { c.setName('MyChart'); await wait(80); r.misc.nameRead = c.getName ? c.getName() : '(no getName)'; }
    } catch (e) { r.misc.setName = 'ERR ' + String(e).slice(0, 120); }
    try {
      let c = await mk(T.Column);
      if (typeof c.setZOrder === 'function') {
        c.setZOrder(3); await wait(120);
        const i = c.getInfo();
        r.misc.zOrder = { infoKeys: Object.keys(i), anchor: i.anchor, z: i.zIndex };
      }
    } catch (e) { r.misc.zOrder = 'ERR ' + String(e).slice(0, 120); }
    // 趋势线字段真名
    try {
      let c = await mk(T.Line, (b) => b.setTrendline({ index: 1 }, { type: 'linear', showFormula: true, showR2: true }));
      const i = c.getInfo();
      r.misc.trendline = i.config.trendlines || null;
    } catch (e) { r.misc.trendline = 'ERR ' + String(e).slice(0, 140); }
    return r;
  });

  console.log(JSON.stringify(out, null, 1).slice(0, 14000));
  if (errs.length) console.log('页面错误', errs.slice(0, 4));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
