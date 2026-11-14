// 诊断：Univer 主题 / 调色板 的 set+get 往返（用户质疑 J5~J7 判定）
//   node _diag_theme_palette.js
const { chromium } = require('playwright');
const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  page.on('pageerror', (e) => console.log('[pageerror]', String(e.message).slice(0, 200)));

  const go = async () => {
    await page.goto(URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  if (!(await go())) { console.log('univerAPI 未就绪'); await browser.close(); process.exit(1); }
  await page.waitForTimeout(3000);

  const out = await page.evaluate(async () => {
    const U = window.univerAPI; const T = U.Enum.ChartTypeString; const E = U.Enum;
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const wb = U.getActiveWorkbook();
    let s = wb.getSheetByName('__diag');
    if (!s) s = wb.insertSheet('__diag');
    wb.setActiveSheet(s);
    const vals = [['Cat', 'S1', 'S2', 'S3']];
    for (let i = 1; i <= 6; i++) vals.push(['C' + i, 10 + i * 3, 20 + i * 5, 30 - i]);
    s.getRange('A1:D7').setValues(vals);
    await wait(300);

    const r = {};
    r.themeNames = Object.keys(E.ChartThemeName || {}).map((k) => k + '=' + E.ChartThemeName[k]);
    r.hasRegisterTheme = typeof U.registerTheme;

    const fresh = async (build) => {
      for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
      const b = s.newChart(T.Column).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320);
      if (build) build(b);
      const c = await s.insertChart(b.build());
      await wait(200);
      return c;
    };

    // ① 默认主题读回
    let c = await fresh();
    r['1_default_config_theme'] = JSON.stringify(c.getInfo().config.theme);
    r['1_default_config_palette'] = JSON.stringify(c.getInfo().config.palette);
    r['1_config_keys'] = Object.keys(c.getInfo().config);

    // ② setTheme 后读回
    c = await fresh((b) => b.setTheme(E.ChartThemeName.UniverGradient1));
    const i2 = c.getInfo();
    r['2_after_setTheme'] = JSON.stringify(i2.config.theme);
    r['2_config_palette'] = JSON.stringify(i2.config.palette);

    // ③ 换成另一个主题
    c = await fresh((b) => b.setTheme(E.ChartThemeName.Univer3));
    r['3_after_setTheme_univer3'] = JSON.stringify(c.getInfo().config.theme);

    // ④ 已存在图表上 setTheme（实例方法）
    c.setTheme(E.ChartThemeName.Univer5);
    await wait(200);
    r['4_instance_setTheme'] = JSON.stringify(c.getInfo().config.theme);

    // ⑤ setPalette 后读回
    c = await fresh((b) => b.setPalette(['#FF0000', '#00FF00', '#0000FF']));
    r['5_after_setPalette'] = JSON.stringify(c.getInfo().config.palette);

    // ⑥ 主题 + 调色板同时存在
    c = await fresh((b) => b.setTheme(E.ChartThemeName.Univer2).setPalette(['#123456', '#654321']));
    const i6 = c.getInfo();
    r['6_theme_and_palette'] = JSON.stringify({ theme: i6.config.theme, palette: i6.config.palette });

    // ⑦ clearPalette / clearTheme
    try { c.clearPalette(); await wait(150); r['7_after_clearPalette'] = JSON.stringify(c.getInfo().config.palette); } catch (e) { r['7_after_clearPalette'] = 'throw:' + e.message; }
    try { c.clearTheme(); await wait(150); r['8_after_clearTheme'] = JSON.stringify(c.getInfo().config.theme); } catch (e) { r['8_after_clearTheme'] = 'throw:' + e.message; }

    // ⑨ 系列实际颜色是否随主题变化（取 series 的 colors/backColor）
    for (const c2 of s.getCharts()) { try { await c2.remove(); } catch (e) {} }
    const t1 = await s.insertChart(s.newChart(T.Column).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(400, 260).setTheme(E.ChartThemeName.Univer1).build());
    await wait(300);
    r['9_univer1_series'] = JSON.stringify((t1.getInfo().config.series || []).map((x) => x.colors || x.color || x.fill));

    for (const c2 of s.getCharts()) { try { await c2.remove(); } catch (e) {} }
    return r;
  });

  for (const k of Object.keys(out)) console.log(k, '=', out[k]);
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
