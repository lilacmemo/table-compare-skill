// 诊断：chart.colorScheme() 到底能不能写？【本轮用户质疑点】
//   文档说 colorScheme(value?) 是「获取或设置」，但探针里传 1/2/3 读回 undefined
//   → 怀疑是探针写法错（传数字 vs 传 ColorScheme 对象）
// 跑法：node tools/diag_color_scheme.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PORT = 8471;
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };

(async () => {
  const srv = await new Promise((resolve) => {
    const s = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'chart-probe.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) return res.writeHead(404).end('nf');
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    s.listen(PORT, '127.0.0.1', () => resolve(s));
  });

  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => console.log('[pageerror]', String(e.message).slice(0, 300)));
  await page.goto(`http://127.0.0.1:${PORT}/chart-probe.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__chartReady, null, { timeout: 60000 });

  const out = await page.evaluate(async () => {
    const C = GC.Spread.Sheets.Charts;
    const sp = window.__spread;
    const sh = sp.getSheetFromName(sp.getSheet(0).name()) || sp.getSheet(0);
    sh.setArray(0, 0, [
      ['Cat', 'S1', 'S2'],
      ['C1', 11, 21], ['C2', 12, 22], ['C3', 13, 23], ['C4', 14, 24],
    ]);
    const r = {};

    // 枚举面
    r.colorSchemesKeys = Object.keys(C.ColorSchemes);
    r.colorSchemeCount = r.colorSchemesKeys.length;
    r.colorRuleKeys = Object.keys(C.ColorRule);
    r.colorRuleVals = C.ColorRule;
    r.palette2_sample = C.ColorSchemes.colorfulPalette2;

    const mk = () => sh.charts.add('c' + Math.random().toString(36).slice(2), C.ChartType.columnClustered, 20, 20, 520, 320, 'A1:C5');

    // ① 初始读回
    let ch = mk();
    r['1_initial'] = JSON.stringify(ch.colorScheme());

    // ② 内置方案：传 ColorSchemes.colorfulPalette2
    ch.colorScheme(C.ColorSchemes.colorfulPalette2);
    r['2_after_set_palette2'] = JSON.stringify(ch.colorScheme());
    r['2_equal_to_palette2'] = JSON.stringify(ch.colorScheme()) === JSON.stringify(C.ColorSchemes.colorfulPalette2);

    // ③ 换 colorfulPalette4 再读
    ch.colorScheme(C.ColorSchemes.colorfulPalette4);
    r['3_after_set_palette4'] = JSON.stringify(ch.colorScheme());
    r['3_equal_to_palette4'] = JSON.stringify(ch.colorScheme()) === JSON.stringify(C.ColorSchemes.colorfulPalette4);

    // ④ 自定义方案
    const mine = [C.ColorRule.cycle, ['#FF0000', '#00AA00', '#0000FF']];
    ch.colorScheme(mine);
    r['4_after_set_custom'] = JSON.stringify(ch.colorScheme());
    r['4_equal_to_custom'] = JSON.stringify(ch.colorScheme()) === JSON.stringify(mine);

    // ⑤ 再读一次（有无缓存/延时）
    await new Promise((s) => setTimeout(s, 300));
    r['5_read_again'] = JSON.stringify(ch.colorScheme());

    // ⑥ 序列化后是否保留（toJSON → 重建）
    const raw = sp.toJSON();
    const json = typeof raw === 'string' ? JSON.parse(raw) : raw;
    r['6_json_sheet_keys'] = JSON.stringify(Object.keys(json)).slice(0, 200);
    const sh0 = json.sheets && (Array.isArray(json.sheets) ? json.sheets[0] : json.sheets[0]);
    r['6_json_sheet0_keys'] = sh0 ? JSON.stringify(Object.keys(sh0)).slice(0, 300) : 'no sh0';
    r['6_json_charts'] = sh0 && sh0.charts ? JSON.stringify(sh0.charts.map((c) => ({ name: c.name, colorScheme: c.colorScheme }))).slice(0, 400) : 'no charts in json';

    // ⑦ 新建图表时通过 add 的第 9 参传入
    const ch2 = sh.charts.add('named2', C.ChartType.columnClustered, 20, 380, 520, 320, 'A1:C5', C.RowCol.columns, C.ColorSchemes.colorfulPalette3);
    r['7_add_with_scheme'] = JSON.stringify(ch2.colorScheme());
    r['7_equal_to_palette3'] = JSON.stringify(ch2.colorScheme()) === JSON.stringify(C.ColorSchemes.colorfulPalette3);

    // ⑧ 系列颜色是否随之变化？（文档说不会，需要手动 reset）
    const sBefore = ch.series().get().map((s) => s.backColor);
    const schemeNow = ch.colorScheme();
    const s2 = ch.series().get();
    for (let i = 0; i < s2.length; i++) {
      s2[i].backColor = C.getColor(schemeNow, i, s2.length);
      ch.series().set(i, s2[i]);
    }
    const sAfter = ch.series().get().map((s) => s.backColor);
    r['8_series_before'] = JSON.stringify(sBefore);
    r['8_series_after_getColor'] = JSON.stringify(sAfter);
    r['8_getColor_0_3'] = C.getColor(mine, 0, 3);
    r['8_getColor_1_3'] = C.getColor(mine, 1, 3);

    // ⑨ 主题：工作表级 currentTheme 能不能影响图表？是否等价于图表主题
    r['9_sheet_theme'] = String(sh.currentTheme && sh.currentTheme().name && sh.currentTheme().name());
    r['9_has_chart_theme'] = typeof ch.theme;
    r['9_has_chart_setTheme'] = typeof ch.setTheme;

    // ⑩ 还原：设为 undefined / null 会怎样
    try { ch.colorScheme(null); r['10_set_null'] = JSON.stringify(ch.colorScheme()); } catch (e) { r['10_set_null'] = 'throw:' + e.message; }

    return r;
  });

  for (const k of Object.keys(out)) console.log(k, '=', out[k]);
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
