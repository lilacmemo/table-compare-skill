// 诊断 9：A10 exportImage 在批量跑批中为什么返回 undefined（单独跑却正常）
//   node _diag_pro_chart9.js
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const SRC = fs.readFileSync(path.join(__dirname, 'probe_pro_chart_api.js'), 'utf8');
// 复用探针里的 PAGE_MAIN
const body = SRC.slice(SRC.indexOf('function PAGE_MAIN()'), SRC.indexOf('(async () => {'));
const PAGE_MAIN = eval('(' + body.trim() + ')');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const go = async () => {
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  if (!(await go())) { console.log('not ready'); await browser.close(); process.exit(1); }
  await page.waitForTimeout(2500);
  await page.evaluate(`(${PAGE_MAIN.toString()})()`);
  console.log('init', JSON.stringify(await page.evaluate(() => window.__ProChartProbe.init())));

  const out = await page.evaluate(async () => {
    const P = window.__ProChartProbe;
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const r = {};
    // 按顺序跑 A1~A9
    for (const id of ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9']) {
      const spec = P.SPEC.find((s) => s.id === id);
      r['run_' + id] = await P.runOne(spec);
    }
    // 紧接着跑 A10
    r['run_A10'] = await P.runOne(P.SPEC.find((s) => s.id === 'A10'));

    // 手工再试：直接建图 + 5 次尝试
    const U = window.univerAPI, T = U.Enum.ChartTypeString;
    const s = window.__chartSheet;
    const next = async (label) => {
      for (const c of s.getCharts()) { try { await c.remove(); } catch (e) {} }
      const c = await s.insertChart(s.newChart(T.Column).setSource({ range: 'A1:D7' }).setPosition({ row: 1, column: 4 }).setSize(480, 320).build());
      const seq = [];
      for (let i = 0; i < 5; i++) { await wait(400); const res = await c.exportImage(); seq.push(typeof res === 'string' ? 'str:' + res.length : String(res)); }
      return { label, active: U.getActiveWorkbook().getActiveSheet().getSheetName(), seq };
    };
    r.manual1 = await next('after-batch');
    // 重新激活工作表 + 滚动
    try { U.getActiveWorkbook().setActiveSheet(s); await wait(400); } catch (e) {}
    try { s.scrollToCell ? s.scrollToCell(0, 4) : null; } catch (e) {}
    r.manual2 = await next('reactivated');
    return r;
  });
  console.log(JSON.stringify(out, null, 1).slice(0, 3000));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
