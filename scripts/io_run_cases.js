// 通用用例跑批：打开对比台 → 跑指定模块 → 落盘 JSON
//   NODE_PATH=../poc-compare/node_modules node tools/run_cases.js charts
//   node tools/run_cases.js charts --no-sample     不导入样本（纯 API 用例）
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const MOD = process.argv[2] || 'charts';
const NO_SAMPLE = process.argv.includes('--no-sample');
const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html';
const OUT = path.join(__dirname, '..', 'out', 'data');

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1680, height: 1000 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push('PAGEERROR ' + String(e).slice(0, 200)));
  page.on('console', (m) => { if (m.type() === 'error') errs.push(m.text().slice(0, 200)); });

  const url = BASE + '?module=' + MOD + (NO_SAMPLE ? '&autoload=0' : '');
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 120000 });
  for (let i = 0; i < 120; i++) {
    const ok = await page.evaluate(() => window.APP && APP.booted && APP.ready.sjs && APP.ready.uni);
    if (ok) break;
    await page.waitForTimeout(1000);
  }
  const boot = await page.evaluate(() => ({ sjs: APP.ready.sjs, uni: APP.ready.uni, anon: APP.anonUser || null }));
  console.log('boot', JSON.stringify(boot));

  // 等样本导入完成（图表探针会在工作簿里加临时表，必须先让导入落定）
  if (!NO_SAMPLE) {
    for (let i = 0; i < 90; i++) {
      const done = await page.evaluate(() => !!(APP.lastImport && APP.lastImport.sjs && APP.lastImport.uni));
      if (done) break;
      await page.waitForTimeout(1000);
    }
    const imp = await page.evaluate(() => APP.lastImport);
    console.log('import', JSON.stringify(imp));
  }

  const t0 = Date.now();
  const summary = await page.evaluate(async (m) => await APP.runCases(m), MOD);
  const secs = ((Date.now() - t0) / 1000).toFixed(1);
  if (!summary) { console.log('没有结果'); await browser.close(); process.exit(1); }

  const payload = {
    module: MOD,
    at: new Date().toISOString(),
    env: {
      sjs: await page.evaluate(() => (window.GC && GC.Spread && GC.Spread.Sheets.Version) ? GC.Spread.Sheets.Version.getProductVersion() : 'unknown'),
      univer: '1.3.1 Pro (local bundle)',
      browser: 'msedge headless',
    },
    ...summary,
  };
  const file = path.join(OUT, MOD + '-cases.json');
  fs.writeFileSync(file, JSON.stringify(payload, null, 1), 'utf8');
  console.log('用时', secs + 's', '总项', summary.total);
  console.log('SpreadJS', JSON.stringify(summary.sjs), 'Univer', JSON.stringify(summary.uni));
  console.log('FAIL 明细：');
  for (const r of summary.results) {
    if (r.sjs !== 'PASS' || r.uni !== 'PASS') console.log(' ', r.id, r.name, '|', r.sjs, r.uni, '|', r.note.slice(0, 160));
  }
  console.log('写出', file);
  if (errs.length) console.log('页面错误', errs.slice(0, 5));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
