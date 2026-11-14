// 批量跑「表 / 排序 / 筛选 / 数据验证」用例（46 条 × 双侧）
//   NODE_PATH=../poc-compare/node_modules node tools/run_cases_table.js [组别]
//   组别：A/B/C/D/E/F，不传则全跑
// 输出：tools/out/cases-table.json + 控制台分组摘要
//
// 判定：返回即 PASS，抛错即 FAIL。SJS 侧若 FAIL，先确认是「API 真不存在」还是
// 「用例里 API 名写错」——后者必须改用例，不能当成产品差异。

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8483);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const GROUP = (process.argv[2] || '').toUpperCase();

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end('not found'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  const pageErrs = [];
  page.on('pageerror', (e) => pageErrs.push(e.message));

  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__cmpRun && window.__cmpCases, null, { timeout: 90000 });
  await page.waitForTimeout(1500);

  const all = await page.evaluate(() => window.__cmpCases
    .filter((c) => /^T-/.test(c.id))
    .map((c) => ({ id: c.id, title: c.title, excel: c.excel, sjs: c.sjs, uni: c.uni })));

  const cases = GROUP ? all.filter((c) => c.id.startsWith('T-' + GROUP)) : all;
  console.log(`共 ${cases.length} 条用例（库内 ${all.length} 条）\n`);

  const results = [];
  for (const c of cases) {
    const one = { id: c.id, title: c.title, excel: c.excel };
    for (const side of ['sjs', 'uni']) {
      await page.evaluate(() => window.__cmpClear('batch'));
      await page.waitForTimeout(120);
      const r = await page.evaluate(({ side, code }) => window.__cmpRun(side, code), { side, code: c[side] });
      one[side] = r;
      await page.waitForTimeout(60);
    }
    results.push(one);
    const mark = (r) => (r.ok ? 'PASS' : 'FAIL');
    console.log(
      `${c.id.padEnd(6)} SJS:${mark(one.sjs)}  UNI:${mark(one.uni)}  ${c.title}`
    );
    if (!one.sjs.ok) console.log(`        ↳ SJS  ${one.sjs.err}`);
    if (!one.uni.ok) console.log(`        ↳ UNI  ${one.uni.err}`);
  }

  // 分组统计
  const stat = {};
  for (const r of results) {
    const g = r.id[2];
    stat[g] = stat[g] || { sjsPass: 0, sjsFail: 0, uniPass: 0, uniFail: 0, n: 0 };
    stat[g].n++;
    r.sjs.ok ? stat[g].sjsPass++ : stat[g].sjsFail++;
    r.uni.ok ? stat[g].uniPass++ : stat[g].uniFail++;
  }

  console.log('\n组别  SJS PASS/FAIL    Univer PASS/FAIL');
  let ts = 0, tf = 0, us = 0, uf = 0;
  for (const g of Object.keys(stat).sort()) {
    const s = stat[g];
    ts += s.sjsPass; tf += s.sjsFail; us += s.uniPass; uf += s.uniFail;
    console.log(`${g}     ${String(s.sjsPass).padStart(2)} / ${String(s.sjsFail).padStart(2)}`
      + `          ${String(s.uniPass).padStart(2)} / ${String(s.uniFail).padStart(2)}`);
  }
  console.log(`合计  ${ts} / ${tf}          ${us} / ${uf}`);

  if (pageErrs.length) console.log('\nPAGEERR:', pageErrs.slice(0, 5).join(' | '));

  const file = path.join(OUT, 'cases-table.json');
  fs.writeFileSync(file, JSON.stringify({ group: GROUP || 'ALL', at: new Date().toISOString(), stat, results }, null, 2), 'utf8');
  console.log('\n写出', file);

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FATAL', e); process.exit(1); });
