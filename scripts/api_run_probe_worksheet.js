// 工作表 / 工作簿层 API 探针跑批器
//   node tools/run_probe_ws.js            → tools/out/api-ws.json
// 依赖：poc-compare/node_modules 里的 playwright（NODE_PATH 指定）
// 浏览器：本机无 Chrome，用 msedge（同 Chromium）

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = 8462;
const CHANNEL = process.env.PW_CHANNEL || 'msedge';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.ttf': 'font/ttf',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      if (!p.startsWith(ROOT)) { res.writeHead(403).end(); return; }
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
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e.message)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });

  await page.waitForFunction(
    () => window.__spread && window.__wb && window.univerAPI && window.ApiProbeWs,
    null, { timeout: 90000 }
  );
  await page.waitForTimeout(1500);

  // 环境自检：确认探针要用的 Facade 入口真的挂上了（缺插件时会 undefined）
  const env = await page.evaluate(() => {
    const wb = window.__wb, sh = wb.getActiveSheet();
    const has = (o, k) => typeof o[k] === 'function';
    return {
      sjsVersion: window.GC.Spread.Sheets.getVersion ? window.GC.Spread.Sheets.getVersion() : '?',
      uniSheets: wb.getNumSheets(),
      findReplace: has(window.univerAPI, 'createTextFinderAsync'),
      crosshair: has(window.univerAPI, 'setCrosshairHighlightEnabled'),
      perm: !!(sh.getWorksheetPermission && sh.getWorksheetPermission()),
      outline: has(sh, 'addRowOutline'),
      sparkline: has(sh, 'addSparkline'),
      shape: has(sh, 'insertShape'),
      print: has(wb, 'updatePrintConfig'),
      specCount: window.ApiProbeWs.SPEC.length,
    };
  });
  console.log('环境自检', JSON.stringify(env));

  const t0 = Date.now();
  const groups = process.argv[2] || '';
  const QUIET = !!process.env.WS_QUIET;

  // 逐项打点：页面内 __wsProgress 由 runAll 每项更新，这里每秒轮询一次，
  // 挂起时能立刻看出卡在哪个 ID（超时整体 kill 后才看得出）
  const poll = setInterval(async () => {
    if (QUIET) return;
    try {
      const p = await page.evaluate(() => window.__wsProgress || null);
      if (p) process.stdout.write(`\r  进度 ${p.done}/${p.total} 当前 ${p.at || '-'}   `);
    } catch (e) {}
  }, 1500);

  const result = await page.evaluate(
    async (g) => await window.ApiProbeWs.runAll({ groups: g || undefined }),
    groups
  );
  clearInterval(poll);
  process.stdout.write('\n');
  const file = path.join(OUT, 'api-ws.json');
  fs.writeFileSync(file, JSON.stringify(result, null, 2), 'utf8');

  const items = result.items || [];
  const tally = (side) => {
    let pass = 0, noApi = 0, err = 0, mismatch = 0, strictFail = 0;
    for (const i of items) {
      const s = i[side] || {};
      if (s.noApi) noApi++;
      else if (s.setOk === false) err++;
      else if (s.verdict === 'FAIL') strictFail++;
      else { pass++; if (!s.match) mismatch++; }
    }
    return { pass, noApi, err, strictFail, mismatch, total: items.length };
  };
  console.log('用时', ((Date.now() - t0) / 1000).toFixed(1) + 's');
  console.log('SJS   ', JSON.stringify(tally('sjs')));
  console.log('Univer', JSON.stringify(tally('uni')));

  const bad = items.filter((i) => (i.sjs && i.sjs.setOk === false) || (i.uni && i.uni.setOk === false));
  if (bad.length) {
    console.log('--- 抛错项（前 25）---');
    for (const b of bad.slice(0, 25)) {
      console.log(b.id, b.name, '| sjs:', b.sjs.err || '-', '| uni:', b.uni.err || '-');
    }
  }
  const sf = items.filter((i) => (i.sjs && i.sjs.verdict === 'FAIL') || (i.uni && i.uni.verdict === 'FAIL'));
  if (sf.length) {
    console.log('--- strict 判 FAIL（有入口但语义未达成）---');
    for (const m of sf) {
      console.log(m.id, m.name, '| sjs:', m.sjs.verdict, m.sjs.exp, '/', m.sjs.got, '| uni:', m.uni.verdict, m.uni.exp, '/', m.uni.got);
    }
  }
  const mis = items.filter((i) => (i.sjs && i.sjs.setOk && !i.sjs.match && i.sjs.verdict !== 'FAIL') || (i.uni && i.uni.setOk && !i.uni.match && i.uni.verdict !== 'FAIL'));
  if (mis.length) {
    console.log('--- 读回不一致（非 strict，只作提示，前 30）---');
    for (const m of mis.slice(0, 30)) {
      console.log(m.id, m.name, '| sjs exp=', m.sjs.exp, 'got=', m.sjs.got, '| uni exp=', m.uni.exp, 'got=', m.uni.got);
    }
  }
  if (errs.length) console.log('页面错误：', errs.slice(0, 6));
  console.log('写出', file);

  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
