// SpreadJS 图表 API 探针跑批器
//   node tools/run_probe_chart_sjs.js            → tools/out/chart-api-sjs.json
// 依赖 poc-compare/node_modules 里的 playwright（NODE_PATH 指定）
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = 8466;
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.ttf': 'font/ttf', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
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
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e.message).slice(0, 200)));
  await page.goto(`http://127.0.0.1:${PORT}/chart-probe.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__chartReady && window.ApiProbeChartSjs, null, { timeout: 60000 });

  const t0 = Date.now();
  const res = await page.evaluate(async () => await window.ApiProbeChartSjs.runAll());
  const file = path.join(OUT, 'chart-api-sjs.json');
  fs.writeFileSync(file, JSON.stringify(res, null, 1), 'utf8');

  const items = res.items || [];
  const pass = items.filter((i) => i.ok).length;
  console.log('用时', ((Date.now() - t0) / 1000).toFixed(1) + 's', '总项', items.length);
  console.log('SJS  ', JSON.stringify({ pass, fail: items.length - pass, rate: (pass / items.length * 100).toFixed(1) + '%' }));
  console.log('--- FAIL ---');
  for (const i of items.filter((x) => !x.ok)) {
    console.log(i.id, i.name, '|', i.reason || ('got=' + i.got + ' exp=' + i.exp), '|', i.note || '');
  }
  const byGroup = {};
  for (const i of items) { byGroup[i.g] = byGroup[i.g] || { pass: 0, fail: 0 }; byGroup[i.g][i.ok ? 'pass' : 'fail']++; }
  console.log('--- 分组 ---', JSON.stringify(byGroup));
  if (errs.length) console.log('页面错误', errs.slice(0, 5));
  console.log('写出', file);

  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
