// sjs_v2.js -- 用 v2 语义用例在同一数据区上跑 SpreadJS，输出 probe_sjs_v2.json
// 与 excel_case_eval.ps1 完全共享 data_block_v2.json / cases_v2.json，保证三方输入一致。
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const http = require('http');

const HERE = __dirname;
const CASES = JSON.parse(fs.readFileSync(path.join(HERE, 'cases_v2.json'), 'utf8'));
const DATA = JSON.parse(fs.readFileSync(path.join(HERE, 'data_block_v2.json'), 'utf8'));
const OUT = path.join(HERE, 'probe_sjs_v2.json');
const PORT = 8971;

const server = http.createServer((req, res) => {
    const p = path.join(HERE, decodeURIComponent(req.url.split('?')[0]));
    fs.readFile(p, (e, d) => {
        if (e) { res.writeHead(404); res.end(); return; }
        const ext = path.extname(p);
        res.writeHead(200, { 'Content-Type': ext === '.html' ? 'text/html' : 'application/octet-stream' });
        res.end(d);
    });
});

(async () => {
    await new Promise(r => server.listen(PORT, r));
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    page.on('pageerror', e => console.error('PAGEERR', String(e).slice(0, 300)));
    await page.goto(`http://localhost:${PORT}/formula_probe.html`);
    await page.waitForFunction(() => typeof window.runFinal === 'function');
    const ver = await page.evaluate(() => window.sjsVersion());
    const t0 = Date.now();
    const r = await page.evaluate(async ([cases, data]) => await window.runFinal(cases, data), [CASES, DATA]);
    console.log('done in', ((Date.now() - t0) / 1000).toFixed(1), 's | SJS', ver, '| allowDynamicArray', r.allowDynamicArray);
    r.sjsVersion = ver;
    fs.writeFileSync(OUT, JSON.stringify(r, null, 1));
    console.log('setErrors:', r.setErrors.length, '| rows:', r.rows.length);
    const b = {};
    for (const x of r.rows) {
        const k = x.valueType === 'error' ? 'ERR ' + x.value.err : x.valueType;
        b[k] = (b[k] || 0) + 1;
    }
    console.log('buckets:', JSON.stringify(Object.entries(b).sort((a, c) => c[1] - a[1])));
    if (r.setErrors.length) console.log('setErr sample:', JSON.stringify(r.setErrors.slice(0, 12)));
    await browser.close(); server.close(); process.exit(0);
})().catch(e => { console.error('FATAL', e); server.close(); process.exit(1); });
