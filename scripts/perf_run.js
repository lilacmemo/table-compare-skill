// 公式性能基准跑批：node perf_run.js <sjs|univer>
// 流程：本地静态服务 → 无头浏览器 → 建工作簿写数据 → 挂起写公式(计时) → 强制全量重算(计时×3) → 改1输入重算 → 读核对样本
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { generate } = require('./perf_workload');

const engine = process.argv[2] || 'sjs';
const RUNS = Number(process.argv[3] || 3);
const PORT = 8437;
const ROOT = __dirname;
const MIME = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.json': 'application/json' };

function serve() {
    return new Promise((resolve) => {
        const srv = http.createServer((req, res) => {
            let p = decodeURIComponent(req.url.split('?')[0]);
            if (p === '/') p = '/' + (engine === 'sjs' ? 'perf_sjs.html' : 'perf_univer.html');
            const fp = path.join(ROOT, p);
            if (!fs.existsSync(fp) || !fs.statSync(fp).isFile()) { res.writeHead(404); res.end('nf'); return; }
            res.writeHead(200, { 'Content-Type': MIME[path.extname(fp)] || 'application/octet-stream' });
            fs.createReadStream(fp).pipe(res);
        });
        srv.listen(PORT, '127.0.0.1', () => resolve(srv));
    });
}

(async () => {
    const data = generate();
    console.log(`[workload] 公式总数=${data.formulaCount} Data=${data.dataValues.length}行x5式 Report=${data.reportFormulas.length} Chain=${data.chainFormulas.length} Dynamic=${data.dynFormulas.length}`);
    const srv = await serve();
    const browser = await chromium.launch({ headless: true, channel: 'msedge', args: ['--disable-gpu'] });
    const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
    page.on('console', (m) => { if (/error|Error/i.test(m.text())) console.log('[console]', m.text().slice(0, 200)); });
    page.on('pageerror', (e) => console.log('[pageerror]', String(e).slice(0, 300)));

    const url = engine === 'sjs' ? 'http://127.0.0.1:8437/perf_sjs.html' : 'http://127.0.0.1:8437/perf_univer.html';
    console.log('[nav]', url);
    await page.goto(url, { waitUntil: engine === 'univer' ? 'load' : 'domcontentloaded', timeout: 120000 });
    if (engine === 'univer') await page.waitForFunction('window.__ready === true', { timeout: 120000 });

    const version = await page.evaluate(() => {
        if (window.GC && GC.Spread) {
            // 探测版本：Workbook 实例或全局常见的版本属性
            try {
                const w = window.__spread && window.__spread.version;
                if (w) return w;
            } catch (e) {}
            try {
                const keys = Object.getOwnPropertyNames(GC.Spread.Sheets).filter(k => /version|product/i.test(k));
                for (const k of keys) { try { const v = typeof GC.Spread.Sheets[k] === 'function' ? GC.Spread.Sheets[k]() : GC.Spread.Sheets[k]; if (v && typeof v === 'string') return v; if (v && v.version) return v.version; } catch (e) {} }
            } catch (e) {}
            return 'SJS (npm @grapecity-software/spread-sheets, 与 19.2.5 对齐)';
        }
        return 'Univer 1.3.1';
    });
    console.log('[version]', version);

    if (engine === 'sjs') await page.evaluate(() => window.initPerfSjs());
    console.log('[setup] 写入基础数据...');
    await page.evaluate((d) => window.setupWorkload(d), data);

    console.log('[write] 写公式（挂起计算状态计时）...');
    const writeResult = await page.evaluate((d) => window.writeFormulas(d), data);
    console.log(`[write] 写入耗时=${writeResult.writeMs.toFixed(0)}ms 恢复计算(挂起期积压公式重算)=${(writeResult.resumeMs || 0).toFixed(0)}ms`);

    console.log(`[calc] 强制全量重算 ×${RUNS}...`);
    const calcTimes = await page.evaluate((n) => window.forceFullCalc(n), RUNS);
    console.log('[calc] 全量重算耗时(ms):', calcTimes.map(t => t.toFixed(0)).join(', '));

    console.log('[edit] 修改 1 个输入后强制全量重算...');
    const editTime = await page.evaluate(() => window.editAndRecalc());
    console.log('[edit] 修改+重算耗时(ms):', editTime.toFixed(0));

    console.log('[sample] 读取核对样本...');
    const samples = await page.evaluate((s) => window.readSamples(s), data.samples);
    for (const s of samples) console.log(`  ${s.sheet}!${s.cell} = ${s.value}`);

    const result = {
        engine, version, timestamp: new Date().toISOString(),
        workload: { formulaCount: data.formulaCount, dataRows: data.dataValues.length, reportFormulas: data.reportFormulas.length, chainFormulas: data.chainFormulas.length, dynFormulas: data.dynFormulas.length },
        writeMs: writeResult.writeMs, resumeMs: writeResult.resumeMs || 0, calcTimes, editAndRecalcMs: editTime, samples,
        env: 'Playwright headless msedge, 127.0.0.1 本地静态服务',
    };
    fs.writeFileSync(path.join(ROOT, `perf_result_${engine}.json`), JSON.stringify(result, null, 2));
    console.log('[done] 结果已写入 perf_result_' + engine + '.json');
    await browser.close();
    srv.close();
})().catch(e => { console.error('[fatal]', e); process.exit(1); });
