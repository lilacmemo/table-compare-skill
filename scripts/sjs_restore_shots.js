// sjs_restore_shots.js —— SpreadJS 导入基线样本"原始状态"逐图导出 v3
// 用 chart.toImageDataURL() 直接导出每个图表对象自身图像（与 Excel COM Chart.Export 同构，
// 叠放图、视口外图都能完整导出，无需隐藏/裁切）。
// 用法: node sjs_restore_shots.js <xlsx路径> <prefix>
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const XLSX = process.argv[2] || path.join(ROOT, '..', '批次1样式样本.xlsx');
const PREFIX = process.argv[3] || 'b1';
const PORT = 8943;
const MIME = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css' };

const server = http.createServer((req, res) => {
    const p = path.join(ROOT, decodeURIComponent(req.url.split('?')[0]));
    fs.readFile(p, (err, data) => {
        if (err) { res.writeHead(404); res.end('nf'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(data);
    });
});

(async () => {
    await new Promise(r => server.listen(PORT, r));
    const b64 = fs.readFileSync(XLSX).toString('base64');
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 900 }, deviceScaleFactor: 2 });
    page.on('pageerror', e => console.error('PAGE_ERROR:', e.message));
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');

    const info = await page.evaluate(async (b64) => {
        const r = await window.runImport(b64, true);
        const out = [];
        for (let i = 0; i < r.sheetCount; i++) {
            const sheet = window.__spread.getSheet(i);
            const charts = sheet.charts.all();
            if (!charts.length) { out.push({ name: sheet.name(), charts: [] }); continue; }
            out.push({
                name: sheet.name(),
                charts: charts.map(c => {
                    try { return { name: c.name(), x: c.x(), y: c.y(), w: c.width(), h: c.height() }; }
                    catch (e) { return { name: c.name(), err: String(e) }; }
                }),
            });
        }
        return { sheets: out, sheetCount: r.sheetCount };
    }, b64);

    const posDump = {};
    for (let i = 0; i < info.sheets.length; i++) {
        const s = info.sheets[i];
        if (!s.charts.length) { console.log('skip:', s.name); continue; }
        const tag = String(i).padStart(2, '0') + '_' + s.name.replace(/[^\w]/g, '');
        const entries = [];
        for (let k = 0; k < s.charts.length; k++) {
            const c = s.charts[k];
            if (c.err) continue;
            const dataUrl = await page.evaluate(async ({ sheetName, idx }) => {
                const sheet = window.__spread.getSheetFromName(sheetName);
                const ch = sheet.charts.all()[idx];
                try {
                    if (typeof ch.toImageDataURL === 'function') return await ch.toImageDataURL();
                    if (typeof ch.toImageSrcAsync === 'function') return await ch.toImageSrcAsync();
                    if (typeof ch.toImageSrc === 'function') return ch.toImageSrc();
                } catch (e) { return 'ERR:' + e.message; }
                return null;
            }, { sheetName: s.name, idx: k });
            if (!dataUrl || dataUrl.startsWith('ERR:') || !dataUrl.startsWith('data:image')) {
                console.log(`  k${k + 1} ${c.name}: FAIL ${String(dataUrl).slice(0, 80)}`);
                continue;
            }
            const file = `restore_${PREFIX}_${tag}_k${k + 1}.png`;
            fs.writeFileSync(path.join(ROOT, file), Buffer.from(dataUrl.split(',')[1], 'base64'));
            entries.push({ file, name: c.name, xPx: +(c.x || 0).toFixed(1), yPx: +(c.y || 0).toFixed(1), wPx: +(c.w || 0).toFixed(1), hPx: +(c.h || 0).toFixed(1) });
            console.log(`  k${k + 1} ${c.name}: ok`);
        }
        posDump[s.name] = entries;
        console.log(`sheet ${s.name}: ${entries.length}/${s.charts.length} exported`);
    }
    fs.writeFileSync(path.join(ROOT, `restore_pos_${PREFIX}.json`), JSON.stringify(posDump, null, 1));
    await browser.close();
    server.close();
    console.log('DONE v3 prefix=' + PREFIX);
    process.exit(0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
