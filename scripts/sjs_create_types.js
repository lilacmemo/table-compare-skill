// SpreadJS 图表类型实测：逐个枚举 ChartType 用代码新建，统计可创建类型 + 降级映射
// 用法：node sjs_create_types.js
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const FILE = process.argv[2] || path.join(ROOT, '..', '批次1样式样本.xlsx');
const PORT = 8937;

const MIME = { '.html': 'text/html', '.js': 'application/javascript' };
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
    const b64 = fs.readFileSync(FILE).toString('base64');
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1400,950'] });
    const page = await browser.newPage();
    page.on('pageerror', e => console.error('PAGE_ERROR:', e.message));
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');
    await page.evaluate(d => window.runImport(d, false), b64);

    const result = await page.evaluate(() => {
        const C = GC.Spread.Sheets.Charts;
        const spread = window.__spread;
        // 用一个全新空表承载新建，避免与导入图混淆
        let ws;
        try { ws = spread.getSheetFromName('__create_test'); } catch (e) {}
        if (!ws) {
            spread.addSheet(spread.getSheetCount(), new GC.Spread.Sheets.Worksheet('__create_test'));
            ws = spread.getSheetFromName('__create_test');
        }
        ws.setArray(0, 0, [['类别', '系列A', '系列B'], ['一月', 30, 25], ['二月', 44, 38], ['三月', 26, 41], ['四月', 52, 33]]);
        // 枚举 dump
        const enumDump = {};
        for (const k of Object.keys(C.ChartType)) {
            if (typeof C.ChartType[k] === 'number') enumDump[k] = C.ChartType[k];
        }
        const results = [];
        let seq = 0;
        for (const [name, val] of Object.entries(enumDump)) {
            let created = null, err = null, readbackType = null, readbackName = null;
            try {
                created = ws.charts.add('T' + (++seq), val, 10, 10, 400, 300, "'__create_test'!A1:C5");
                if (created) {
                    readbackType = created.chartType();
                    // 数字反查枚举名
                    for (const [n2, v2] of Object.entries(enumDump)) if (v2 === readbackType) readbackName = n2;
                }
            } catch (e) { err = String(e && e.message || e); }
            results.push({ enumName: name, requested: val, created: !!created, readbackType, readbackName, downgraded: created ? readbackType !== val : null, err });
        }
        return { enumCount: Object.keys(enumDump).length, results };
    });

    fs.writeFileSync(path.join(ROOT, 'sjs_create_types.json'), JSON.stringify(result, null, 1));
    const ok = result.results.filter(r => r.created);
    const fail = result.results.filter(r => !r.created);
    const dg = ok.filter(r => r.downgraded);
    console.log(`ChartType 枚举总数: ${result.enumCount}`);
    console.log(`可创建: ${ok.length} | 创建失败: ${fail.length} | 创建成功但读回类型变化(降级映射): ${dg.length}`);
    if (fail.length) console.log('创建失败:', fail.map(r => r.enumName).join(', '));
    if (dg.length) console.log('降级映射:', dg.map(r => `${r.enumName} -> ${r.readbackName}`).join(', '));
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
