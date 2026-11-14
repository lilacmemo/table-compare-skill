// SJS 补测：股价图（按要求的列序供数）+ 组合图（series().add 逐系列组合）
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const PORT = 8939;
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
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1400,950'] });
    const page = await browser.newPage();
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');
    // 先加载一个样本以初始化 spread
    const b64 = fs.readFileSync(path.join(ROOT, '..', '批次1样式样本.xlsx')).toString('base64');
    await page.evaluate(d => window.runImport(d, false), b64);

    const result = await page.evaluate(() => {
        const C = GC.Spread.Sheets.Charts;
        const spread = window.__spread;
        spread.addSheet(spread.getSheetCount(), new GC.Spread.Sheets.Worksheet('__stock_test'));
        const ws = spread.getSheetFromName('__stock_test');
        const out = { stock: [], combo: null };

        // 股价数据：5 列布局（日期, 成交量, 开盘, 最高, 最低, 收盘）覆盖四种列序
        ws.setArray(0, 0, [
            ['日期', '成交量', '开盘', '最高', '最低', '收盘'],
            ['1日', 1000, 12, 15, 11, 14],
            ['2日', 1200, 14, 16, 13, 15],
            ['3日', 900, 15, 17, 14, 16],
            ['4日', 1500, 16, 18, 15, 17],
        ]);
        // HLC: 高-低-收 列序（用 C:E 即 最高/最低/收盘）
        const stockTests = [
            { name: 'stockHLC', type: C.ChartType.stockHLC, range: "'__stock_test'!A1:A5,C1:E5" },
            { name: 'stockOHLC', type: C.ChartType.stockOHLC, range: "'__stock_test'!A1:A5,C1:F5" },
            { name: 'stockVHLC', type: C.ChartType.stockVHLC, range: "'__stock_test'!A1:B5,C1:E5" },
            { name: 'stockVOHLC', type: C.ChartType.stockVOHLC, range: "'__stock_test'!A1:F5" },
        ];
        let seq = 0;
        for (const t of stockTests) {
            try {
                const ch = ws.charts.add('ST' + (++seq), t.type, 10, 10, 400, 280, t.range);
                out.stock.push({ name: t.name, ok: !!ch, readback: ch ? ch.chartType() : null });
            } catch (e) {
                out.stock.push({ name: t.name, ok: false, err: String(e && e.message || e).slice(0, 150) });
            }
        }
        // 组合图：先建柱形图，再逐系列追加折线系列
        ws.setArray(0, 8, [['类', '系列A', '系列B'], ['一', 30, 25], ['二', 44, 38], ['三', 26, 41]]);
        try {
            const ch = ws.charts.add('COMBO', C.ChartType.columnClustered, 500, 10, 400, 280, "'__stock_test'!I1:K4");
            ch.series().add({ chartType: C.ChartType.line, name: '系列B', xValues: "'__stock_test'!I2:I4", yValues: "'__stock_test'!J2:J4" });
            const types = ch.series().get().map(s => s.chartType);
            out.combo = { ok: true, seriesTypes: types, enumColumn: C.ChartType.columnClustered, enumLine: C.ChartType.line };
        } catch (e) {
            out.combo = { ok: false, err: String(e && e.message || e).slice(0, 200) };
        }
        return out;
    });

    fs.writeFileSync(path.join(ROOT, 'sjs_create_types2.json'), JSON.stringify(result, null, 1));
    console.log('股价图补测:', JSON.stringify(result.stock));
    console.log('组合图补测:', JSON.stringify(result.combo));
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
