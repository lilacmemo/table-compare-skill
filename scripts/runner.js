// SpreadJS xlsx 图表导入还原实测
// 依赖安装（当前官方 scope）：
//   npm i playwright @grapecity-software/spread-sheets @grapecity-software/spread-sheets-shapes @grapecity-software/spread-sheets-charts @grapecity-software/spread-sheets-io
// 运行：node runner.js <样本.xlsx>
// 报告必须注明：SpreadJS 包版本号（npm ls @grapecity-software/spread-sheets）
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const XLSX = process.argv[2] || path.join(ROOT, '..', 'Excel原生图表效果.xlsx');
const PORT = 8931;

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
    const b64 = fs.readFileSync(XLSX).toString('base64');
    const browser = await chromium.launch();
    const page = await browser.newPage();
    page.on('pageerror', e => console.error('PAGE_ERROR:', e.message));
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');

    for (const preserve of [false, true]) {
        try {
            const r = await page.evaluate(([d, p]) => window.runImport(d, p), [b64, preserve]);
            const real = r.sheets.filter(s => !/evaluation/i.test(s.sheet));
            const evalSheets = r.sheets.filter(s => /evaluation/i.test(s.sheet));
            const total = real.reduce((s, x) => s + x.chartCount, 0);
            console.log(`\n=== preserveUnsupportedChart=${preserve} : 共 ${r.sheetCount} 表（含授权提示表 ${evalSheets.length} 个，已排除），有效图表存活 ${total}/63 ===`);
            console.log('  API 诊断:', JSON.stringify(r.apiInfo));
            for (const s of real) {
                const types = s.charts.map(c => c.type.replace(/^(xl|)/, '')).join(', ');
                console.log(`  [${s.sheet}] ${s.chartCount} 张${s.charts.length ? ' → ' + types : ''}`);
            }
            for (const s of r.sheets) {
                const types = s.charts.map(c => c.type).join(', ');
                console.log(`  [${s.sheet}] ${s.chartCount} 张${s.charts.length ? ' → ' + types : ''}`);
            }
        } catch (e) {
            console.error(`preserveUnsupportedChart=${preserve} 导入失败:`, e.message);
        }
        // 截图验证渲染效果
        for (const target of ['柱形图', '雷达气泡曲面图', '新式图表']) {
            try {
                await page.evaluate((name) => window.__gotoSheet(name), target);
                await page.waitForTimeout(500);
                await page.screenshot({
                    path: path.join(ROOT, `shot_p${preserve ? 1 : 0}_${target}.png`),
                    clip: { x: 0, y: 0, width: 1200, height: 800 }
                });
                if (preserve === false && target === '雷达气泡曲面图') {
                    // 向下滚动看气泡图/曲面图区域
                    await page.mouse.move(600, 400);
                    await page.mouse.wheel(0, 1500);
                    await page.waitForTimeout(500);
                    await page.screenshot({
                        path: path.join(ROOT, 'shot_p0_曲面图区域.png'),
                        clip: { x: 0, y: 0, width: 1200, height: 800 }
                    });
                }
            } catch (e) { console.error(`截图失败 ${target}:`, e.message); }
        }
    }
    await browser.close();
    server.close();
})();
