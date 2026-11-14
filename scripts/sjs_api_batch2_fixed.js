// sjs_api_batch2_fixed.js —— SJS 侧与 univer_api_batch2_fixed.js 完全同口径的复刻：
// 导入批次2 → 对 09高级元素表 k1~k4 重放相同修改（J1/J2 次轴切换、J3a 趋势线改红+公式R²、J3b period=3、k4 无代码入口）
// → chart.toImageDataURL() 逐图导出 → _sjs_fixed_k1..4.png（与 Univer _fixed_k*.png 逐图对应）
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const ROOT = __dirname;
const FILE = path.join(ROOT, '..', '批次2系列与高级元素样本.xlsx');
const PORT = 8935;

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
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 900 }, deviceScaleFactor: 2 });
    page.on('pageerror', e => console.error('PAGE_ERROR:', e.message));
    await page.goto(`http://localhost:${PORT}/test.html`);
    await page.waitForFunction(() => typeof window.runImport === 'function');
    await page.evaluate(d => window.runImport(d, false), b64);
    await page.waitForTimeout(2000);

    const log = await page.evaluate(() => {
        const GC_ = GC.Spread.Sheets, C = GC_.Charts;
        const spread = window.__spread;
        const sheet = spread.getSheetFromName('09高级元素');
        const charts = sheet.charts.all();
        const out = { total: charts.length, steps: [] };
        const spaceOf = c => { const j = c.toJSON(); return j.shapeData?.graphicFrame?.graphic?.graphicData?.chart?.chartSpace?.chart; };
        const enumName = (en, v) => { try { return Object.keys(en).find(k => en[k] === v) || String(v); } catch (e) { return String(v); } };

        // 与 Univer 脚本一致：按谓词定位 k1~k4（次序无关，标签决定身份）
        let k1 = null, k2 = null, k3 = null, k4 = null;
        for (const c of charts) {
            try {
                const sp = spaceOf(c);
                const j = JSON.stringify(sp);
                if (!k1 && sp?.plotArea?.chartGroups && JSON.parse(j).plotArea.chartGroups.length >= 2) { k1 = c; continue; }
                if (!k2 && (j.indexOf('dispRSqr') >= 0 || j.indexOf('dispEq') >= 0)) { k2 = c; continue; }
                if (!k3) {
                    const s0 = c.series().get(0);
                    if (s0 && s0.trendlines && s0.trendlines[0] && !(s0.trendlines[0].displayEquation || s0.trendlines[0].forward)) { k3 = c; continue; }
                }
                if (!k4) {
                    const s0 = c.series().get(0);
                    if (s0 && s0.errorBars && s0.errorBars.vertical) { k4 = c; continue; }
                }
            } catch (e) { out.steps.push('probe err: ' + String(e).slice(0, 60)); }
        }
        out.steps.push('located: k1=' + (k1 && k1.name()) + ' k2=' + (k2 && k2.name()) + ' k3=' + (k3 && k3.name()) + ' k4=' + (k4 && k4.name()));

        // k1 J1/J2 次轴切换：与 Univer 修改方向一致（rightAxis:false → 线系列切回主轴）
        if (k1) {
            const sc = k1.series();
            for (let i = 0; i < 6; i++) {
              try {
                const s = sc.get(i);
                if (!s) break;
                if (s && enumName(C.ChartType, s.chartType).toLowerCase().indexOf('line') >= 0) {
                    s.chartType = C.ChartType.line;
                    s.axisGroup = C.AxisGroup.primary;
                    sc.set(i, s);
                    out.steps.push('k1: line series#' + i + ' axisGroup->primary, readback=' + enumName(C.AxisGroup, k1.series().get(i).axisGroup));
                    break;
                }
              } catch (e) { break; }
            }
        }
        // k2 J3a 趋势线改红 + 公式R²
        if (k2) {
            const s = k2.series().get(0);
            s.trendlines = [{ type: C.TrendlineType.linear, displayEquation: true, displayRSquared: true, style: { color: 'red', width: 2 } }];
            k2.series().set(0, s);
            out.steps.push('k2: trendline red+eq+R2, readback color=' + JSON.stringify(k2.series().get(0).trendlines?.[0]?.style));
        }
        // k3 J3b 移动平均 period=3
        if (k3) {
            const mvKey = Object.keys(C.TrendlineType).find(k => /oving/i.test(k));
            const s = k3.series().get(0);
            s.trendlines = [{ type: C.TrendlineType[mvKey], period: 3, style: { color: 'blue' } }];
            k3.series().set(0, s);
            out.steps.push('k3: movingAvg period=3, readback period=' + k3.series().get(0).trendlines?.[0]?.period);
        }
        // k4 无代码入口，不改
        return out;
    });
    console.log(JSON.stringify(log, null, 1));
    await page.waitForTimeout(2500);

    // 逐图导出（k1..k4 → _sjs_fixed_k1..4.png）
    const b64s = await page.evaluate(() => {
        const sheet = window.__spread.getSheetFromName('09高级元素');
        const cs = sheet.charts.all();
        // 重新按谓词排序导出，保证 k1~k4 顺序一致
        const spaceOf = c => { const j = c.toJSON(); return j.shapeData?.graphicFrame?.graphic?.graphicData?.chart?.chartSpace?.chart; };
        const order = [];
        const used = new Set();
        const tryAdd = (pred) => { for (const c of cs) { if (used.has(c.name())) continue; try { if (pred(c)) { used.add(c.name()); order.push(c); return; } } catch (e) { } } };
        tryAdd(c => { try { const sp = spaceOf(c); return sp?.plotArea?.chartGroups && JSON.parse(JSON.stringify(sp)).plotArea.chartGroups.length >= 2; } catch (e) { return false; } });
        tryAdd(c => { try { const j = JSON.stringify(spaceOf(c)); return j.indexOf('dispRSqr') >= 0 || j.indexOf('dispEq') >= 0; } catch (e) { return false; } });
        tryAdd(c => { try { const s0 = c.series().get(0); return s0 && s0.trendlines && s0.trendlines[0] && !(s0.trendlines[0].displayEquation || s0.trendlines[0].forward); } catch (e) { return false; } });
        tryAdd(c => { try { const s0 = c.series().get(0); return s0 && s0.errorBars && s0.errorBars.vertical; } catch (e) { return false; } });
        for (const c of cs) if (!used.has(c.name())) order.push(c);
        return Promise.all(order.slice(0, 4).map(async c => { try { return await c.toImageDataURL(); } catch (e) { return null; } }));
    });
    b64s.forEach((d, i) => {
        if (d && d.startsWith('data:image')) fs.writeFileSync(path.join(ROOT, `_sjs_fixed_k${i + 1}.png`), Buffer.from(d.split(',')[1], 'base64'));
        console.log(`k${i + 1}: ${d ? 'ok' : 'null'}`);
    });
    await browser.close();
    server.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
