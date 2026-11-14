// Univer 图表类型实测：按官方文档范式 sheet.newChart(ChartTypeString).setSource().setPosition().build() → insertChart
// 用法：node univer_create_types.js
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1500,950'] });
    const page = await (await browser.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
    page.on('pageerror', e => console.error('PAGE_ERROR:', e.message));
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await sleep(8000);

    const result = await page.evaluate(() => {
        const api = window.univerAPI;
        const out = { chartTypeEnum: null, created: [], failed: [], results: [], finalCount: 0 };
        try { out.chartTypeEnum = api.Enum.ChartTypeString; } catch (e) { out.chartTypeEnum = String(e); }
        const entries = out.chartTypeEnum && typeof out.chartTypeEnum === 'object' ? Object.entries(out.chartTypeEnum).filter(([, v]) => typeof v === 'string') : [];
        out.chartTypeEnum = Object.fromEntries(entries);

        const wb = api.getActiveWorkbook();
        const sheet = wb.getSheets()[0];
        try {
            sheet.getRange('H1:J6').setValues([['类', 'A', 'B'], ['一', 30, 25], ['二', 44, 38], ['三', 26, 41], ['四', 52, 33], ['五', 18, 29]]);
        } catch (e) { out.failed.push('setValues: ' + String(e).slice(0, 120)); }

        let i = 0;
        const tryOne = async (label, enumVal) => {
            try {
                const info = sheet.newChart(enumVal).setSource('H1:J6').setPosition(`M${2 + (i % 25) * 2}`).setSize(400, 260).build();
                const fChart = await sheet.insertChart(info);
                i++;
                let readType = null;
                try { readType = fChart.getType(); } catch (e) {}
                out.created.push(label);
                out.results.push({ type: label, ok: true, readType, downgraded: readType !== enumVal ? `${enumVal} -> ${readType}` : false });
            } catch (e) {
                out.failed.push(label);
                out.results.push({ type: label, ok: false, err: String(e && e.message || e).slice(0, 200) });
            }
        };

        const run = async () => {
            for (const [name, val] of Object.entries(out.chartTypeEnum || {})) {
                await tryOne('newChart:' + name, val);
            }
            try { out.finalCount = sheet.getCharts().length; } catch (e) { out.finalCount = -1; }
        };
        return run().then(() => out);
    });

    fs.writeFileSync(path.join(__dirname, 'univer_create_types.json'), JSON.stringify(result, null, 1));
    console.log('ChartTypeString 枚举:', JSON.stringify(result.chartTypeEnum));
    console.log('可创建:', result.created.length, '| 失败:', result.failed.length, '| 最终工作表图表数:', result.finalCount);
    if (result.failed.length) console.log('失败明细:', JSON.stringify(result.failed));
    const dg = result.results.filter(r => r.ok && r.downgraded);
    if (dg.length) console.log('类型变化:', dg.map(r => r.downgraded).join(', '));
    // 渲染截图：视口内前若干个图表
    await page.screenshot({ path: path.join(__dirname, 'univer_createtype_render_check.png') });
    // 滚动画布截第二张
    await page.keyboard.press('PageDown').catch(() => {});
    await sleep(1500);
    await page.screenshot({ path: path.join(__dirname, 'univer_createtype_render_check2.png') });
    console.log('截图: univer_createtype_render_check.png / _check2.png');
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
