/**
 * probe_setters.js —— 逐个验证 Univer Facade 图表 setter 的真实能力
 * 针对：图例字体/颜色、轴标签字体、网格线、数据标签、update() 任意配置
 * 用法: node probe_setters.js <xlsx>
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const FILE = process.argv[2];
if (!FILE) { console.error('用法: node probe_setters.js <xlsx>'); process.exit(1); }

(async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2 });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(10000);

    const b64 = fs.readFileSync(FILE).toString('base64');
    const r = await page.evaluate(async (b64) => {
        const bin = atob(b64); const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 't.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await new Promise(r2 => setTimeout(r2, 8000));
        const wb = univerAPI.getActiveWorkbook();
        wb.setActiveSheet(wb.getSheets()[0]);
        await new Promise(r2 => setTimeout(r2, 1500));
        const chart = wb.getSheets()[0].getCharts()[0];

        const out = {};
        // 0. 方法签名（源码即文档）
        for (const m of ['setLegend', 'setTitle', 'setAppearance', 'setTheme', 'setPalette', 'update']) {
            try { out['src_' + m] = chart[m].toString().slice(0, 400); } catch (e) { out['src_' + m] = 'ERR'; }
        }
        // 1. setLegend 尝试带 fontSize/color（文档未写的字段）
        try {
            chart.setLegend({ visible: true, position: 'top', fontSize: 22, color: '#FF00FF', bold: true });
            await new Promise(r2 => setTimeout(r2, 800));
            const lg = chart.getInfo().config.legend;
            out.after_setLegend = lg;
        } catch (e) { out.after_setLegend = 'ERR:' + String(e).slice(0, 150); }
        // 2. setTitle 带颜色字号
        try {
            chart.setTitle('标题测试', { fontSize: 24, color: '#00AA00' });
            await new Promise(r2 => setTimeout(r2, 800));
            out.after_setTitle = chart.getInfo().config.title;
        } catch (e) { out.after_setTitle = 'ERR:' + String(e).slice(0, 150); }
        // 3. update() 带轴标签样式/网格线（关键：是否接受任意 config）
        try {
            chart.update({
                axes: {
                    x: { label: { fontSize: 18, color: '#112233' } },
                    y: { gridLine: { visible: true, color: '#FF0000', width: 3 } },
                },
            });
            await new Promise(r2 => setTimeout(r2, 800));
            const ax = chart.getInfo().config.axes;
            out.after_update_x_label = ax.x.label;
            out.after_update_y_gridLine = ax.y.gridLine;
        } catch (e) { out.after_update = 'ERR:' + String(e).slice(0, 200); }
        // 4. dataLabels：series[].label 可否改
        try {
            chart.update({ series: [{ selector: { index: 1 }, label: { visible: true, fontSize: 16, color: '#AA0000', bold: true } }] });
            await new Promise(r2 => setTimeout(r2, 800));
            out.after_update_series_label = chart.getInfo().config.series;
        } catch (e) { out.after_update_series = 'ERR:' + String(e).slice(0, 200); }
        return out;
    }, b64);

    console.log(JSON.stringify(r, null, 1).slice(0, 5000));
    await browser.close();
})().catch(e => { console.error('FAIL:', e.message); process.exit(1); });
