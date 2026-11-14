/**
 * verify_config_preserve.js —— 配置项「支不支持」的三向验证
 *   Excel 原文件 XML  →  Univer 导入后 getInfo()  →  Univer 再导出 XML
 * 用途：文档"未提及"的配置项，用导入/导出行为判断内部模型是否支持
 * 用法: node verify_config_preserve.js <xlsx>
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = process.argv[2];
const OUT = path.join(__dirname, 'exported_from_univer.xlsx');
if (!FILE) { console.error('用法: node verify_config_preserve.js <xlsx>'); process.exit(1); }

(async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 2, acceptDownloads: true });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(10000);

    const b64 = fs.readFileSync(FILE).toString('base64');
    const cfg = await page.evaluate(async (b64) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 't.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await new Promise(r => setTimeout(r, 9000));
        const wb = univerAPI.getActiveWorkbook();
        const s = wb.getSheets()[0];
        wb.setActiveSheet(s);
        const chart = s.getCharts()[0];
        const info = chart.getInfo();
        return { sheet: s.getSheetName(), id: chart.getId(), type: String(chart.getType()), config: info.config, size: info.size, position: info.position, anchor: info.anchor };
    }, b64);

    console.log('=== Univer 导入后 getInfo() ===');
    console.log(JSON.stringify(cfg, null, 1).slice(0, 4000));
    fs.writeFileSync(path.join(__dirname, 'univer_config_dump.json'), JSON.stringify(cfg, null, 1));

    // 导出回 xlsx，供 XML 对比
    try {
        const [dl] = await Promise.all([
            page.waitForEvent('download', { timeout: 180000 }),
            page.evaluate(async () => {
                const wb = univerAPI.getActiveWorkbook();
                const file = await univerAPI.exportSheetBySnapshotAsync(wb.save(), 'xlsx');
                if (!file) throw new Error('export null');
                univerAPI.downloadFile(file, 'exported_from_univer', 'xlsx');
            }),
        ]);
        await dl.saveAs(OUT);
        console.log('\n导出成功:', OUT);
    } catch (e) {
        console.error('\n导出失败:', e.message);
    }

    await browser.close();
})().catch(e => { console.error('FAIL:', e.message); process.exit(1); });
