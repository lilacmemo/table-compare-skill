// Univer 导出方向：导入批次1样本 → exportSheetBySnapshotAsync → 保存 xlsx
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const FILE = process.argv[2] || path.join(__dirname, '..', '批次1样式样本.xlsx');
const OUT = process.argv[3] || path.join(__dirname, 'univer_roundtrip_batch1.xlsx');
const sleep = ms => new Promise(r => setTimeout(r, ms));
(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1500,950'] });
    const page = await (await browser.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await sleep(8000);
    const b64 = fs.readFileSync(FILE).toString('base64');
    const r = await page.evaluate(async (b64) => {
        const sleep = ms => new Promise(r2 => setTimeout(r2, ms));
        const bin = atob(b64); const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 't.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await sleep(9000);
        const wb = univerAPI.getActiveWorkbook();
        // 统计导入后图表
        let chartCount = 0;
        for (const s of wb.getSheets()) { try { chartCount += s.getCharts().length; } catch (e) { } }
        const outBlob = await univerAPI.exportSheetBySnapshotAsync(wb.save());
        const b = await new Promise((res) => { const rd = new FileReader(); rd.onload = () => res(rd.result.split(',')[1]); rd.readAsDataURL(outBlob); });
        return { chartCount, b64: b };
    }, b64);
    fs.writeFileSync(OUT, Buffer.from(r.b64, 'base64'));
    console.log('Univer 导入图表数:', r.chartCount, '| 导出文件:', OUT, fs.statSync(OUT).size, 'bytes');
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
