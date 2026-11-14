// univer_restore_shots.js —— Univer 导入基线样本，逐图 setSize 修复 0×0 后渲染并导出
// 用户方法：导入图表高度为 0 看不见 → 先调整图表宽高，再观察配置项渲染效果。
// 每图尝试：1) setSize(原始 Excel 宽高 px) 2) exportImage() 逐图导出；导出不可用则整表截图。
// 用法: node univer_restore_shots.js <xlsx路径> <prefix>
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = process.argv[2] || path.join(__dirname, '..', '批次1样式样本.xlsx');
const PREFIX = process.argv[3] || 'b1';

(async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1500, height: 900 }, deviceScaleFactor: 2 });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(8000);

    // Excel 端位置（pt）用于还原尺寸：px = pt * 4/3
    const baseline = JSON.parse(fs.readFileSync(path.join(__dirname, 'baseline_chart_pos.json'), 'utf-8').toString().replace(/^\uFEFF/, ''));
    const b64 = fs.readFileSync(FILE).toString('base64');

    const result = await page.evaluate(async ({ b64, baseline, prefix }) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 'restore.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await new Promise(r => setTimeout(r, 12000));
        const wb = univerAPI.getActiveWorkbook();
        const dump = {};
        for (const batch of Object.keys(baseline)) {
            for (const sheetName of Object.keys(baseline[batch])) {
                // 只处理当前批次对应的表（b1/b2 文件分别调用）
                if (batch !== prefix) continue;
                const sheet = wb.getSheets().find(s => s.getSheetName() === sheetName);
                if (!sheet) { dump[sheetName] = { err: 'no sheet' }; continue; }
                wb.setActiveSheet(sheet);
                await new Promise(r => setTimeout(r, 800));
                const charts = sheet.getCharts();
                const entries = [];
                const cosList = baseline[batch][sheetName];
                for (let k = 0; k < charts.length; k++) {
                    const c = charts[k];
                    const e = cosList[k] || null;
                    let sizeBefore = null, sizeAfter = null, img = null, imgErr = null;
                    try { sizeBefore = JSON.parse(JSON.stringify(c.getInfo().size || null)); } catch (err) {}
                    // setSize 修复 0×0（原始 Excel 尺寸 pt→px）
                    if (e && typeof c.setSize === 'function') {
                        try { c.setSize(Math.round(e.wPt * 4 / 3), Math.round(e.hPt * 4 / 3)); } catch (err) {}
                    }
                    try { sizeAfter = JSON.parse(JSON.stringify(c.getInfo().size || null)); } catch (err) {}
                    // 逐图导出渲染图
                    if (typeof c.exportImage === 'function') {
                        try { img = await c.exportImage(); } catch (err) { imgErr = String(err); }
                    }
                    entries.push({ k: k + 1, sizeBefore, sizeAfter, hasImg: !!img && img.startsWith('data:image'), imgErr, img });
                }
                dump[sheetName] = { count: charts.length, entries };
            }
        }
        return dump;
    }, { b64, baseline, prefix: PREFIX });

    // 落盘：图片 + 位置 JSON
    const posDump = {};
    for (const [sheetName, d] of Object.entries(result)) {
        if (d.err) { console.log(sheetName, 'ERR', d.err); continue; }
        posDump[sheetName] = [];
        for (const e of d.entries) {
            const file = `univer_restore_${PREFIX}_${sheetName.replace(/[^\w]/g, '')}_k${e.k}.png`;
            if (e.hasImg) {
                fs.writeFileSync(path.join(__dirname, file), Buffer.from(e.img.split(',')[1], 'base64'));
            }
            posDump[sheetName].push({ k: e.k, file: e.hasImg ? file : null, sizeBefore: e.sizeBefore, sizeAfter: e.sizeAfter, imgErr: e.imgErr });
            console.log(`${sheetName} k${e.k}: sizeBefore=${JSON.stringify(e.sizeBefore)} sizeAfter=${JSON.stringify(e.sizeAfter)} img=${e.hasImg ? 'ok' : 'FAIL ' + (e.imgErr || 'no api')}`);
        }
    }
    fs.writeFileSync(path.join(__dirname, `univer_restore_pos_${PREFIX}.json`), JSON.stringify(posDump, null, 1));
    await browser.close();
    console.log('DONE univer_restore prefix=' + PREFIX);
    process.exit(0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
