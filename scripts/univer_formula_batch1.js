// univer_formula_batch1.js —— 公式批次1 Univer 侧实测：
// playground 导入公式批次1样本.xlsx → 等待重算 → wb.save() 快照解析出 {sheets:{name:{cells:{addr:{v,f}}}}}
// 产物：univer_formula_batch1.json（与 sjs_formula_batch1.json 同构，供 compare_formula_batch1.py 复用）
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = path.join(__dirname, '..', '公式批次1样本.xlsx');
const OUT = path.join(__dirname, 'univer_formula_batch1.json');
const sleep = ms => new Promise(r => setTimeout(r, ms));

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    const page = await (await browser.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await sleep(8000);
    const b64 = fs.readFileSync(FILE).toString('base64');

    const r = await page.evaluate(async (b64) => {
        const sleep = ms => new Promise(r2 => setTimeout(r2, ms));
        const bin = atob(b64); const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 'batch1.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await sleep(15000); // 等待公式重算

        const wb = univerAPI.getActiveWorkbook();
        const raw = wb.save();
        const colName = c => { let s = ''; c = c + 1; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
        const sheets = {};
        const sheetNames = [];
        for (const [sid, s] of Object.entries(raw.sheets || {})) {
            const name = s.name || sid;
            sheetNames.push(name);
            const cells = {};
            const cd = s.cellData || {};
            for (const [rIdx, row] of Object.entries(cd)) {
                for (const [cIdx, cell] of Object.entries(row)) {
                    if (!cell || (cell.v === undefined && !cell.f && !cell.si)) continue;
                    const addr = colName(Number(cIdx)) + (Number(rIdx) + 1);
                    let v = cell.v;
                    if (v && typeof v === 'object') v = JSON.stringify(v);
                    cells[addr] = { v: v === undefined ? null : v, f: cell.f || null, si: cell.si || undefined };
                }
            }
            sheets[name] = { cells, rowCount: s.rowCount, colCount: s.columnCount };
        }
        return { sheetNames, sheets };
    }, b64);

    fs.writeFileSync(OUT, JSON.stringify(r, null, 1));
    console.log('sheets:', r.sheetNames.join(','));
    const total = Object.values(r.sheets).reduce((a, s) => a + Object.keys(s.cells).length, 0);
    console.log('cells with content:', total, '->', OUT, fs.statSync(OUT).size, 'bytes');
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
