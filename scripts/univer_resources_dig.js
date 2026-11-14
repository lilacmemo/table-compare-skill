// batch1_resources_dig.js —— 取回 Univer 导入后的完整 resources 内部模型
// 目的：判定 8 个 ⚠ 项（Facade config 无键的项）在内部模型层是否保留
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = process.argv[2] || path.join(__dirname, '..', '批次1样式样本.xlsx');

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: false, args: ['--window-size=1560,980'] });
    const ctx = await browser.newContext({ viewport: null });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(8000);

    const b64 = fs.readFileSync(FILE).toString('base64');
    const result = await page.evaluate(async (b64) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 'batch1.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        // 完整 resources（不截断）——内部模型层
        const resources = {};
        for (const [k, v] of Object.entries(snap.resources || {})) {
            resources[k] = typeof v === 'string' ? v : JSON.stringify(v);
        }
        return { resources };
    }, b64);

    fs.writeFileSync(path.join(__dirname, 'batch1_resources_full.json'), JSON.stringify(result.resources, null, 1));
    for (const [k, v] of Object.entries(result.resources)) {
        console.log(`资源键: ${k}  (${v.length} chars, 含chart=${v.toLowerCase().includes('chart')})`);
    }
    await browser.close();
    console.log('DONE');
    process.exit(0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
