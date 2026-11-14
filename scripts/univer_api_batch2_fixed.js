// univer_api_batch2_fixed.js —— Univer 导入批次2 → setSize 修复 0×0 → 重放 I/J 组代码修改 → 截图
// 修改调用与 univer_api_verify2.js round2 完全一致（I1/I2、J3a、J3b、J1/J2）
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const FILE = path.join(__dirname, '..', '批次2系列与高级元素样本.xlsx');

(async () => {
    const browser = await chromium.launch({ channel: 'msedge', headless: true });
    const ctx = await browser.newContext({ viewport: { width: 1500, height: 2200 }, deviceScaleFactor: 2 });
    const page = await ctx.newPage();
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/charts', { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await page.waitForTimeout(8000);

    const baseline = JSON.parse(fs.readFileSync(path.join(__dirname, 'baseline_chart_pos.json'), 'utf-8').toString().replace(/^\uFEFF/, ''));
    const b64 = fs.readFileSync(FILE).toString('base64');

    const r = await page.evaluate(async ({ b64, baseline }) => {
        const bin = atob(b64);
        const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const f = new File([bytes], 'batch2.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const snap = await univerAPI.importSheetToSnapshotAsync(f);
        univerAPI.createWorkbook(snap);
        await new Promise(r => setTimeout(r, 12000));
        const wb = univerAPI.getActiveWorkbook();
        const log = [];
        const chartsOf = (n) => { const s = wb.getSheets().find(x => x.getSheetName() === n); return s ? s.getCharts() : []; };

        // 1) 全部图 setSize 修复（Excel pt × 4/3）
        for (const sheetName of ['08系列样式', '09高级元素']) {
            const cos = baseline['b2'][sheetName];
            const charts = chartsOf(sheetName);
            charts.forEach((c, k) => {
                const e = cos[k];
                if (e) { try { c.setSize(Math.round(e.wPt * 4 / 3), Math.round(e.hPt * 4 / 3)); } catch (err) { log.push('setSize err ' + sheetName + ' k' + k); } }
            });
            log.push(sheetName + ' setSize done: ' + charts.length);
        }
        await new Promise(r => setTimeout(r, 4000));

        const applyUpd = async (chart, variants, tag) => {
            const tried = [];
            for (const v of variants) {
                try { await chart.update(v); tried.push(JSON.stringify(v).slice(0, 80)); } catch (e) { tried.push('ERR:' + String(e).slice(0, 60)); }
            }
            await new Promise(r => setTimeout(r, 1000));
            return tried;
        };

        // 2) I1/I2：系列填充改红 + 绿边框
        const c8 = chartsOf('08系列样式')[0];
        const s8 = c8.getInfo().config.series || [];
        const selA = s8[0].selector;
        await applyUpd(c8, [{ series: [{ selector: selA, color: '#FF0000', border: { color: '#00FF00', visible: true, width: 3 } }] }], 'I1/I2');

        const c9 = chartsOf('09高级元素');
        // J3a 趋势线改红+公式R²（09 表第 2 图 = c2 趋势线图）
        const t1 = c9[1].getInfo().config.trendlines;
        if (t1 && t1[0]) await applyUpd(c9[1], [{ trendlines: [Object.assign({}, t1[0], { color: '#FF0000', showFormula: true, showR2: true })] }], 'J3a');
        // J3b 移动平均（第 3 图）
        const t2 = c9[2] && c9[2].getInfo().config.trendlines;
        if (t2 && t2[0]) await applyUpd(c9[2], [{ trendlines: [Object.assign({}, t2[0], { period: 3 })] }], 'J3b');
        // J1/J2 次轴切换（第 1 图组合图）
        const combo = c9[0].getInfo().config.series || [];
        const lineSer = combo.find(s => s.type === 'line');
        if (lineSer) await applyUpd(c9[0], [{ series: [{ selector: lineSer.selector, rightAxis: false }] }], 'J1/J2');

        await new Promise(r => setTimeout(r, 3000));
        return { log };
    }, { b64, baseline });
    console.log(JSON.stringify(r.log));

    // 3) 逐图 exportImage（组合图/趋势线/移动平均/误差线）
    const imgs = await page.evaluate(async () => {
        const wb = univerAPI.getActiveWorkbook();
        const s = wb.getSheets().find(x => x.getSheetName() === '09高级元素');
        wb.setActiveSheet(s);
        await new Promise(r => setTimeout(r, 1500));
        const charts = s.getCharts();
        const out = [];
        for (const k of [0, 1, 2, 3]) {
            try { const d = await charts[k].exportImage(); out.push(d && d.startsWith('data:image') ? d : null); }
            catch (e) { out.push(null); }
        }
        return out;
    });
    imgs.forEach((d, i) => { if (d) fs.writeFileSync(path.join(__dirname, `_fixed_k${i + 1}.png`), Buffer.from(d.split(',')[1], 'base64')); });
    console.log('imgs:', imgs.map(d => !!d).join(','));
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
