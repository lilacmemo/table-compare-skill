// Univer 挂起机制语义验证 v2：初始化语义 + 运行时写公式是否受控
'use strict';
const { chromium } = require('playwright');
const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';

function buildData(formulaRows) {
    const cellData = {};
    for (let r = 0; r < formulaRows; r++) {
        cellData[r] = {
            0: { v: r + 1 },
            1: { f: `=A${r + 1}*2` },   // B = A*2
        };
    }
    return {
        name: 'probe',
        sheetOrder: ['s1'],
        sheets: {
            s1: { id: 's1', name: 'S', rowCount: formulaRows + 10, columnCount: 5, cellData },
        },
    };
}

(async () => {
    const browser = await chromium.launch({ headless: true, channel: 'msedge', args: ['--disable-gpu'] });
    const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
    await page.goto(URL, { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForTimeout(3000);

    const out = await page.evaluate(async (data) => {
        const res = {};
        const f = window.univerAPI.getFormula();
        // 页面内重建数据（Case 4 用）
        const buildData = (formulaRows) => {
            const cellData = {};
            for (let r = 0; r < formulaRows; r++) cellData[r] = { 0: { v: r + 1 }, 1: { f: `=A${r + 1}*2` } };
            return { name: 'probe-big', sheetOrder: ['s1'], sheets: { s1: { id: 's1', name: 'S', rowCount: formulaRows + 10, columnCount: 5, cellData } } };
        };
        const readB = (wb, r) => {
            try { const s = wb.getSheetBySheetId('s1'); const v = s.getRange(r, 1).getValue(); return v === undefined || v === null ? String(v) : `${v}(${typeof v})`; }
            catch (e) { return 'ERR:' + String(e).slice(0, 60); }
        };
        const wait = (ms) => new Promise(r => setTimeout(r, ms));

        // Case 1: NO_CALCULATION 下建簿（初始化不计算）
        f.setInitialFormulaComputing(2);
        const wb1 = window.univerAPI.createWorkbook(data);
        await wait(800);
        res.noCalc_before = readB(wb1, 0);
        const t = performance.now();
        f.executeCalculation();
        try { await f.onCalculationResultApplied(30000); } catch (e) {}
        await wait(800);
        res.noCalc_after = readB(wb1, 0);
        res.noCalc_calcMs = performance.now() - t;

        // Case 2: 默认 WHEN_EMPTY 建簿（初始化自动计算）
        f.setInitialFormulaComputing(1);
        const t2 = performance.now();
        const wb2 = window.univerAPI.createWorkbook(data);
        try { await f.onCalculationResultApplied(30000); } catch (e) {}
        await wait(800);
        res.whenEmpty_initial = readB(wb2, 0);
        res.whenEmpty_calcMs = performance.now() - t2;

        // Case 3: NO_CALCULATION 建簿后再写新公式 —— 是否运行时也不计算？
        f.setInitialFormulaComputing(2);
        const wb3 = window.univerAPI.createWorkbook(data);
        await wait(500);
        const s3 = wb3.getSheetBySheetId('s1');
        s3.getRange(5, 3).setFormula('=A6*100'); // D6
        try { await f.onCalculationResultApplied(15000); } catch (e) {}
        await wait(800);
        res.runtimeWrite_after = (() => { const v = s3.getRange(5, 3).getValue(); return v === undefined || v === null ? String(v) : `${v}(${typeof v})`; })();

        // Case 4: stopCalculation 在计算进行中调用是否中断（大簿触发后立刻 stop）
        f.setInitialFormulaComputing(1);
        const big = buildData(5000);
        const wb4 = window.univerAPI.createWorkbook(big);
        let stopped = false;
        try { f.executeCalculation(); f.stopCalculation(); stopped = true; } catch (e) { res.stopErr = String(e).slice(0, 80); }
        await wait(1500);
        res.stopCalled = stopped;
        res.big_afterStop = readB(wb4, 0);
        // 再完整算一次
        try { f.executeCalculation(); await f.onCalculationResultApplied(60000); } catch (e) {}
        await wait(800);
        res.big_afterFull = readB(wb4, 0);
        return res;
    }, buildData(20));

    console.log('== Case 1: NO_CALCULATION 建簿（初始化语义）==');
    console.log('  建簿后立即读 B1:', out.noCalc_before);
    console.log('  executeCalculation 后读 B1:', out.noCalc_after, `(计算耗时 ${out.noCalc_calcMs.toFixed(0)}ms)`);
    console.log('\n== Case 2: 默认 WHEN_EMPTY 建簿 ==');
    console.log('  建簿后读 B1:', out.whenEmpty_initial, `(建簿+计算 ${out.whenEmpty_calcMs.toFixed(0)}ms)`);
    console.log('\n== Case 3: NO_CALCULATION 建簿后运行时写新公式 ==');
    console.log('  写入 D6=A6*100 后读值:', out.runtimeWrite_after);
    console.log('\n== Case 4: stopCalculation 中断完整性 ==');
    console.log('  stop 可调用:', out.stopCalled, out.stopErr ? `err=${out.stopErr}` : '');
    console.log('  stop 后 B1:', out.big_afterStop, '| 完整重算后 B1:', out.big_afterFull);

    await browser.close();
})().catch(e => { console.error('[fatal]', String(e).slice(0, 400)); process.exit(1); });
