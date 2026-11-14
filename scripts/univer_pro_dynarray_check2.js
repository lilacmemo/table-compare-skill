// 复测 v2：修正期望值（按现场数据重算）+ Pro 构建强特征校验
'use strict';
const { chromium } = require('playwright');
const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';

(async () => {
    const browser = await chromium.launch({ headless: true, channel: 'msedge', args: ['--disable-gpu'] });
    const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
    await page.goto(URL, { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForTimeout(3000);

    // Pro 强特征：bundle 内是否含 @univerjs-pro 标识 + Pro 独有 Facade/插件
    const proCheck = await page.evaluate(async () => {
        const src = document.querySelector('script[src*="assets/index"]')?.src;
        let bundleHasPro = false, bundleLen = 0;
        if (src) {
            const txt = await (await fetch(src)).text();
            bundleLen = txt.length;
            bundleHasPro = /@univerjs-pro|univerjs-pro|sheets-chart-ui|UniverPro/i.test(txt);
        }
        // Pro 图表 Facade 探测
        let chartFacade = false;
        try { const wb = window.univerAPI.getActiveWorkbook(); chartFacade = typeof wb.insertChart === 'function' || typeof window.univerAPI.newChart === 'function'; } catch (e) {}
        return { bundleHasPro, bundleLen, chartFacade, assetSrc: src };
    });
    console.log('== Pro 构建特征 ==');
    console.log('bundle 含 @univerjs-pro 标识:', proCheck.bundleHasPro, `(bundle ${(proCheck.bundleLen / 1048576).toFixed(1)}MB)`);
    console.log('图表 Facade 存在(Pro 专属功能):', proCheck.chartFacade);

    // 动态数组复测（期望值按现场数据 A1:A6=1..6, B1:B6=6..1 重新计算）
    // FILTER(A,B>4): B>4 是行1(B=6)、行2(B=5) → [1,2] → INDEX=1
    // UNIQUE(1..6): 原序 [1..6] → INDEX(...,3,1)=3
    // SORT(B): B=6..1 排序后 1..6 → INDEX=1
    // TAKE(A,2): [1,2] → INDEX(...,2,1)=2
    // DROP(A,2): [3,4,5,6] → INDEX=3
    // SEQUENCE(5) 求和=15
    // VSTACK(A1:A3,B1:B3): [1,2,3,6,5,4] → INDEX(...,4,1)=6
    // XLOOKUP(2,H,I): I 第2行=1.1
    const cases = [
        { name: 'FILTER',   formula: '=INDEX(FILTER(A1:A6,B1:B6>4),1,1)', excel: 1 },
        { name: 'UNIQUE',   formula: '=INDEX(UNIQUE(A1:A6),3,1)',          excel: 3 },
        { name: 'SORT',     formula: '=INDEX(SORT(B1:B6),1,1)',            excel: 1 },
        { name: 'TAKE',     formula: '=INDEX(TAKE(A1:A6,2),2,1)',          excel: 2 },
        { name: 'DROP',     formula: '=INDEX(DROP(A1:A6,2),1,1)',          excel: 3 },
        { name: 'SEQUENCE', formula: '=SUM(SEQUENCE(5))',                  excel: 15 },
        { name: 'VSTACK',   formula: '=INDEX(VSTACK(A1:A3,B1:B3),4,1)',    excel: 6 },
        { name: 'XLOOKUP',  formula: '=XLOOKUP(2,H1:H3,I1:I3)',            excel: 1.1 },
    ];
    const values = await page.evaluate(async (cases) => {
        const wb = window.univerAPI.getActiveWorkbook();
        const sheet = wb.getActiveSheet();
        for (let i = 0; i < 6; i++) sheet.getRange(i, 0).setValue(i + 1);
        for (let i = 0; i < 6; i++) sheet.getRange(i, 1).setValue(6 - i);
        sheet.getRange(0, 7).setValue(2); sheet.getRange(1, 7).setValue(1); sheet.getRange(2, 7).setValue(3);
        sheet.getRange(0, 8).setValue(2.2); sheet.getRange(1, 8).setValue(1.1); sheet.getRange(2, 8).setValue(3.3);
        for (let i = 0; i < cases.length; i++) sheet.getRange(i, 3).setFormula(cases[i].formula);
        const f = window.univerAPI.getFormula();
        try { f.executeCalculation(); await f.onCalculationResultApplied(15000); } catch (e) {}
        await new Promise(r => setTimeout(r, 1500));
        return cases.map((c, i) => ({ name: c.name, excel: c.excel, got: String(sheet.getRange(i, 3).getDisplayValue()) }));
    }, cases);
    console.log('\n== 动态数组复测 v2（期望值按现场数据重算） ==');
    let pass = 0;
    for (const v of values) {
        const ok = v.got === String(v.excel);
        if (ok) pass++;
        console.log(`${v.name.padEnd(9)} 期望=${v.excel} 实得=${v.got} ${ok ? 'OK' : '<<< 不一致'}`);
    }
    console.log(`\n结果: ${pass}/${cases.length} 一致`);
    await browser.close();
})().catch(e => { console.error('[fatal]', String(e).slice(0, 400)); process.exit(1); });
