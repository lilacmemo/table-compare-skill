// univer_v2.js -- 用 v2 语义用例在同一数据区上跑 Univer Pro（playground），输出 probe_univer_v2.json
// 输入与 Excel / SpreadJS 完全一致：data_block_v2.json + cases_v2.json
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const HERE = __dirname;
const CASES = JSON.parse(fs.readFileSync(path.join(HERE, 'cases_v2.json'), 'utf8'));
const DATA = JSON.parse(fs.readFileSync(path.join(HERE, 'data_block_v2.json'), 'utf8'));
const OUT = path.join(HERE, 'probe_univer_v2.json');
const URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const sleep = ms => new Promise(r => setTimeout(r, ms));

function addr(a) {
    const m = /^([A-Z]+)(\d+)$/.exec(a);
    let c = 0;
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
}

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    const page = await (await browser.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
    page.on('pageerror', e => console.error('PAGEERR', String(e).slice(0, 200)));
    await page.goto(URL, { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
    await page.waitForFunction(() => typeof window.univerAPI !== 'undefined', null, { timeout: 60000 }).catch(() => {});
    await sleep(8000);

    const r = await page.evaluate(async ([cases, data]) => {
        const sleep2 = ms => new Promise(r2 => setTimeout(r2, ms));
        const addr2 = (a) => {
            const m = /^([A-Z]+)(\d+)$/.exec(a);
            let c = 0;
            for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
            return { r: parseInt(m[2], 10) - 1, c: c - 1 };
        };
        const out = { setErrors: [], rows: [] };
        univerAPI.createWorkbook({});
        await sleep2(6000);
        const wb = univerAPI.getActiveWorkbook();
        const sheet = wb.getActiveSheet();
        try { if (sheet.setRowCount) sheet.setRowCount(cases.length * 12 + 200); } catch (e) { out.resizeErr = String(e); }
        try { if (sheet.setColumnCount) sheet.setColumnCount(40); } catch (e) { out.resizeErr2 = String(e); }
        out.apiProbe = { univerVersion: (univerAPI.getVersion && String(univerAPI.getVersion())) || null };

        // 数据区（与 Excel / SpreadJS 完全一致）
        for (const cell of (data.cells || [])) {
            const p = addr2(cell.cell);
            try { sheet.getRange(p.r, p.c).setValue(cell.value); } catch (e) { out.setErrors.push({ cell: cell.cell, err: String(e).slice(0, 80) }); }
        }
        for (const f of (data.formulaCells || [])) {
            const p = addr2(f.cell);
            try { sheet.getRange(p.r, p.c).setFormula(f.formula); } catch (e) { out.setErrors.push({ cell: f.cell, err: String(e).slice(0, 80) }); }
        }

        // 公式列 Z（索引 25），行间隔 12
        for (let i = 0; i < cases.length; i++) {
            const row = 2 + i * 12, c = cases[i];
            let ok = true;
            try { sheet.getRange(row, 25).setFormula(c.formula); }
            catch (e) { ok = false; out.setErrors.push({ id: c.id, name: c.name, err: String(e && e.message || e).slice(0, 150) }); }
            out.rows.push({ id: c.id, name: c.name, formula: c.formula, setOk: ok, value: null, valueType: null, rawObj: null });
        }
        await sleep2(6000);
        try {
            const calc = wb.getCalculation ? wb.getCalculation() : null;
            if (calc && calc.executeCalculation) { await calc.executeCalculation(); }
        } catch (e) { }
        await sleep2(6000);

        const raw = wb.save();
        const sheetId = Object.keys(raw.sheets)[0];
        const cd = (raw.sheets[sheetId] && raw.sheets[sheetId].cellData) || {};
        for (let i = 0; i < cases.length; i++) {
            const row = 2 + i * 12;
            const cell = (cd[row] || {})[25];
            if (!cell) { out.rows[i].valueType = 'null'; continue; }
            const v = cell.v;
            if (v && typeof v === 'object') {
                const j = JSON.stringify(v);
                out.rows[i].valueType = 'error';
                out.rows[i].rawObj = j.slice(0, 200);
                out.rows[i].value = { code: (v.errorType !== undefined ? v.errorType : null), raw: j.slice(0, 120) };
            } else if (v === undefined || v === null) {
                out.rows[i].valueType = 'null';
            } else {
                out.rows[i].value = v;
                out.rows[i].valueType = typeof v;
            }
        }
        out.sheetDims = { rowCount: raw.sheets[sheetId] && raw.sheets[sheetId].rowCount, colCount: raw.sheets[sheetId] && raw.sheets[sheetId].columnCount };
        return out;
    }, [CASES, DATA]);

    fs.writeFileSync(OUT, JSON.stringify(r, null, 1));
    console.log('univer version:', r.apiProbe && r.apiProbe.univerVersion, '| dims:', JSON.stringify(r.sheetDims));
    console.log('setErrors:', r.setErrors.length, '| rows:', r.rows.length);
    const b = {};
    for (const x of r.rows) { const k = x.valueType === 'error' ? 'ERR ' + (x.value && x.value.code) : x.valueType; b[k] = (b[k] || 0) + 1; }
    console.log('buckets:', JSON.stringify(Object.entries(b).sort((a, c) => c[1] - a[1])));
    if (r.setErrors.length) console.log('setErr sample:', JSON.stringify(r.setErrors.slice(0, 10)));
    const samples = r.rows.filter(x => x.rawObj).slice(0, 8);
    if (samples.length) console.log('errObj sample:', JSON.stringify(samples.map(x => [x.name, x.rawObj])));
    await browser.close();
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
