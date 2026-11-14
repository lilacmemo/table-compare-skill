// env_probe.js -- 环境依赖 19 函数双端支持性实测（SJS 本地 + Univer playground）
// 判定：结果文本含 #NAME? → 不支持；其余（错误值/任意值）→ 引擎认识该函数 = 支持。值不对比。
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const http = require('http');

const HERE = __dirname;
const DATA = JSON.parse(fs.readFileSync(path.join(HERE, 'data_block_v2.json'), 'utf8'));
const OUT = path.join(HERE, 'probe_env.json');
const PORT = 8972;
const UNIVER_URL = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const sleep = ms => new Promise(r => setTimeout(r, ms));

const ENV = [
    ['CUBEKPIMEMBER',      '=CUBEKPIMEMBER("[kpi]","[value]",1)'],
    ['CUBEMEMBER',         '=CUBEMEMBER("[cube]","[member]")'],
    ['CUBEMEMBERPROPERTY', '=CUBEMEMBERPROPERTY("[cube]","[member]","p")'],
    ['CUBERANKEDMEMBER',   '=CUBERANKEDMEMBER("[cube]",CUBESET("[cube]","x"),1)'],
    ['CUBESET',            '=CUBESET("[cube]","x")'],
    ['CUBESETCOUNT',       '=CUBESETCOUNT(CUBESET("[cube]","x"))'],
    ['CUBEVALUE',          '=CUBEVALUE("[cube]","[member]")'],
    ['CALL',               '=CALL(1,"a","b",1)'],
    ['REGISTER.ID',        '=REGISTER.ID(1,"a","b")'],
    ['EUROCONVERT',        '=EUROCONVERT(100,"EUR","DEM")'],
    ['GETPIVOTDATA',       '=GETPIVOTDATA("a",$A$1)'],
    ['RTD',                '=RTD("a","","b","c")'],
    ['IMAGE',              '=IMAGE("https://example.com/x.png")'],
    ['STOCKHISTORY',       '=STOCKHISTORY("MSFT",DATE(2024,1,1),DATE(2024,2,1))'],
    ['TRANSLATE',          '=TRANSLATE("hello","en","zh")'],
    ['WEBSERVICE',         '=WEBSERVICE("https://example.com")'],
    ['JIS',                '=JIS("abc")'],
    ['SHEET',              '=SHEET()'],
    ['SHEETS',             '=SHEETS()'],
];

// ---------- SJS 侧 ----------
async function runSJS(browser) {
    const server = http.createServer((req, res) => {
        const p = path.join(HERE, decodeURIComponent(req.url.split('?')[0]));
        fs.readFile(p, (e, d) => {
            if (e) { res.writeHead(404); res.end(); return; }
            const ext = path.extname(p);
            res.writeHead(200, { 'Content-Type': ext === '.html' ? 'text/html' : 'application/octet-stream' });
            res.end(d);
        });
    });
    await new Promise(r => server.listen(PORT, r));
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    await page.goto(`http://localhost:${PORT}/formula_probe.html`);
    await page.waitForFunction(() => typeof window.runFinal === 'function');
    const ver = await page.evaluate(() => window.sjsVersion());
    const r = await page.evaluate(async ([cases, data]) => await window.runFinal(cases, data), [ENV.map(([n, f]) => ({ id: n, name: n, family: 'env', formula: f })), DATA]);
    await browser.close(); server.close();
    return { engine: 'SpreadJS ' + ver, rows: r.rows.map(x => ({ name: x.name, value: x.value, valueType: x.valueType, display: x.display !== undefined ? x.display : null })) };
}

// ---------- Univer 侧 ----------
async function runUniver(browser) {
    const page = await (await browser.newContext({ viewport: { width: 1500, height: 900 } })).newPage();
    page.on('pageerror', e => console.error('PAGEERR', String(e).slice(0, 200)));
    await page.goto(UNIVER_URL, { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
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
        for (const cell of (data.cells || [])) {
            const p = addr2(cell.cell);
            try { sheet.getRange(p.r, p.c).setValue(cell.value); } catch (e) {}
        }
        for (let i = 0; i < cases.length; i++) {
            const row = 2 + i * 6;
            try { sheet.getRange(row, 1).setFormula(cases[i][1]); } catch (e) { out.setErrors.push({ name: cases[i][0], err: String(e).slice(0, 120) }); }
        }
        await sleep2(6000);
        try { const c = wb.getCalculation; if (c && wb.getCalculation().executeCalculation) await wb.getCalculation().executeCalculation(); } catch (e) {}
        await sleep2(6000);
        for (let i = 0; i < cases.length; i++) {
            const row = 2 + i * 6;
            const rec = { name: cases[i][0] };
            try { rec.getValue = sheet.getRange(row, 1).getValue(); } catch (e) { rec.getValueErr = String(e).slice(0, 80); }
            try { rec.display = String(sheet.getRange(row, 1).getDisplayValue()); } catch (e) { rec.displayErr = String(e).slice(0, 80); }
            out.rows.push(rec);
        }
        const raw = wb.save();
        const sid = Object.keys(raw.sheets)[0];
        const cd = (raw.sheets[sid] && raw.sheets[sid].cellData) || {};
        for (let i = 0; i < cases.length; i++) {
            const cell = (cd[2 + i * 6] || {})[1];
            out.rows[i].snapshotCell = cell ? JSON.stringify(cell).slice(0, 200) : null;
        }
        return out;
    }, [ENV, DATA]);
    await browser.close();
    return { engine: 'Univer playground', rows: r.rows, setErrors: r.setErrors };
}

(async () => {
    const browser = await chromium.launch({ channel: 'chrome', headless: true });
    console.log('== SJS ==');
    const sjs = await runSJS(browser);
    const browser2 = await chromium.launch({ channel: 'chrome', headless: true });
    console.log('== Univer ==');
    const uni = await runUniver(browser2);
    const out = { sjs, univer: uni };
    fs.writeFileSync(OUT, JSON.stringify(out, null, 1));
    console.log(JSON.stringify(out, null, 1));
    process.exit(0);
})().catch(e => { console.error('FATAL', e); process.exit(1); });
