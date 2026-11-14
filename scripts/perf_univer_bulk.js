// Univer 同口径复测：NO_CALCULATION 一次性建簿（数据+公式）→ 只计 executeCalculation
'use strict';
const http = require('http');
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const { generate } = require('./perf_workload');

const PORT = 8439;
const ROOT = __dirname;

function serve() {
    return new Promise((resolve) => {
        const srv = http.createServer((req, res) => {
            const p = decodeURIComponent(req.url.split('?')[0]);
            if (p === '/') { res.writeHead(200, { 'Content-Type': 'text/html' }); res.end(PAGE); return; }
            const fp = path.join(ROOT, p);
            if (!fs.existsSync(fp) || !fs.statSync(fp).isFile()) { res.writeHead(404); res.end('nf'); return; }
            res.writeHead(200, { 'Content-Type': p.endsWith('.js') ? 'application/javascript' : p.endsWith('.css') ? 'text/css' : 'application/octet-stream' });
            fs.createReadStream(fp).pipe(res);
        });
        srv.listen(PORT, '127.0.0.1', () => resolve(srv));
    });
}

const PAGE = `<!DOCTYPE html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="/dist_univer/univer_bundle.css">
<script src="/dist_univer/univer_bundle.js"></script>
<style>html,body{margin:0;width:100%;height:100%}</style></head>
<body><div id="app" style="width:1400px;height:900px"></div>
<script>
window.__ready = false;
(async () => { await window.initPerfUniver('app'); window.__ready = true; })();

function mkSheet(name, rows, cols, cellData) {
  return { id: 'sid_' + name, name, rowCount: rows, columnCount: cols, cellData };
}

window.buildAndCalc = async function (data) {
  const toCells = (vals, r0, c0, formulas) => {
    const cd = {};
    const put = (r, c, obj) => { if (!cd[r]) cd[r] = {}; cd[r][c] = obj; };
    for (let r = 0; r < vals.length; r++)
      for (let c = 0; c < vals[r].length; c++) {
        const v = vals[r][c];
        if (v !== null && v !== undefined) put(r0 + r, c0 + c, { v });
      }
    if (formulas) for (const f of formulas) {
      const m = /^([A-Z]+)(\\d+)$/.exec(f.cell);
      let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      put(parseInt(m[2]) - 1, col - 1, { f: f.formula });
    }
    return cd;
  };
  // Data：值 + 公式 合并到同一 cellData
  const dataCells = toCells(data.dataValues, 0, 0, null);
  for (let r = 1; r < data.dataFormulas.length; r++) {
    const row = data.dataFormulas[r];
    for (let c = 0; c < 4; c++) {
      const f = row[5 + c];
      if (f) { if (!dataCells[r]) dataCells[r] = {}; dataCells[r][5 + c] = { f }; }
    }
  }
  const sheets = {
    sid_Data: mkSheet('Data', data.dataValues.length + 1, 9, dataCells),
    sid_Params: mkSheet('Params', 4, 2, toCells(data.paramsValues, 0, 0, null)),
    sid_Products: mkSheet('Products', 21, 2, toCells(data.productsValues, 0, 0, null)),
    sid_Report: (() => { const cd = toCells(data.reportValues, 0, 0, null); const cd2 = toCells(data.reportPart2, 8, 0, null); for (const k in cd2) cd[k] = Object.assign(cd[k] || {}, cd2[k]); return mkSheet('Report', 20, 21, cd); })(),
    sid_Chain: mkSheet('Chain', data.chainValues.length, 4, toCells(data.chainValues, 0, 0, null)),
    sid_Dynamic: mkSheet('Dynamic', data.dataValues.length + 1, 30, {}),
  };
  // Report / Chain / Dynamic 公式合并
  const mergeFormulas = (sheet, list) => {
    for (const f of list) {
      const m = /^([A-Z]+)(\\d+)$/.exec(f.cell);
      let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      const r = parseInt(m[2]) - 1;
      if (!sheet.cellData[r]) sheet.cellData[r] = {};
      sheet.cellData[r][col] = { f: f.formula };
    }
  };
  mergeFormulas(sheets.sid_Report, data.reportFormulas);
  mergeFormulas(sheets.sid_Chain, data.chainFormulas);
  mergeFormulas(sheets.sid_Dynamic, data.dynFormulas);

  const f = window.univerAPI.getFormula();
  f.setInitialFormulaComputing(2); // NO_CALCULATION：建簿不计算
  const t0 = performance.now();
  const wb = window.univerAPI.createWorkbook({ name: 'bulk', sheetOrder: ['sid_Data', 'sid_Params', 'sid_Products', 'sid_Report', 'sid_Chain', 'sid_Dynamic'], sheets });
  const buildMs = performance.now() - t0;
  window.__wb = wb;
  window.__sheets = { Data: wb.getSheetBySheetId('sid_Data'), Report: wb.getSheetBySheetId('sid_Report'), Chain: wb.getSheetBySheetId('sid_Chain'), Dynamic: wb.getSheetBySheetId('sid_Dynamic') };
  await new Promise(r => setTimeout(r, 500));
  const before = (() => { try { return String(window.__sheets.Data.getRange(1, 5).getValue()); } catch (e) { return 'ERR'; } })();

  const times = [];
  for (let i = 0; i < 3; i++) {
    const t = performance.now();
    f.executeCalculation();
    await new Promise((resolve) => {
      let done = false;
      const d = f.onCalculationResultApplied(120000);
      Promise.resolve(d).then(() => { if (!done) { done = true; resolve(); } });
      setTimeout(() => { if (!done) { done = true; resolve(); } }, 120000);
    });
    await new Promise(r => setTimeout(r, 800));
    times.push(performance.now() - t);
  }
  const samples = [];
  for (const s of data.samples) {
    const ws = window.__sheets[s.sheet] || window.__wb.getSheetBySheetId('sid_' + s.sheet);
    const m = /^([A-Z]+)(\\d+)$/.exec(s.cell);
    let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    samples.push({ sheet: s.sheet, cell: s.cell, value: ws.getRange(parseInt(m[2]) - 1, col - 1).getValue() });
  }
  return { buildMs, beforeCalc: before, calcTimes: times, samples };
};
</script></body></html>`;

(async () => {
    const data = generate(Number(process.argv[2] || 2000));
    console.log(`[workload] 公式=${data.formulaCount}`);
    const srv = await serve();
    const browser = await chromium.launch({ headless: true, channel: 'msedge', args: ['--disable-gpu'] });
    const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
    page.on('pageerror', (e) => console.log('[pageerror]', String(e).slice(0, 200)));
    await page.goto(`http://127.0.0.1:${PORT}/`, { waitUntil: 'load', timeout: 120000 });
    await page.waitForFunction('window.__ready === true', { timeout: 120000 });
    const r = await page.evaluate((d) => window.buildAndCalc(d), data);
    console.log(`[build] NO_CALCULATION 一次性建簿(含公式): ${r.buildMs.toFixed(0)}ms`);
    console.log(`[build] 建簿后未计算样本 F2 = ${JSON.stringify(r.beforeCalc)}（期望未计算）`);
    console.log(`[calc] executeCalculation ×3: ${r.calcTimes.map(t => t.toFixed(0)).join(', ')} ms`);
    for (const s of r.samples) console.log(`  ${s.sheet}!${s.cell} = ${s.value}`);
    fs.writeFileSync(path.join(ROOT, 'perf_result_univer_bulk.json'), JSON.stringify({ buildMs: r.buildMs, calcTimes: r.calcTimes, samples: r.samples, env: 'office.univer.ai? no — local preset 1.3.1, NO_CALCULATION init' }, null, 2));
    await browser.close();
    srv.close();
})().catch(e => { console.error('[fatal]', String(e).slice(0, 500)); process.exit(1); });
