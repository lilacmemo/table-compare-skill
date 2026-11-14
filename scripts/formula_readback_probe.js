// Univer 公式读回口径探针：getValues() 拿到的到底是"数值"还是"显示文本"
//   用法：NODE_PATH=<poc-compare/node_modules> node tools/probe_formula_readback.js
//
// 为什么单独探：跑批里 Univer 的财务函数读回是 "(¥188.71)"、"8%"、"¥5,299.07"，
// 与 Excel 的 -188.712336440109 / 0.07713847 / 5299.07063 差在"显示精度"。
// 这类差异可能是
//   ① 引擎算错了（真差异）
//   ② 引擎算对了，但 getValues() 返回的是套了数字格式的显示值（测量口径问题）
// 两者的报告结论完全相反，必须分开。做法：同一格同时读 API 值、显示值、快照原始 v。
const { chromium } = require('playwright');
const path = require('path');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?module=formula';

const CASES = [
  ['PMT', '=PMT(0.05/12,60,10000)', -188.712336440109],
  ['IPMT', '=IPMT(0.05/12,1,60,10000)', -41.6666666666667],
  ['FV', '=FV(0.05/12,60,-100)', 6800.60828408434],
  ['PV', '=PV(0.05/12,60,-100)', 5299.07063239275],
  ['NPV', '=NPV(0.05,100,200,300,400)', 849.7],   // 仅作对照，实际值不作判据
  ['IRR', '=IRR({-1000,300,400,500})', 0.0771384729520437],
  ['TIME', '=TIME(10,20,30)', 0.430902777777778],
  ['ABS', '=ABS(-3.5)', 3.5],
];

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1680, height: 1000 } });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => console.log('PAGEERROR', String(e.message).slice(0, 160)));
  await p.goto(BASE, { waitUntil: 'domcontentloaded', timeout: 120000 });
  await p.waitForFunction(() => window.APP && APP.booted === true && APP.ready && APP.ready.uni, null, { timeout: 120000 });
  // 等样本导入落定（复算列 51 在样本默认宽度之外，先撑列）
  await p.waitForFunction(() => !!(window.APP.lastImport && APP.lastImport.uni), null, { timeout: 120000 });
  await p.waitForTimeout(2500);

  const out = await p.evaluate(async (cases) => {
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const res = {};
    const U = window.univerAPI;
    const wb = U.getActiveWorkbook();
    const sh = wb.getSheetByName('Excel实测') || wb.getActiveSheet();
    res.sheet = sh.getSheetName();

    // 复算列在 51 列（0-based 50），行从 2 开始；先撑列 + 清空 + 强制常规格式
    const R = 50;
    try { if (sh.getMaxColumns() < R + 2) sh.setColumnCount(R + 4); } catch (e) { res.eCols = String(e.message).slice(0, 120); }
    res.maxCols = sh.getMaxColumns();
    try { sh.getRange(1, R, cases.length, 1).clearContent(); } catch (e) { res.eClear = String(e.message).slice(0, 120); }
    try { sh.getRange(1, R, cases.length, 1).setNumberFormat('General'); } catch (e) { res.eFmt = String(e.message).slice(0, 120); }
    res.fmtBefore = (() => { try { return sh.getRange(1, R).getNumberFormat ? sh.getRange(1, R).getNumberFormat() : 'no-getter'; } catch (e) { return 'err:' + e.message.slice(0, 60); } })();

    await sh.getRange(1, R, cases.length, 1).setFormulas(cases.map((c) => [c[1]]));
    // 等重算落定
    let last = -1;
    for (let k = 0; k < 30; k++) {
      await wait(400);
      const v = sh.getRange(1, R, cases.length, 1).getValues() || [];
      const filled = v.filter((x) => x && String(x[0]) !== '').length;
      if (filled >= cases.length) break;
      if (filled === last && k > 4) break;
      last = filled;
    }

    const vals = sh.getRange(1, R, cases.length, 1).getValues() || [];
    // 快照原始值：v 是存储值，t 是类型，s 是样式
    let snap = null;
    try {
      snap = wb.getSnapshot ? wb.getSnapshot() : (wb.save ? wb.save().getSnapshot() : null);
    } catch (e) { res.eSnap = String(e.message).slice(0, 120); }
    const sid = sh.getSheetId ? sh.getSheetId() : null;
    let cells = null;
    if (snap && snap.sheets) {
      const s = snap.sheets.find((x) => x.id === sid) || snap.sheets[0];
      cells = s && s.cellData;
    }

    cases.forEach((c, i) => {
      const row = 1 + i;
      const o = { got: vals[i] ? vals[i][0] : undefined, typeof: typeof (vals[i] ? vals[i][0] : undefined) };
      try { o.display = sh.getRange(row, R).getDisplayValue ? sh.getRange(row, R).getDisplayValue() : undefined; } catch (e) { o.eDisp = String(e.message).slice(0, 80); }
      try { o.numFmt = sh.getRange(row, R).getNumberFormat ? sh.getRange(row, R).getNumberFormat() : undefined; } catch (e) { o.eFmt2 = String(e.message).slice(0, 80); }
      if (cells && cells[row] && cells[row][R]) {
        const cell = cells[row][R];
        o.rawV = cell.v;
        o.rawT = cell.t;
        o.rawS = cell.s ? JSON.stringify(cell.s).slice(0, 160) : cell.s;
        o.rawF = cell.f ? String(cell.f).slice(0, 60) : cell.f;
      }
      o.expect = c[2];
      res[c[0]] = o;
    });

    try { sh.getRange(1, R, cases.length, 1).clearContent(); } catch (e) {}
    return res;
  }, CASES);

  console.log(JSON.stringify(out, null, 2));
  await b.close();
})();
