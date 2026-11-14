// 第二轮自检：定位上一轮剩下的疑点
//   node tools/probe_ws_diag2.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PORT = 8464;
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css', '.png': 'image/png', '.ttf': 'font/ttf' };

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end(); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

(async () => {
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => console.log('[pageerror]', e.message));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__spread && window.__wb && window.univerAPI, null, { timeout: 90000 });
  await page.waitForTimeout(1200);

  const out = await page.evaluate(async () => {
    const SJS = window.GC.Spread.Sheets;
    const sp = window.__spread, wb = window.__wb;
    const us = wb.getActiveSheet();
    const R = {};
    const race = (p, ms) => Promise.race([Promise.resolve(p), new Promise((r) => setTimeout(() => r('TIMEOUT'), ms || 4000))]);

    // ---- 1. Univer find：跳过 ensureCompleteAsync，直接 findAll / findNext ----
    R.find = {};
    const seed = [['ALPHA', 10], ['alpha', 20], ['Alpha test', 30], ['ALPHABET', 40], ['a*b', 50]];
    for (let i = 0; i < seed.length; i++) {
      us.getRange(i + 1, 0).setValue(seed[i][0]);
      us.getRange(i + 1, 1).setValue(seed[i][1]);
    }
    us.getRange(7, 1).setFormula('=SUM(B2:B6)');
    await new Promise((r) => setTimeout(r, 250));
    try {
      const tf = await race(window.univerAPI.createTextFinderAsync('ALPHA'), 4000);
      R.find.created = typeof tf === 'object' && tf !== 'TIMEOUT';
      if (R.find.created) {
        const a1 = await race(Promise.resolve().then(() => tf.findAll()), 4000);
        R.find.findAllRaw = a1 === 'TIMEOUT' ? 'TIMEOUT' : (Array.isArray(a1) ? a1.length : String(a1));
        const n1 = await race(Promise.resolve().then(() => tf.findNext()), 4000);
        R.find.findNext = n1 === 'TIMEOUT' ? 'TIMEOUT' : (n1 ? JSON.stringify({ r: n1.getRow(), c: n1.getColumn() }) : 'null');
        const en = await race(Promise.resolve().then(() => tf.ensureCompleteAsync()), 4000);
        R.find.ensure = en === 'TIMEOUT' ? 'TIMEOUT' : String(JSON.stringify(en)).slice(0, 120);
        const a2 = await race(Promise.resolve().then(() => tf.findAll()), 4000);
        R.find.findAllAfterEnsure = a2 === 'TIMEOUT' ? 'TIMEOUT' : (Array.isArray(a2) ? a2.length : String(a2));
      }
    } catch (e) { R.find.err = 'ERR ' + e.message; }

    // ---- 2. 背景图 ----
    R.bg = {};
    try { R.bg.hasSetter = typeof us.setBackgroundImage; } catch (e) {}
    try { us.setBackgroundImage('./fonts/bg.png'); R.bg.set = 'ok'; } catch (e) { R.bg.set = 'ERR ' + e.message; }
    try { R.bg.get = String(us.getBackgroundImage ? us.getBackgroundImage() : 'no getter'); } catch (e) { R.bg.get = 'ERR ' + e.message; }

    // ---- 3. outline：getLevel / getState 真实读数 ----
    const os = new SJS.Worksheet('__d3'); sp.addSheet(sp.getSheetCount(), os); os.setRowCount(50);
    R.outline = {};
    os.rowOutlines.group(2, 5);
    R.outline.getLevel3 = os.rowOutlines.getLevel(3);
    R.outline.getLevel0 = os.rowOutlines.getLevel(0);
    R.outline.getMaxLevel = os.rowOutlines.getMaxLevel();
    R.outline.getState3 = String(os.rowOutlines.getState(3));
    os.rowOutlines.expand(1, false);
    R.outline.stateAfterCollapse = String(os.rowOutlines.getState(3));
    R.outline.jsonAfter = JSON.stringify(os.rowOutlines.toJSON().itemsData);
    os.rowOutlines.ungroupRange(2, 5);
    R.outline.jsonAfterUngroup = JSON.stringify(os.rowOutlines.toJSON().itemsData);
    // 嵌套两级
    const os2 = new SJS.Worksheet('__d4'); sp.addSheet(sp.getSheetCount(), os2); os2.setRowCount(50);
    os2.rowOutlines.group(1, 8);
    os2.rowOutlines.group(3, 3);
    R.outline.nested = JSON.stringify(os2.rowOutlines.toJSON().itemsData);
    R.outline.nestedLevels = [1, 2, 3, 4, 5, 6].map((r) => os2.rowOutlines.getLevel(r));

    // ---- 4. SJS search 是否只返回首个命中：换条件验证 ----
    const ss = new SJS.Worksheet('__d5'); sp.addSheet(sp.getSheetCount(), ss);
    ['ALPHA', 'alpha', 'Alpha test', 'ALPHABET', 'a*b'].forEach((v, i) => ss.setValue(i + 1, 0, v));
    const mk = (s, flags, target, sheetIdx) => {
      const c = new SJS.Search.SearchCondition();
      c.searchString = s; c.searchFlags = flags; c.searchTarget = target;
      c.startSheetIndex = sheetIdx === undefined ? 0 : sheetIdx;
      c.endSheetIndex = sheetIdx === undefined ? sp.getSheetCount() - 1 : sheetIdx;
      return c;
    };
    R.search = {};
    const r1 = sp.search(mk('ALPHA', SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellText));
    R.search.ignoreCase = r1 ? { row: r1.foundRowIndex, col: r1.foundColumnIndex, sheet: r1.foundSheetIndex, str: r1.foundString, flag: r1.searchFoundFlag } : null;
    const r2 = sp.search(mk('ALPHA', SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellText));
    R.search.caseSensitive = r2 ? { row: r2.foundRowIndex, str: r2.foundString } : null;
    const r3 = sp.search(mk('alpha', SJS.Search.SearchFlags.ignoreCase | SJS.Search.SearchFlags.exactMatch, SJS.Search.SearchFoundFlags.cellText));
    R.search.exact = r3 ? { row: r3.foundRowIndex, str: r3.foundString } : null;
    const r4 = sp.search(mk('a*b', SJS.Search.SearchFlags.useWildCards | SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellText));
    R.search.wildcard = r4 ? { row: r4.foundRowIndex, str: r4.foundString } : null;
    const r5 = sp.search(mk(/^alpha$/i, SJS.Search.SearchFlags.none, SJS.Search.SearchFoundFlags.cellText));
    R.search.regex = r5 ? { row: r5.foundRowIndex, str: r5.foundString } : 'null';
    const ssIdx = (() => { for (let i = 0; i < sp.getSheetCount(); i++) if (sp.getSheet(i).name() === '__d5') return i; return 0; })();
    const r6 = sp.search(mk('NUM', SJS.Search.SearchFlags.ignoreCase, SJS.Search.SearchFoundFlags.cellFormula, ssIdx));
    R.search.formula = r6 ? { row: r6.foundRowIndex, str: r6.foundString } : 'null';
    // 是否有 count / next 类接口
    R.search.resultProto = r1 ? Object.getOwnPropertyNames(Object.getPrototypeOf(r1)) : null;

    // ---- 5. SJS undo：走命令而非 setValue ----
    R.undo = {};
    sp.options.allowUndo = true;
    const undoSh = new SJS.Worksheet('__d6'); sp.addSheet(sp.getSheetCount(), undoSh);
    sp.setActiveSheet('__d6');
    sp.undoManager().clear();
    const exec = (cmd, opts) => sp.commandManager().execute(Object.assign({ cmd, sheetName: '__d6' }, opts));
    R.undo.editCellRet = exec('editCell', { row: 1, col: 0, newValue: 'CMDVAL' });
    R.undo.afterEdit = String(undoSh.getValue(1, 0));
    R.undo.stackLen = sp.undoManager().getUndoStack().length;
    R.undo.undoRet = sp.undoManager().undo();
    R.undo.afterUndo = String(undoSh.getValue(1, 0));
    if (!R.undo.stackLen) {
      // 试 editCell 的另一种参数写法
      sp.undoManager().clear();
      undoSh.setValue(2, 0, 'X');
      R.undo.afterSetValueStack = sp.undoManager().getUndoStack().length;
    }

    // ---- 6. Univer executeCommand 参数修正 ----
    R.cmd = {};
    try {
      await race(window.univerAPI.executeCommand('sheet.command.set-range-values', {
        value: { v: 'CMD_OK' },
        range: { startRow: 20, startColumn: 0, endRow: 20, endColumn: 0 },
      }), 3000);
      R.cmd.ok = String(us.getRange(20, 0).getValue());
    } catch (e) { R.cmd.err = 'ERR ' + e.message; }

    // ---- 7. 清理 ----
    for (let i = sp.getSheetCount() - 1; i >= 0; i--) {
      if (String(sp.getSheet(i).name()).indexOf('__d') === 0) { try { sp.removeSheet(i); } catch (e) {} }
    }
    return R;
  });

  console.log(JSON.stringify(out, null, 2));
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
