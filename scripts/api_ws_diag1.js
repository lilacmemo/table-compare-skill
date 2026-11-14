// 探针自检：把「我可能写错」的几处真实结构打出来，避免把写法错误判成产品缺失。
//   node tools/probe_ws_diag.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PORT = 8463;
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
    const sh = sp.getActiveSheet(), us = wb.getActiveSheet();
    const tick = (ms) => new Promise((r) => setTimeout(r, ms || 150));
    const R = {};
    const keys = (o) => { try { return Object.keys(o).slice(0, 25); } catch (e) { return 'ERR:' + e.message; } };
    const proto = (o) => {
      const s = new Set(); let c = o;
      while (c && c !== Object.prototype) { Object.getOwnPropertyNames(c).forEach((k) => s.add(k)); c = Object.getPrototypeOf(c); }
      return [...s].sort().slice(0, 60);
    };

    // 0) 主表状态（前一轮跑批可能把活动表切走 / 内容清空）
    R.baseline = { sjsActive: sh.name(), uniActive: us.getSheetName(), uniSheets: wb.getNumSheets() };

    // 1) SJS search 返回结构
    sh.setValue(1, 0, 'ALPHA'); sh.setValue(2, 0, 'alpha'); sh.setValue(3, 0, 'Alpha test');
    const sc = new SJS.Search.SearchCondition();
    sc.searchString = 'ALPHA';
    sc.searchFlags = SJS.Search.SearchFlags.ignoreCase;
    sc.searchTarget = SJS.Search.SearchFoundFlags.cellText;
    sc.startSheetIndex = 0; sc.endSheetIndex = 0;
    const sr = sp.search(sc);
    R.sjsSearch = { keys: keys(sr), foundCount: sr && sr.foundCount, proto: sr ? proto(sr).slice(0, 30) : null };
    try { R.sjsSearch.sample = sr && sr.getFoundCell ? JSON.stringify(sr.getFoundCell(0)) : 'no getFoundCell'; } catch (e) { R.sjsSearch.sample = 'ERR ' + e.message; }

    // 2) rowOutlines 真实结构与语义
    const os = new SJS.Worksheet('__d1'); sp.addSheet(sp.getSheetCount(), os); os.setRowCount(50);
    os.rowOutlines.group(2, 5);
    R.outline = { proto: proto(os.rowOutlines).slice(0, 40) };
    try { const f = os.rowOutlines.find(3); R.outline.find3 = f ? JSON.stringify(f) : 'null'; } catch (e) { R.outline.find3 = 'ERR ' + e.message; }
    try { R.outline.find3_l0 = JSON.stringify(os.rowOutlines.find(3, 0)); } catch (e) { R.outline.find3_l0 = 'ERR'; }
    try { R.outline.all = JSON.stringify(os.rowOutlines.toJSON ? os.rowOutlines.toJSON() : os.rowOutlines); } catch (e) { R.outline.all = 'ERR ' + e.message; }
    try { R.outline.expandRet = String(os.rowOutlines.expand(1, false)); } catch (e) { R.outline.expandRet = 'ERR'; }
    // 分组是否真的落在 XML 上（用 toJSON 看）
    try { R.outline.json = JSON.stringify(os.toJSON().rowOutlines || os.toJSON().outlines || 'NO_KEY'); } catch (e) { R.outline.json = 'ERR'; }

    // 3) undo：setValue 在 allowUndo 下是否进栈
    sp.options.allowUndo = true;
    const us2 = new SJS.Worksheet('__d2'); sp.addSheet(sp.getSheetCount(), us2);
    sp.setActiveSheet('__d2');
    sp.undoManager().clear();
    us2.setValue(1, 0, 'BEFORE');
    us2.setValue(1, 0, 'AFTER');
    R.undo = { stackLen: sp.undoManager().getUndoStack().length, stack: JSON.stringify(sp.undoManager().getUndoStack()).slice(0, 300) };
    R.undo.undoRet = sp.undoManager().undo();
    R.undo.after = String(us2.getValue(1, 0));

    // 4) getSelections 元素结构
    us2.setSelection(3, 2, 4, 2);
    const sels = us2.getSelections();
    R.sel = { len: sels.length, first: sels[0] ? JSON.stringify(sels[0]) : null, proto: sels[0] ? proto(sels[0]).slice(0, 20) : null };

    // 5) getUsedRange 各 type
    R.used = {};
    for (const t of ['undefined', 'data', 'style', 'all']) {
      try {
        const arg = t === 'undefined' ? undefined : SJS.UsedRangeType[t];
        const r = arg === undefined ? us2.getUsedRange() : us2.getUsedRange(arg);
        R.used[t] = r ? { row: r.row, col: r.col, rc: r.rowCount, cc: r.colCount } : 'null';
      } catch (e) { R.used[t] = 'ERR ' + e.message; }
    }

    // 6) Univer FRange 的行列读取方法
    const ur = us.getRange(1, 0, 3, 2);
    R.uniRange = { proto: proto(ur).filter((k) => /row|col|width|height|range|data/i.test(k)) };
    try { R.uniRange.getRow = ur.getRow(); R.uniRange.getColumn = ur.getColumn(); } catch (e) { R.uniRange.getRow = 'ERR ' + e.message; }
    try { R.uniRange.dr = JSON.stringify(us.getDataRange ? { r: us.getDataRange().getRow(), c: us.getDataRange().getColumn() } : 'no getDataRange'); } catch (e) { R.uniRange.dr = 'ERR ' + e.message; }

    // 7) 冻结语义
    R.freeze = {};
    try { us.setFrozenRows(2, 5); R.freeze.afterSet25 = JSON.stringify(us.getFreeze()); R.freeze.getFrozenRows = us.getFrozenRows(); } catch (e) { R.freeze.err = 'ERR ' + e.message; }
    try { us.cancelFreeze(); us.setFrozenRows(3); R.freeze.afterSet3 = JSON.stringify(us.getFreeze()); } catch (e) { R.freeze.err2 = 'ERR ' + e.message; }
    try { us.cancelFreeze(); } catch (e) {}

    // 8) find-replace 真实用法
    R.find = { proto: proto(wb).filter((k) => /find|search/i.test(k)) };
    try {
      const tf = await Promise.race([
        window.univerAPI.createTextFinderAsync('ALPHA'),
        new Promise((r) => setTimeout(() => r('TIMEOUT'), 4000)),
      ]);
      if (tf === 'TIMEOUT') { R.find.create = 'TIMEOUT'; }
      else if (!tf) { R.find.create = 'null'; }
      else {
        R.find.create = 'ok'; R.find.tfProto = proto(tf);
        const c1 = await Promise.race([tf.ensureCompleteAsync(), new Promise((r) => setTimeout(() => r('TIMEOUT'), 4000))]);
        R.find.ensure = c1 === 'TIMEOUT' ? 'TIMEOUT' : JSON.stringify(c1).slice(0, 200);
        const all = await Promise.race([Promise.resolve(tf.findAll()), new Promise((r) => setTimeout(() => r('TIMEOUT'), 4000))]);
        R.find.findAll = all === 'TIMEOUT' ? 'TIMEOUT' : (all ? all.length + ' hits' : 'null');
      }
    } catch (e) { R.find.err = 'ERR ' + e.message; }
    R.find.activeSheet = wb.getActiveSheet().getSheetName();

    // 9) executeCommand 真实参数（set-range-values）
    R.cmd = {};
    try {
      await Promise.race([
        window.univerAPI.executeCommand('sheet.command.set-range-values', {
          range: { startRow: 16, startColumn: 0, endRow: 16, endColumn: 0 },
          cellValue: { v: 'CMD' },
        }),
        new Promise((r) => setTimeout(() => r('TIMEOUT'), 3000)),
      ]);
      R.cmd.cellValue = String(us.getRange(16, 0).getValue());
    } catch (e) { R.cmd.cellValue = 'ERR ' + e.message; }
    // 换一种参数形态
    try {
      await Promise.race([
        window.univerAPI.executeCommand('sheet.command.set-range-values', {
          range: { startRow: 17, startColumn: 0, endRow: 17, endColumn: 0 },
          value: 'CMD2',
        }),
        new Promise((r) => setTimeout(() => r('TIMEOUT'), 3000)),
      ]);
      R.cmd.value = String(us.getRange(17, 0).getValue());
    } catch (e) { R.cmd.value = 'ERR ' + e.message; }

    // 10) save() 快照的 id 处理
    try {
      const snap = wb.save();
      R.snap = { id: snap.id, keys: keys(snap) };
      const snap2 = JSON.parse(JSON.stringify(snap));
      delete snap2.id;
      const wb2 = window.univerAPI.createWorkbook(snap2, { makeCurrent: false });
      R.snap.reCreated = !!wb2;
      try { window.univerAPI.disposeUnit(wb2.getId()); } catch (e) {}
    } catch (e) { R.snap = { err: 'ERR ' + e.message }; }

    // 11) 背景图 / drawing 插件
    R.drawing = { hasSetBg: typeof us.setBackgroundImage, hasDrawingPlugin: !!window.univerAPI.Enum };

    await tick(200);
    // 清理临时表
    for (let i = sp.getSheetCount() - 1; i >= 0; i--) {
      if (String(sp.getSheet(i).name()).indexOf('__d') === 0) { try { sp.removeSheet(i); } catch (e) {} }
    }
    return R;
  });

  console.log(JSON.stringify(out, null, 2));
  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
