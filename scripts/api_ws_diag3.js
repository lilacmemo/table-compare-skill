// 第三轮自检：FSelection 方法名 + SJS 事件名是否存在
const http = require('http'); const fs = require('fs'); const path = require('path');
const ROOT = path.resolve(__dirname, '..'); const PORT = 8465;
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css', '.png': 'image/png', '.ttf': 'font/ttf' };
function serve() { return new Promise((r) => { const s = http.createServer((q, res) => { const u = decodeURIComponent(q.url.split('?')[0]); const p = path.join(ROOT, u === '/' ? 'index.html' : u); fs.readFile(p, (e, b) => { if (e) { res.writeHead(404).end(); return; } res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' }); res.end(b); }); }); s.listen(PORT, '127.0.0.1', () => r(s)); }); }
(async () => {
  const srv = await serve();
  const { chromium } = require('playwright');
  const b = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const p = await b.newPage(); p.on('pageerror', (e) => console.log('[pageerror]', e.message));
  await p.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await p.waitForFunction(() => window.__spread && window.__wb && window.univerAPI, null, { timeout: 90000 });
  await p.waitForTimeout(1000);
  const out = await p.evaluate(async () => {
    const SJS = window.GC.Spread.Sheets; const sp = window.__spread, wb = window.__wb, us = wb.getActiveSheet();
    const R = {};
    const proto = (o) => { const s = new Set(); let c = o; while (c && c !== Object.prototype) { Object.getOwnPropertyNames(c).forEach((k) => s.add(k)); c = Object.getPrototypeOf(c); } return [...s].filter((k) => /^[a-zA-Z]/.test(k)).sort(); };
    // 1) SJS 事件名是否存在
    const evNames = ['ActiveSheetChanged', 'ActiveSheetChanging', 'SheetChanged', 'SheetChanging', 'SheetMoved',
      'CellChanged', 'RangeGroupStateChanged', 'RowChanged', 'ColumnChanged', 'EditEnded', 'SelectionChanged', 'RangeChanged'];
    R.sjsEvents = {};
    for (const n of evNames) R.sjsEvents[n] = SJS.Events[n] || 'MISSING';
    // 2) 主动触发 ActiveSheetChanged：建表、切表、切回
    const h = { activeSheet: 0, sheetChanged: 0, sheetMoved: 0 };
    sp.bind(SJS.Events.ActiveSheetChanged, () => h.activeSheet++);
    sp.bind(SJS.Events.SheetChanged, () => h.sheetChanged++);
    sp.bind(SJS.Events.SheetMoved, () => h.sheetMoved++);
    const s1 = new SJS.Worksheet('__e1'); sp.addSheet(sp.getSheetCount(), s1);
    await new Promise((r) => setTimeout(r, 120));
    const afterAdd = { ...h };
    sp.setActiveSheet('__e1');
    await new Promise((r) => setTimeout(r, 120));
    const afterActivate = { ...h };
    sp.changeSheetPosition('__e1', 0);
    await new Promise((r) => setTimeout(r, 120));
    R.sjsEventHits = { afterAdd, afterActivate, afterMove: { ...h } };
    // 3) FSelection 方法
    us.setActiveSelection(us.getRange(3, 2, 4, 2));
    await new Promise((r) => setTimeout(r, 200));
    const sel = us.getSelection();
    R.sel = { has: !!sel, proto: sel ? proto(sel) : null };
    if (sel) {
      try { const ar = sel.getActiveRange(); R.sel.activeRange = ar ? ar.getRow() + ',' + ar.getColumn() : null; } catch (e) { R.sel.activeRange = 'ERR ' + e.message; }
      try { const rs = sel.getRanges(); R.sel.ranges = Array.isArray(rs) ? rs.length : String(rs); } catch (e) { R.sel.ranges = 'ERR ' + e.message; }
    }
    // 4) FWorksheet 上选区相关方法
    R.wsSel = proto(us).filter((k) => /select|active/i.test(k));
    // 5) SheetApi 里有没有 getSelections
    for (let i = sp.getSheetCount() - 1; i >= 0; i--) if (String(sp.getSheet(i).name()).indexOf('__e') === 0) { try { sp.removeSheet(i); } catch (e) {} }
    return R;
  });
  console.log(JSON.stringify(out, null, 2));
  await b.close(); srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
