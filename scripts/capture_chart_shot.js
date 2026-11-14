// 逐 case 抓「单张图表」截图（SpreadJS / Univer 两侧）→ out/shots/chart/{id}-sjs.png / -uni.png
//
// 两侧取图方式故意不同（能力差异决定的，不是偷懒）：
//   SpreadJS : chart.toImageSrcAsync() —— 图表控件自己渲染出的 PNG，和 Excel 的 Chart.Export 同源。
//              绕开 Designer 的三个致命坑（见下），也不会有网格线/邻居图/样本自带灰条混进来。
//   Univer   : 没有等价 API（导出是 undefined），只能截视口 → 按「锚点单元格 + 偏移」算剪裁框。
//
// Designer 截图路线的三个坑（上一版踩过，别改回去）：
//  1) workbook.setActiveSheet(i) 不生效：activeSheetIndex 不变、非活动表 getCellRect 一律返回 {}。
//     要切表得点一下网格再按 Ctrl+PageDown/PageUp。本脚本已不用截图路线，故不需要。
//  2) 滚动是异步渲染，滚完立刻 getCellRect 会拿到空对象；而且要滚十几屏才能到工作表下半部分，
//     420ms 等不够 → 图表没渲染完就截，出来是空白网格（B1/B9/C6/G4/G5/H*/I*/K1 曾全白）。
//  3) getCellRect(0,0) 在滚走之后返回 {}，得在滚动前缓存成 window.__R0_SJS。
//
// Univer 坐标换算：
//   scrollToCell(row,col,0,0) 把锚点单元格推到内容区左上角（左上角 = 主画布 + 行表头 46 / 列表头 20），
//   但最右列滚不过去（A7 在第 24 列，scrollToCell 夹在 sheetViewStartColumn=21, offsetX=9.4），
//   所以还要减去「首列 + 滚动偏移」的绝对坐标差。
//
// 用法：
//   node tools/capture_chart_shot.js              # 全部 91 个 case
//   node tools/capture_chart_shot.js A5,B1,H12    # 指定 case
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const SAMPLE = process.env.CHART_SAMPLE || 'samples/图表用例载体-逐项配置.xlsx';
const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'out', 'shots', 'chart');
const SCALE = 2;   // Univer 侧 2 倍存图，报告里缩一半显示更清晰
const PAD = 6;     // 剪裁外扩（左/上/右，CSS px）
// 下边界只留 1px：样本里每张图正下方 2pt 处有一个灰色说明文本框（样本自带的 case 标注），
// 留太多会把那条灰带框进来，而 Excel 的 Chart.Export 只导出图表本体、没有这条灰带，三方图会不一致。
const PAD_B = 1;

const only = (process.argv.slice(2).find((a) => !a.startsWith('--')) || '').split(',').filter(Boolean);

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 2240, height: 1500 }, deviceScaleFactor: SCALE });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => console.log('PAGEERROR', e.message.slice(0, 160)));
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });
  await p.evaluate((f) => APP.loadSample(f, '图表用例载体-逐项配置.xlsx'), SAMPLE);
  await p.waitForTimeout(9000);
  // 收起结果面板，把两栏拉到最高
  await p.evaluate(() => { const rp = document.getElementById('rp'); if (rp && !rp.classList.contains('min')) document.getElementById('btnFold').click(); });
  await p.waitForTimeout(500);

  const list = await p.evaluate(() => {
    const sp = APP.sjs;
    let idx = -1;
    for (let i = 0; i < sp.getSheetCount(); i++) if (sp.getSheet(i).name() === '用例') idx = i;
    const sh = sp.getSheet(idx);
    const wb = window.univerAPI.getActiveWorkbook();
    const ws = wb.getSheets().find((s) => s.getSheetName() === '用例');
    wb.setActiveSheet(ws);
    // 样本里每张图正下方 2pt 有个灰色说明文本框（Shapes.AddTextbox）。Excel 的 Chart.Export 不含它，
    // 不删就会压进 Univer 的剪裁框（A7 面板底部那条灰带就是这么来的）。
    (ws.getShapes ? ws.getShapes() || [] : []).forEach((x) => { try { x.remove(); } catch (e) {} });
    const ids = sh.charts.all().map((c) => c.name());        // 两侧顺序一致，用 SJS 的名字作主键
    // Univer 全部图表原始位置快照：截图时把「非目标图」全部挪走，剪裁框里就只剩目标图，
    // 免去「按读回尺寸算重叠邻居」这套不可靠的几何推算（实测读回尺寸和渲染尺寸对不上）。
    const uniPos = ws.getCharts().map((x) => { const p = x.getInfo().position || {}; return { x: p.x, y: p.y }; });
    return {
      sjs: sh.charts.all().map((c) => ({
        id: c.name(),
        row: c.startRow(), col: c.startColumn(), ro: c.startRowOffset(), co: c.startColumnOffset(),
        w: c.width(), h: c.height(),
      })),
      ids,
      uniPos,
      uniCount: ws.getCharts().length,
      uniHeader: (() => { const r = ws.getRange(0, 0).getCellRect(); return { x: r.x, y: r.y }; })(),
      zoom: { sjs: sh.zoom(), uni: ws.getZoom ? ws.getZoom() : 1 },
    };
  });
  const pairs = list.sjs.map((s, i) => ({ id: s.id, sjs: s, uniIndex: i }));
  const todo = only.length ? pairs.filter((x) => only.includes(x.id)) : pairs;
  console.log('charts', pairs.length, 'uni', list.uniCount, 'todo', todo.length,
    'zoom', JSON.stringify(list.zoom), 'uniHeader', JSON.stringify(list.uniHeader));
  if (list.uniCount !== pairs.length) console.log('  ! 两侧图表数不一致，按下标对位可能错位');

  // Univer 内容区原点（页面坐标）
  const origin = await p.evaluate((hdr) => {
    const pane = document.querySelector('#paneUni .host');
    const pr = pane.getBoundingClientRect();
    let best = null;
    pane.querySelectorAll('canvas').forEach((c) => {
      const r = c.getBoundingClientRect();
      if (r.width < 500 || r.height < 300) return;
      if (!best || r.width * r.height > best.w * best.h) best = { x: r.x, y: r.y, w: r.width, h: r.height };
    });
    return { x: best.x + hdr.x, y: best.y + hdr.y, w: best.w - hdr.x, h: best.h - hdr.y };
  }, list.uniHeader);
  console.log('ORIGIN', JSON.stringify(origin));

  const report = [];
  for (const it of todo) {
    // ---------- SpreadJS：图表自带导出 ----------
    const dataUrl = await p.evaluate(async (cid) => {
      const sp = APP.sjs;
      let idx = -1;
      for (let i = 0; i < sp.getSheetCount(); i++) if (sp.getSheet(i).name() === '用例') idx = i;
      const sh = sp.getSheet(idx);
      const c = sh.charts.all().find((x) => x.name() === cid);
      if (!c) return null;
      try { return await c.toImageSrcAsync(); } catch (e) { return 'ERR:' + e; }
    }, it.id);
    if (typeof dataUrl === 'string' && dataUrl.startsWith('data:image/png;base64,')) {
      const buf = Buffer.from(dataUrl.split(',')[1], 'base64');
      const file = path.join(OUT, `${it.id}-sjs.png`);
      fs.writeFileSync(file, buf);
      const sz = pngSize(buf);
      report.push({ id: it.id, engine: 'sjs', method: 'toImageSrcAsync', size: sz });
      if (process.env.SHOT_DEBUG) console.log('   ', it.id, 'sjs export', sz);
    } else {
      console.log('  ! sjs export failed', it.id, String(dataUrl).slice(0, 60));
    }

    // ---------- Univer：截视口 ----------
    const u = it.sjs;   // 两侧锚点一致，直接用 SJS 读到的 row/col/offset
    await p.evaluate(({ row, col }) => {
      const wb = window.univerAPI.getActiveWorkbook();
      const ws = wb.getSheets().find((s) => s.getSheetName() === '用例');
      wb.setActiveSheet(ws);
      ws.scrollToCell(row, col, 0, 0);
    }, { row: u.row, col: u.col });
    await p.waitForTimeout(420);
    const rel = await p.evaluate(({ row, col, idx, pos }) => {
      try {
        const wb = window.univerAPI.getActiveWorkbook();
        const ws = wb.getSheets().find((s) => s.getSheetName() === '用例');
        const st = ws.getScrollState();
        const r = ws.getRange(row, col).getCellRect();
        const s = ws.getRange(st.sheetViewStartRow || 0, st.sheetViewStartColumn || 0).getCellRect();
        const all = ws.getCharts();
        const t = all[idx].getInfo();
        return { x: r.x - s.x - (st.offsetX || 0), y: r.y - s.y - (st.offsetY || 0),
                 size: { w: t.size.width, h: t.size.height }, raw: r, start: s, state: st };
      } catch (e) { return null; }
    }, { row: u.row, col: u.col, idx: it.uniIndex, pos: list.uniPos });
    const d = rel || { x: 0, y: 0 };
    const z = list.zoom.uni || 1;
    const o = origin;
    const uw = (d.size && d.size.w) || u.w;
    const uh = (d.size && d.size.h) || u.h;
    // 右边/下边多留 MARGIN：读回的 size 未必等于实际渲染尺寸（实测偏大），留余量后靠 compose 去白边收边，
    // 这样即使实际渲染偏小也不会把图裁掉，而且周围的邻居图已经挪走 → 白底可安全裁。
    const MARGIN = 40;
    const wantX = o.x + d.x + u.co * z - PAD;
    const wantY = o.y + d.y + u.ro * z - PAD;
    const cw = Math.round(uw * z + PAD * 2 + MARGIN);
    const chh = Math.round(uh * z + PAD + PAD_B + MARGIN);
    const maxW = Math.round(o.x + o.w - wantX);
    const maxH = Math.round(o.y + o.h - wantY);
    const clip = { x: Math.max(o.x, Math.round(wantX)), y: Math.max(o.y, Math.round(wantY)),
                   width: Math.max(1, Math.min(cw, maxW)), height: Math.max(1, Math.min(chh, maxH)) };
    if (!(clip.width > 2 && clip.height > 2)) {
      console.log('  ! skip uni', it.id, 'clip', JSON.stringify(clip), 'rel', JSON.stringify(d));
      continue;
    }
    if (process.env.SHOT_DEBUG) console.log('   ', it.id, 'uni rel', JSON.stringify(d), 'clip', JSON.stringify(clip));
    // 把除目标图以外的全部图表挪到屏幕外（记原坐标），截完立刻原样挪回
    await p.evaluate(({ idx, pos }) => {
      const wb = window.univerAPI.getActiveWorkbook();
      const ws = wb.getSheets().find((s) => s.getSheetName() === '用例');
      const all = ws.getCharts();
      all.forEach((c, i) => {
        if (i === idx) return;
        // 必须往「正方向」挪：setAbsolutePosition 传负坐标会被夹到 0，90 张图会全叠到左上角，剪裁框里一片花
        try { c.setAbsolutePosition({ x: (pos[i] || {}).x + 20000, y: (pos[i] || {}).y + 20000 }); } catch (e) {}
      });
    }, { idx: it.uniIndex, pos: list.uniPos });
    await p.waitForTimeout(320);
    await p.screenshot({ path: path.join(OUT, `${it.id}-uni.png`), clip });
    await p.evaluate(({ idx, pos }) => {
      const wb = window.univerAPI.getActiveWorkbook();
      const ws = wb.getSheets().find((s) => s.getSheetName() === '用例');
      const all = ws.getCharts();
      all.forEach((c, i) => {
        if (i === idx) return;
        try { c.setAbsolutePosition({ x: (pos[i] || {}).x, y: (pos[i] || {}).y }); } catch (e) {}
      });
    }, { idx: it.uniIndex, pos: list.uniPos });
    await p.waitForTimeout(220);
    report.push({ id: it.id, engine: 'uni', method: 'screenshot', clip, rel: d, reported: { w: uw, h: uh } });
    if (todo.length <= 20) console.log('shot', it.id);
  }
  fs.writeFileSync(path.join(ROOT, 'out', 'data', 'chart-shots.json'), JSON.stringify({ origin, list: report }, null, 1), 'utf8');
  console.log('wrote', report.length, 'shots ->', OUT);
  await b.close();
})();

// 读 PNG IHDR 拿宽高（base64 → Buffer）
function pngSize(buf) {
  if (buf.length < 24) return null;
  return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
}
