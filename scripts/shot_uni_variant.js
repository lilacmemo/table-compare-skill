// Univer 侧单图截图：用「只保留该图的变体工作簿」，一张表只有一张图，剪裁框里不可能混进邻居。
//
// 为什么不用共享样本 + 藏邻居（试过，都失败）：
//   · chart.remove() 掉其余图表后，幸存的那张不再绘制（getCharts() 数量、size、position 都对，画布上就是空的）；
//   · chart.setAbsolutePosition() 传负坐标会被夹到 0，90 张图全叠到左上角，剪裁框一片花。
// 变体工作簿由 tools/build_single_chart_variants.py 生成（drawing1.xml 只留该图那个 twoCellAnchor，
// 顺带丢掉样本自带的灰色说明文本框）。
//
// 坐标换算：getCellRect(row,col) 是整表绝对坐标（不随滚动变），减掉「视口左上角那格 + 滚动偏移」就是屏幕位置。
// 图表尺寸用 getInfo().size —— 注意这个值在刚导入、布局还没稳定时读会偏大（A7 第一次读到 792，稳定后 721.6），
// 所以滚动后要等一会儿再读。
//
// 用法：node tools/shot_uni_variant.js A5,A7,...   （不给参数就跑 out/data/chart-variants.json 全部）
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'out', 'shots', 'chart');
const SCALE = 2;
const PAD = 6;
const PAD_B = 1;
const MARGIN = 0;    // 不再留余量：实测 Univer 实际渲染尺寸 == getInfo().size，留余量只会把下方的网格线框进来

const only = (process.argv.slice(2).find((a) => !a.startsWith('--')) || '').split(',').filter(Boolean);

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const idx = JSON.parse(fs.readFileSync(path.join(ROOT, 'out', 'data', 'chart-variants.json'), 'utf8'));
  const todo = only.length ? idx.filter((x) => only.includes(x.id)) : idx;
  console.log('todo', todo.length);

  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 2240, height: 1500 }, deviceScaleFactor: SCALE });
  const report = [];

  for (const it of todo) {
    const p = await ctx.newPage();
    p.on('pageerror', (e) => console.log('  PAGEERROR', e.message.slice(0, 120)));
    await p.goto(BASE, { waitUntil: 'networkidle' });
    await p.waitForFunction(() => window.APP && window.APP.booted === true, null, { timeout: 90000 });
    await p.evaluate(({ f, name }) => APP.loadSample(f, name), { f: it.file, name: it.file.split('/').pop() });
    await p.waitForTimeout(7000);
    // 收起结果面板，把编辑区拉到最高
    await p.evaluate(() => { const rp = document.getElementById('rp'); if (rp && !rp.classList.contains('min')) document.getElementById('btnFold').click(); });
    await p.waitForTimeout(400);

    const pre = await p.evaluate(() => {
      const pane = document.querySelector('#paneUni .host');
      let best = null;
      pane.querySelectorAll('canvas').forEach((c) => {
        const r = c.getBoundingClientRect();
        if (r.width < 500 || r.height < 300) return;
        if (!best || r.width * r.height > best.w * best.h) best = { x: r.x, y: r.y, w: r.width, h: r.height };
      });
      const ws = window.univerAPI.getActiveWorkbook().getSheets().find((s) => s.getSheetName() === '用例');
      window.univerAPI.getActiveWorkbook().setActiveSheet(ws);
      const hs = ws.getRange(0, 0).getCellRect();
      const g = ws.getCharts().length ? ws.getCharts()[0].getInfo() : null;
      // 单图变体的「已用范围」可能只到数据末尾，而图表锚在第 193 行（C6）之类的地方 ——
      // 不撑开范围就滚不过去，图会停在视口下方（size 也读成 1px 的退化值）。
      if (g) { try { ws.getRange(Math.max(g.anchor.row + 60, 900), 30).setValue('.'); } catch (e) {} }
      if (g) ws.scrollToCell(g.anchor.row, g.anchor.column, 0, 0);
      return { canvas: best, hdr: { x: hs.x, y: hs.y }, n: ws.getCharts().length, anchor: g && g.anchor };
    });
    if (!pre.n) { console.log('  ! 变体里没有图表', it.id); await p.close(); continue; }
    await p.waitForTimeout(1600);
    // 只用 position 算屏幕位置，不去 getRange(anchor.row, anchor.column)：
    // 图表的锚点行列可能落在工作表数据范围之外（C6 在第 193 行），getRange 会直接抛异常。
    // position 与 getCellRect 都是同一套绝对坐标，减掉「视口首格 + 滚动偏移」就是屏幕位置。
    const info = await p.evaluate(() => {
      const ws = window.univerAPI.getActiveWorkbook().getSheets().find((s) => s.getSheetName() === '用例');
      const g = ws.getCharts()[0].getInfo();
      const st = ws.getScrollState();
      const s = ws.getRange(st.sheetViewStartRow || 0, st.sheetViewStartColumn || 0).getCellRect();
      const pos = g.position || { x: 0, y: 0 };
      return {
        size: g.size, pos, anchor: g.anchor, state: st,
        rel: { x: pos.x - s.x - (st.offsetX || 0), y: pos.y - s.y - (st.offsetY || 0) },
      };
    });

    const o = { x: pre.canvas.x + pre.hdr.x, y: pre.canvas.y + pre.hdr.y,
                w: pre.canvas.w - pre.hdr.x, h: pre.canvas.h - pre.hdr.y };
    const wantX = o.x + info.rel.x - PAD;      // rel 已经含了锚点偏移（position 就是图表左上角）
    const wantY = o.y + info.rel.y - PAD;
    const cw = Math.round(info.size.width + PAD * 2 + MARGIN);
    const chh = Math.round(info.size.height + PAD + PAD_B + MARGIN);
    const clip = {
      x: Math.max(o.x, Math.round(wantX)), y: Math.max(o.y, Math.round(wantY)),
      width: Math.max(1, Math.min(cw, Math.round(o.x + o.w - wantX))),
      height: Math.max(1, Math.min(chh, Math.round(o.y + o.h - wantY))),
    };
    await p.screenshot({ path: path.join(OUT, `${it.id}-uni.png`), clip });
    report.push({ id: it.id, engine: 'uni', method: 'variant', clip, info });
    console.log('shot', it.id, 'size', JSON.stringify(info.size), 'clip', JSON.stringify(clip));
    await p.close();
  }
  const f = path.join(ROOT, 'out', 'data', 'chart-shots-uni.json');
  fs.writeFileSync(f, JSON.stringify(report, null, 1), 'utf8');
  console.log('wrote', report.length, '->', f);
  await b.close();
})();
