// Univer 打印输出捕获：print() 会离屏渲染分页到 .printing-canvas 的 canvas 上，
// 在 window.print() 之前把每页 canvas 抓成 PNG，就等于拿到"打印预览"。
//   NODE_PATH=../poc-compare/node_modules node tools/univer_print_capture.js
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out', 'print');
const PORT = Number(process.env.PORT || 8488);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png', '.ttf': 'font/ttf',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, url === '/' ? 'index.html' : url);
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end('not found'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  page.on('pageerror', (e) => console.log('PAGEERR', e.message));

  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.univerAPI && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1200);

  // 在两个视图都放一遍？不：只操作 Univer 面板，避免影响 SJS。
  const res = await page.evaluate(async () => {
    const u = window.univerAPI;
    const wb = window.__wb;
    const ws = wb.getActiveSheet();
    const R = { pages: [], errors: [] };

    // ---------- 造内容：中文标题 + 数字 + 百分比 + 合并 + 条件格式 + 长表 ----------
    ws.getRange('A1:H1').merge();
    ws.getRange('A1').setValue('季度销售报表 Quarterly Sales Report');
    ws.getRange('A1').setFontSize(18).setFontWeight('bold');
    const hdr = ['区域', 'Q1', 'Q2', 'Q3', 'Q4', '合计', '占比', '备注'];
    for (let c = 0; c < hdr.length; c++) ws.getRange(2, c).setValue(hdr[c]).setFontWeight('bold');
    const regions = ['华东', '华南', '华北', '西南', '华中', '东北', '西北', '东南'];
    let seed = 7;
    const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return 120 + (seed % 860); };
    for (let i = 0; i < 40; i++) {
      const r = 3 + i;
      ws.getRange(r, 0).setValue(regions[i % 8] + ' ' + (Math.floor(i / 8) + 1));
      let sum = 0;
      for (let c = 1; c <= 4; c++) { const v = rnd(); sum += v; ws.getRange(r, c).setValue(v); }
      ws.getRange(r, 5).setValue(sum);
      ws.getRange(r, 6).setValue(sum / 4000);
      ws.getRange(r, 6).setNumberFormat('0.0%');
      ws.getRange(r, 7).setValue('自动生成');
    }
    // 条件格式：数据条 + 色阶（OSS 有的能力）
    try {
      const cf = u.newConditionalFormatting();
      cf.addDataBarRule?.('B4:B43');
      u.getActiveWorkbook().getActiveSheet().getRange('C4:C43')
        .setBackgroundColor?.('#dbeafe');
    } catch (e) { R.errors.push('cf: ' + e.message); }

    // ---------- 打印配置 ----------
    try {
      wb.updatePrintConfig({
        area: u.Enum.PrintArea.CurrentSheet,
        paperSize: u.Enum.PrintPaperSize.A4,
        direction: u.Enum.PrintDirection.Landscape,
        scale: u.Enum.PrintScale.FitWidth,
        margin: u.Enum.PrintPaperMargin.Normal,
        freeze: [u.Enum.PrintFreeze.Row],
      });
      R.layout = 'ok';
    } catch (e) { R.layout = 'ERR ' + e.message; }

    try {
      // isCustomHeaderFooter 必须为 true，否则 headerFooterSetting 整块被忽略，
      // 只剩 headerFooter 里的枚举（页眉会显示 sheet 名）
      wb.updatePrintRenderConfig({
        gridlines: true,
        hAlign: u.Enum.PrintAlign.Middle,
        headerFooter: [u.Enum.PrintHeaderFooter.WorksheetTitle],
        isCustomHeaderFooter: true,
        headerFooterSetting: {
          topCenter: '季度销售报表',
          bottomLeft: '由 Univer 生成',
          // Univer 的页码占位符是 @Page / @TotalPage，不是 Excel 的 &P / &N
          bottomRight: '@Page / @TotalPage',
        },
      });
      R.render = 'ok';
    } catch (e) { R.render = 'ERR ' + e.message; }

    // ---------- 触发打印并抓 canvas ----------
    const grab = () => {
      const nodes = Array.from(document.querySelectorAll('.printing-canvas canvas, .printing-canvas-container canvas'));
      return nodes.map((cv, i) => ({
        i, w: cv.width, h: cv.height,
        data: cv.toDataURL('image/png'),
      }));
    };

    let captured = [];
    try {
      const ret = wb.print();
      if (ret && ret.then) await ret;
      // print() 内部会渲染完再调 window.print()；headless 下 window.print 是 no-op，
      // 容器会留在 DOM 里，但渲染是异步的，要轮询等 canvas 出现
      for (let t = 0; t < 20; t++) {
        await new Promise((r) => setTimeout(r, 250));
        const c = grab();
        if (c.length) { captured = c; break; }
      }
      await new Promise((r) => setTimeout(r, 800));
      if (!captured.length) captured = grab();
      R.printReturn = String(ret && ret.constructor && ret.constructor.name);
    } catch (e) { R.errors.push('print: ' + e.message); }

    R.domPrint = {
      container: !!document.querySelector('.printing-canvas-container'),
      canvases: document.querySelectorAll('.printing-canvas canvas').length,
    };
    R.pages = captured.map((c) => ({ i: c.i, w: c.w, h: c.h, data: c.data }));
    return R;
  });

  console.log('layout:', res.layout, '| render:', res.render);
  console.log('printReturn:', res.printReturn);
  console.log('domPrint:', JSON.stringify(res.domPrint));
  console.log('errors:', res.errors);
  console.log('pages:', res.pages.map((p) => `${p.i}:${p.w}x${p.h}`).join(' '));

  res.pages.forEach((p) => {
    const b64 = String(p.data).split(',')[1] || '';
    const f = path.join(OUT, `univer-print-p${p.i + 1}.png`);
    fs.writeFileSync(f, Buffer.from(b64, 'base64'));
    console.log('saved', f, p.w + 'x' + p.h);
  });

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
