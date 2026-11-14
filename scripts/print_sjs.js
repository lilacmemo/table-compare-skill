// SpreadJS 侧：导入多元素样本 → 读回 printInfo → savePDF 落盘
//   NODE_PATH=../poc-compare/node_modules node tools/print_sjs.js
// 输出：tools/out/print/sjs.pdf + sjs-printinfo.json + 截图在 screenshots/
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out', 'print');
const PORT = Number(process.env.PORT || 8485);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const SAMPLE = '打印-PDF-多元素样本.xlsx';

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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
  await page.waitForFunction(() => window.__spread && window.univerAPI, null, { timeout: 90000 });
  await page.waitForTimeout(1200);

  const info = await page.evaluate(async (sampleName) => {
    const sp = window.__spread;

    // PDF 导出默认只嵌 14 种西文标准字体，中文会变成乱码方块 ——
    // 必须 registerFont 传 TTF ArrayBuffer，并用 fallbackFont 兜住未注册字体。
    // 注意 .ttc（simsun.ttc / msyh.ttc）不支持，只能是 .ttf/.otf。
    let fontBytes = 0;
    try {
      const fb = await fetch('/fonts/simhei.ttf').then((r) => r.arrayBuffer());
      fontBytes = fb.byteLength;
      const F = GC.Spread.Sheets.PDF.PDFFontsManager;
      F.registerFont('SimHei', { normal: fb, bold: fb });
      F.registerFont('宋体', { normal: fb, bold: fb });
      F.registerFont('等线', { normal: fb, bold: fb });
      F.fallbackFont = function () { return fb; };
    } catch (e) {
      fontBytes = -1;
    }

    const buf = await fetch('/samples/' + encodeURIComponent(sampleName)).then((r) => r.arrayBuffer());
    // v19 签名是回调式：import(file, successCallback, errorCallback, importOptions)，
    // 返回 void。写成 await sp.import(blob, opts) 不会报错但什么也不导（只剩空表）。
    const t0 = Date.now();
    await new Promise((res, rej) => {
      sp.import(new Blob([buf]), res, rej, { fileType: GC.Spread.Sheets.FileType.excel });
    });
    await new Promise((r) => setTimeout(r, 800));

    const out = { fontBytes, sheets: [], elements: [], printInfo: [] };
    sp.sheets.forEach((sh, i) => {
      out.sheets.push({ i, name: sh.name(), visible: sh.visible(), rowCount: sh.getRowCount(), colCount: sh.getColumnCount() });
      const charts = sh.charts ? sh.charts.all() : [];
      const shapes = sh.shapes ? sh.shapes.all() : [];
      const spark = sh.getSparklineGroups ? (sh.getSparklineGroups() || []) : [];
      let images = 0;
      try { images = sh.pictures ? sh.pictures.all().length : 0; } catch (e) {}
      let comments = 0;
      try {
        // v19：批注在 sheet.comments 集合上；旧版 sheet.getComment 已移除
        for (let r = 0; r < Math.min(sh.getRowCount(), 40); r++) {
          for (let c = 0; c < Math.min(sh.getColumnCount(), 12); c++) {
            if (sh.comments && sh.comments.get(r, c)) comments++;
          }
        }
      } catch (e) { comments = -1; }
      out.elements.push({
        i, name: sh.name(), charts: charts.length, shapes: shapes.length,
        images, sparklineGroups: spark.length, comments,
        tables: sh.tables ? sh.tables.all().length : 0,
      });
      const pi = sh.printInfo();
      out.printInfo.push({
        i, name: sh.name(),
        orientation: pi.orientation(),
        paperSize: pi.paperSize && pi.paperSize(),
        printArea: sh.getPrintArea ? String(sh.getPrintArea(0) || '') : '',
        rowStart: pi.rowStart(), rowEnd: pi.rowEnd(), columnStart: pi.columnStart(), columnEnd: pi.columnEnd(),
        repeatRowStart: pi.repeatRowStart(), repeatRowEnd: pi.repeatRowEnd(),
        fitPagesWide: pi.fitPagesWide(), fitPagesTall: pi.fitPagesTall(),
        centering: pi.centering && pi.centering(),
        zoomFactor: pi.zoomFactor && pi.zoomFactor(),
        showGridLine: pi.showGridLine(), showBorder: pi.showBorder(),
        headerCenter: pi.pageHeaderFooter() && pi.pageHeaderFooter().normal && pi.pageHeaderFooter().normal.header
          ? pi.pageHeaderFooter().normal.header.center : null,
        margin: pi.margin(),
      });
    });

    // 逐个 sheet 导出 PDF（跳过试用版 Evaluation 表），base64 回传
    const pdfs = [];
    for (const s of out.sheets) {
      if (/evaluation/i.test(s.name) || !s.visible) continue;
      try {
        const blob = await new Promise((res, rej) => sp.savePDF(res, rej, { title: s.name }, s.i));
        const b64 = await new Promise((res) => {
          const fr = new FileReader();
          fr.onload = () => res(String(fr.result).split(',')[1]);
          fr.readAsDataURL(blob);
        });
        pdfs.push({ i: s.i, name: s.name, size: blob.size, b64 });
      } catch (e) {
        pdfs.push({ i: s.i, name: s.name, error: String(e && e.message || e) });
      }
    }
    return { ...out, pdfs };
  }, SAMPLE);

  const cfg = { sheets: info.sheets, elements: info.elements, printInfo: info.printInfo };
  fs.writeFileSync(path.join(OUT, 'sjs-printinfo.json'), JSON.stringify(cfg, null, 2));

  console.log('=== sheets ===');
  info.sheets.forEach((s) => console.log(s.i, s.name, 'visible=' + s.visible, s.rowCount + 'x' + s.colCount));
  console.log('=== elements ===');
  info.elements.forEach((e) => console.log(JSON.stringify(e)));
  console.log('=== printInfo ===');
  info.printInfo.forEach((p) => console.log(JSON.stringify(p)));
  console.log('=== pdf ===');
  info.pdfs.forEach((p) => {
    if (p.error) { console.log(p.i, p.name, 'ERR', p.error); return; }
    const f = path.join(OUT, 'sjs-' + p.i + '-' + p.name + '.pdf');
    fs.writeFileSync(f, Buffer.from(p.b64, 'base64'));
    console.log(p.i, p.name, p.size, 'bytes ->', f);
  });

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
