// 打印 / PDF API 字段级矩阵：逐字段「设置 → 读回 / 观测」判 PASS/FAIL
//   NODE_PATH=../poc-compare/node_modules node tools/print_api_matrix.js
// 判定口径（用户定）：完全设置成功 = PASS；抛错、无入口、设置后读回不一致 = FAIL
//   SpreadJS：printInfo 是 getter/setter 二合一，set 后同名无参调用读回
//   Univer  ：没有 getter（getPrintConfig 不存在），只能判"设置是否被接受"，
//             生效与否另用打印预览观测（见 tools/univer_print_capture.js）
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out', 'print');
const PORT = Number(process.env.PORT || 8490);
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
  await page.waitForFunction(() => window.univerAPI && window.__wb && window.__spread, null, { timeout: 90000 });
  await page.waitForTimeout(1000);

  const result = await page.evaluate(async () => {
    const sp = window.__spread;
    const u = window.univerAPI;
    const wait = (ms) => new Promise((r) => setTimeout(r, ms));
    const safe = (v) => {
      try { return typeof v === 'object' ? JSON.stringify(v) : String(v); } catch (e) { return String(v); }
    };

    // ---------- SpreadJS：一表一字段，避免相互污染 ----------
    const sjsCases = [
      ['打印区域 起始行 rowStart', 'rowStart', 2],
      ['打印区域 结束行 rowEnd', 'rowEnd', 20],
      ['打印区域 起始列 columnStart', 'columnStart', 1],
      ['打印区域 结束列 columnEnd', 'columnEnd', 5],
      ['重复标题行 起 repeatRowStart', 'repeatRowStart', 0],
      ['重复标题行 止 repeatRowEnd', 'repeatRowEnd', 1],
      ['重复标题列 起 repeatColumnStart', 'repeatColumnStart', 0],
      ['重复标题列 止 repeatColumnEnd', 'repeatColumnEnd', 1],
      ['显示网格线 showGridLine', 'showGridLine', false],
      ['显示边框 showBorder', 'showBorder', false],
      ['黑白打印 blackAndWhite', 'blackAndWhite', true],
      ['页面方向 orientation', 'orientation', 2],
      ['纸张大小 paperSize', 'paperSize', { width: 842, height: 1191, kind: 8 }],
      ['页序 pageOrder', 'pageOrder', 1],
      ['指定页码范围 pageRange', 'pageRange', '1-2'],
      ['起始页码 firstPageNumber', 'firstPageNumber', 5],
      ['缩放 适配页宽 fitPagesWide', 'fitPagesWide', 2],
      ['缩放 适配页高 fitPagesTall', 'fitPagesTall', 3],
      ['缩放百分比 zoomFactor', 'zoomFactor', 0.8],
      ['首页不同 differentFirstPage', 'differentFirstPage', true],
      ['奇偶页不同 differentOddAndEvenPages', 'differentOddAndEvenPages', true],
      ['页脚左 footerLeft', 'footerLeft', 'FL'],
      ['页脚中 footerCenter', 'footerCenter', 'FC'],
      ['页脚右 footerRight', 'footerRight', 'FR'],
      ['页眉左 headerLeft', 'headerLeft', 'HL'],
      ['页眉中 headerCenter', 'headerCenter', 'HC'],
      ['页眉右 headerRight', 'headerRight', 'HR'],
      ['页边距 margin', 'margin', { top: 100, bottom: 100, left: 40, right: 40, header: 20, footer: 20 }],
      ['居中 centering', 'centering', { horizontalCentered: true, verticalCentered: true }],
      ['最佳列宽 bestFitColumns', 'bestFitColumns', true],
      ['最佳行高 bestFitRows', 'bestFitRows', true],
      ['打印质量 qualityFactor', 'qualityFactor', 2],
    ];

    const sjs = [];
    for (const [label, method, value] of sjsCases) {
      sp.addSheet('tmp_' + method);
      const sh = sp.getSheet(sp.getSheetCount() - 1);
      sh.setArray(0, 0, [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]]);
      const pi = sh.printInfo();
      let status = 'PASS', back = null, err = '';
      try {
        if (typeof pi[method] !== 'function') throw new Error('方法不存在');
        pi[method](value);
        back = pi[method]();
      } catch (e) { status = 'FAIL'; err = e && e.message; }

      // 读回比较：数字容差、对象逐键、字符串全等
      const same = (a, b) => {
        if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 0.001;
        if (a && b && typeof a === 'object' && typeof b === 'object') {
          // 只校验我们显式设置的键：读回对象可能多带内部字段（如 width/height/kind + 原型）
          return Object.keys(a).every((k) => same(a[k], b[k]));
        }
        return a === b;
      };
      let match = null;
      if (status === 'PASS') {
        // 双通道：结构比较 + 序列化比较（读回对象可能是 Number 包装等特殊形态）
        match = same(value, back) ||
          (typeof value === 'object' && JSON.stringify(value) === JSON.stringify(back));
        if (!match) status = 'FAIL';
      }
      sjs.push({ label, method, set: safe(value), got: safe(back), status, err });
      sp.removeSheet(sp.getSheetCount() - 1);
    }

    // 带图片/水印的项单独测（需要图片数据）
    sp.addSheet('tmp_extra');
    const shx = sp.getSheet(sp.getSheetCount() - 1);
    const pix = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const extras = [];
    try {
      const pi = shx.printInfo();
      pi.headerLeftImage(pix);
      extras.push({ label: '页眉图片 headerLeftImage', set: '1px png', got: String(pi.headerLeftImage()).slice(0, 32), status: 'PASS' });
    } catch (e) { extras.push({ label: '页眉图片 headerLeftImage', status: 'FAIL', err: e.message }); }
    try {
      const pi = shx.printInfo();
      pi.watermark({ imageSrc: pix, x: 10, y: 10, width: 100, height: 40, page: 'all' });
      extras.push({ label: '水印 watermark', set: 'image+pos', got: 'no getter', status: 'PASS' });
    } catch (e) { extras.push({ label: '水印 watermark', status: 'FAIL', err: e.message }); }
    try {
      const pi = shx.printInfo();
      pi.pageHeaderFooter({ normal: { header: { left: 'A', center: 'B', right: 'C' } } });
      const b = pi.pageHeaderFooter();
      extras.push({
        label: '页眉页脚结构 pageHeaderFooter', set: 'left/center/right',
        got: b && b.normal && b.normal.header ? b.normal.header.center : 'n/a',
        status: b && b.normal && b.normal.header && b.normal.header.center === 'B' ? 'PASS' : 'FAIL',
      });
    } catch (e) { extras.push({ label: '页眉页脚结构 pageHeaderFooter', status: 'FAIL', err: e.message }); }
    sp.removeSheet(sp.getSheetCount() - 1);

    // ---------- Univer：无 getter，只能判"设置是否被接受" ----------
    const wb = window.__wb;
    const wsu = wb.getActiveSheet();
    wsu.getRange('A1').setValue('matrix');
    const E = u.Enum;
    const uniCases = [
      ['打印区域 area', { area: E.PrintArea.AllSelection }],
      ['纸张 paperSize', { paperSize: E.PrintPaperSize.A3 }],
      ['方向 direction', { direction: E.PrintDirection.Portrait }],
      ['缩放 scale', { scale: E.PrintScale.FitPage }],
      ['自定义缩放 customScale', { scale: E.PrintScale.Custom, customScale: 80 }],
      ['打印重复行 freeze', { freeze: [E.PrintFreeze.Row, E.PrintFreeze.Column] }],
      ['页边距 margin', { margin: E.PrintPaperMargin.Narrow }],
      ['自定义页边距 marginCustom', { margin: E.PrintPaperMargin.Custom, marginCustom: { top: 10, right: 10, bottom: 10, left: 10 } }],
      ['每页最多行 maxRowsEachPage', { maxRowsEachPage: 10 }],
      ['每页最多列 maxColumnsEachPage', { maxColumnsEachPage: 5 }],
    ];
    const uni = [];
    for (const [label, cfg] of uniCases) {
      try {
        wb.updatePrintConfig(cfg);
        uni.push({ label, cfg: JSON.stringify(cfg), status: 'PASS' });
      } catch (e) { uni.push({ label, cfg: JSON.stringify(cfg), status: 'FAIL', err: e.message }); }
      await wait(60);
    }

    const uniRender = [
      ['网格线 gridlines', { gridlines: false }],
      ['水平对齐 hAlign', { hAlign: E.PrintAlign.End }],
      ['垂直对齐 vAlign', { vAlign: E.PrintAlign.Middle }],
      ['页眉页脚枚举 headerFooter', { headerFooter: [E.PrintHeaderFooter.PageSize, E.PrintHeaderFooter.DateA] }],
      ['自定义页眉页脚 headerFooterSetting', { isCustomHeaderFooter: true, headerFooterSetting: { topLeft: 'L', topCenter: 'C', topRight: 'R' } }],
      ['水印 watermark', { watermark: 'CONFIDENTIAL' }],
    ];
    const uni2 = [];
    for (const [label, cfg] of uniRender) {
      try {
        wb.updatePrintRenderConfig(cfg);
        uni2.push({ label, cfg: JSON.stringify(cfg), status: 'PASS' });
      } catch (e) { uni2.push({ label, cfg: JSON.stringify(cfg), status: 'FAIL', err: e.message }); }
      await wait(60);
    }

    // 有无 getter
    const getters = ['getPrintConfig', 'getPrintRenderConfig'].map((k) => [k, typeof wb[k]]);

    return { sjs, extras, uni, uni2, getters };
  });

  const safeLog = (o) => o;
  fs.writeFileSync(path.join(OUT, 'print-api-matrix.json'), JSON.stringify(result, null, 2));

  const fmt = (rows, kind) => {
    rows.forEach((r) => {
      console.log([r.status.padEnd(4), kind, r.label, '| set=' + (r.set || r.cfg || ''), '| got=' + (r.got === undefined ? '' : r.got), r.err ? '| ERR ' + r.err : ''].join(' '));
    });
  };
  console.log('=== SpreadJS PrintInfo ===');
  fmt(result.sjs, 'SJS');
  fmt(result.extras, 'SJS');
  console.log('=== Univer updatePrintConfig ===');
  fmt(result.uni, 'UNI');
  console.log('=== Univer updatePrintRenderConfig ===');
  fmt(result.uni2, 'UNI');
  console.log('=== getter ===', JSON.stringify(result.getters));

  const count = (rows) => rows.filter((r) => r.status === 'PASS').length + '/' + rows.length;
  console.log('SJS 通过', count(result.sjs.concat(result.extras)), '| Univer 通过', count(result.uni.concat(result.uni2)));

  await browser.close();
  srv.close();
})().catch((e) => { console.error('FAILED', e); process.exit(1); });
