// 表模块「一眼看得见的差异」证据图：两侧跑同一套演示，各截一张，再拼左右。
//   左侧 SpreadJS：表样式(medium2) + 镶边行 + 首列高亮 + 筛选按钮 + 计算列 + 汇总行
//   右侧 Univer  ：同样的 addTable + 样式 ID + 计算列，但无镶边/首列/汇总行入口
// 用法：NODE_PATH=../poc-compare/node_modules node tools/shoot_table_showcase.js
// 输出：tools/out/show-sjs.png、tools/out/show-uni.png（拼图由 PIL 另做）

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8486);

const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json', '.png': 'image/png',
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

const DATA = `[['区域','产品','数量','单价','金额'],
 ['华东','苹果',10,5,null],['华北','香蕉',4,30,null],
 ['华南','橙子',7,15,null],['华东','葡萄',2,50,null]]`;

const SJS = `
sheet.setArray(0, 0, ${DATA});
var t = sheet.tables.add('TblShow', 0, 0, 5, 5, GC.Spread.Sheets.Tables.TableThemes.medium2);
t.bandRows(true);
t.highlightFirstColumn(true);
t.filterButtonVisible(true);
for (var r = 1; r <= 4; r++) sheet.setFormula(r, 4, '=[@数量]*[@单价]');
t.showFooter(true);
t.setColumnValue(0, '总计');
t.setColumnFormula(4, '=SUBTOTAL(109,[金额])');
[80, 80, 70, 70, 90].forEach(function (w, i) { sheet.setColumnWidth(i, w); });
spread.getActiveSheet().setRowHeight(0, 28);
return '汇总=' + sheet.getValue(5, 4) + ' 表数=' + sheet.tables.all().length;`;

const UNI = `
sheet.getRange('A1:E5').setValues(${DATA});
await sheet.addTable('TblShow', sheet.getRange('A1:E5').getRange(), 'tid-show',
  { tableStyleId: 'table-default-2', showAutoFilter: true });
var info = sheet.getSubTableInfos()[0] || {};
var cols = info.subTableColumns || info.columns || [];
var colId = cols[4] && (cols[4].id || cols[4].columnId);
await sheet.setTableColumnFormula('tid-show', colId, '=[@数量]*[@单价]');
for (var i = 0; i < 5; i++) { try { sheet.setColumnWidth(i, [80, 80, 70, 70, 90][i]); } catch (e) {} }
await new Promise(function (r) { setTimeout(r, 500); });
return '金额列=' + [1, 2, 3, 4].map(function (r) { return sheet.getRange(r, 4).getValue(); }).join(',')
  + ' 表字段=[' + Object.keys(info).join(',') + ']';`;

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 900 }, deviceScaleFactor: 2 });
  page.on('pageerror', (e) => console.log('PAGEERR', String(e.message).slice(0, 160)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__cmpRun && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1500);
  await page.evaluate(() => window.__cmpClear('证据图'));

  const rs = await page.evaluate((c) => window.__cmpRun('sjs', c), SJS);
  console.log('SJS :', JSON.stringify(rs).slice(0, 300));
  const ru = await page.evaluate((c) => window.__cmpRun('uni', c), UNI);
  console.log('UNI :', JSON.stringify(ru).slice(0, 300));

  await page.waitForTimeout(1200);
  for (const [side, sel] of [['sjs', '#sjsPane'], ['uni', '#sjsPane ~ section.pane']]) {
    const el = await page.$(sel);
    if (!el) { console.log('未找到面板', sel); continue; }
    const f = path.join(OUT, `show-${side}.png`);
    await el.screenshot({ path: f });
    console.log('已截图', f);
  }
  await browser.close();
  srv.close();
})();
