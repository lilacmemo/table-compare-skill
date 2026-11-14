// 缺口复核探针：把「方法名里没看到」的那些能力，改用真实调用再验一遍。
// 起因：E5（按颜色筛选）最初只查 FFilter 的方法名就判 FAIL，实际 IFilterColumn 里
//       有 colorFilters 字段；D 组（排序）也只查了 Facade 方法名，而表内排序其实是
//       一条命令 sheet.command.set-table-sort-state。**没找到方法名 ≠ 做不到。**
// 用法：NODE_PATH=../poc-compare/node_modules node tools/probe_table_gaps.js

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8487);
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

// ---------- 1. SJS 按颜色筛选 ----------
const SJS_COLOR = `
sheet.setArray(0, 0, [['区域','值'],['华东',10],['华南',20],['华东',30],['华北',40]]);
sheet.getCell(1, 0).backColor('Cyan');
sheet.getCell(2, 0).backColor('Yellow');
sheet.getCell(3, 0).backColor('Cyan');
var F = GC.Spread.Sheets.Filter, C = GC.Spread.Sheets.ConditionalFormatting;
var rf = new F.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 5, 2));
sheet.rowFilter(rf);
var cond = new C.Condition(C.ConditionType.colorCondition,
  { compareType: C.ColorCompareType.backgroundColor, expected: 'Cyan' });
rf.addFilterItem(0, cond);
rf.filter(0);
var vis = [];
for (var r = 1; r < 5; r++) if (sheet.getRowVisible(r)) vis.push(sheet.getValue(r, 0));
return '按底色 Cyan 筛选后可见=' + vis.join(',') + '（期望 华东,华东）';`;

// ---------- 2. SJS 前 N 项筛选：ConditionType 里有 top10Condition，试出正确参数 ----------
const SJS_TOP = `
var C = GC.Spread.Sheets.ConditionalFormatting, F = GC.Spread.Sheets.Filter;
var condTypes = Object.keys(C.ConditionType).filter(function (k) { return isNaN(+k); });
// 正确参数出自官方 Demo（custom-header）：{ type: Top10ConditionType, expected: N }
// compareType 对这个条件不起作用，传了也不报错、结果仍是「最小的 N 个」——容易被误读成"参数生效"
var shapes = [
  { type: 0, expected: 2 },
  { type: 1, expected: 2 }
];
var out = [];
for (var i = 0; i < shapes.length; i++) {
  sheet.clear(0, 0, 20, 5);
  sheet.setArray(0, 0, [['产品','金额'],['a',7600],['b',8800],['c',9500],['d',10300],['e',12000]]);
  var rf = new F.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 6, 2));
  sheet.rowFilter(rf);
  try {
    var c = new C.Condition(C.ConditionType.top10Condition, shapes[i]);
    rf.addFilterItem(1, c);
    rf.filter(1);
    var vis = [];
    for (var r = 1; r < 6; r++) if (sheet.getRowVisible(r)) vis.push(sheet.getValue(r, 1));
    out.push(JSON.stringify(shapes[i]) + '→可见=' + vis.join(','));
  } catch (e) { out.push(JSON.stringify(shapes[i]) + '→抛错:' + e.message); }
}
return 'ConditionType 含 top10Condition=' + (condTypes.indexOf('top10Condition') >= 0) + ' | ' + out.join(' | ');`;

// ---------- 3. Univer 按颜色筛选 ----------
const UNI_COLOR = `
sheet.getRange('A1:B5').setValues([['区域','值'],['华东',10],['华南',20],['华东',30],['华北',40]]);
sheet.getRange('A2').setBackgroundColor('#00ffff');
sheet.getRange('A3').setBackgroundColor('#ffff00');
sheet.getRange('A4').setBackgroundColor('#00ffff');
sheet.getRange('A1:B5').createFilter();
var f = sheet.getRange('A1:B5').getFilter();
var cands = ['#00ffff', 'rgb(0,255,255)', 'rgb(0, 255, 255)', 'rgba(0, 255, 255, 1)'];
var tried = [];
for (var i = 0; i < cands.length; i++) {
  try {
    f.setColumnFilterCriteria(0, { colId: 0, colorFilters: { cellFillColors: [cands[i]] } });
    tried.push(cands[i] + '→隐藏行=' + JSON.stringify(f.getFilteredOutRows()));
  } catch (e) { tried.push(cands[i] + '→抛错:' + e.message); }
}
// 再对比一次「按值筛选」是否正常工作，用来排除「筛选整体无效」
f.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['华东'] } });
var byValue = JSON.stringify(f.getFilteredOutRows());
return tried.join(' | ') + ' || 按值筛选(华东)隐藏行=' + byValue;`;

// ---------- 4. Univer 表内排序（命令，不是 Facade） ----------
const UNI_SORT = `
sheet.getRange('A1:B6').setValues([['产品','金额'],['a',7600],['b',8800],['c',9500],['d',10300],['e',12000]]);
await sheet.addTable('TblSort', sheet.getRange('A1:B6').getRange(), 'tid-sort', {});
var before = [1,2,3,4,5].map(function (r) { return sheet.getRange(r, 1).getValue(); }).join(',');
var unitId = wb.getId();
var log = [];
try {
  var ok = await univerAPI.executeCommand('sheet.command.set-table-sort-state',
    { unitId: unitId, tableId: 'tid-sort', sortInfo: { columnIndex: 1, sortState: 'asc' } });
  log.push('executeCommand asc=' + ok);
} catch (e) { log.push('executeCommand 抛错:' + e.message); }
await new Promise(function (r) { setTimeout(r, 500); });
var after = [1,2,3,4,5].map(function (r) { return sheet.getRange(r, 1).getValue(); }).join(',');
var info = sheet.getSubTableInfos()[0] || {};
return '排序前=' + before + ' 排序后=' + after + ' | ' + log.join(' | ')
  + ' | subTable 里有没有 sortInfo=' + (JSON.stringify(info).match(/sort/i) ? '有' : '无');`;

// ---------- 5. Univer 文本长度验证有没有别的入口 ----------
const UNI_TEXTLEN = `
var dv = univerAPI.newDataValidation();
var proto = [];
var o = dv;
while (o && o !== Object.prototype) {
  Object.getOwnPropertyNames(o).forEach(function (n) { if (n !== 'constructor') proto.push(n); });
  o = Object.getPrototypeOf(o);
}
var uniq = proto.filter(function (v, i, a) { return a.indexOf(v) === i; });
return 'builder 全部方法=' + uniq.sort().join(',') + ' | 含 text/len 的='
  + (uniq.filter(function (n) { return /text|len/i.test(n); }).join(',') || '无');`;


// ---------- 6. Univer 汇总行：showFooter / hasTotalRow 到底哪个生效 ----------
// A6 已经证明 showHeader 的 @todo 注释是过时的（传 false 真的生效），
// 所以汇总行不能只试 hasTotalRow 一个字段。
const UNI_FOOTER = `
var variants = [
  ['不传', {}],
  ['hasTotalRow', { hasTotalRow: true }],
  ['showFooter', { showFooter: true }],
  ['两者都传', { showFooter: true, hasTotalRow: true }]
];
var out = [];
for (var i = 0; i < variants.length; i++) {
  var top = i * 8;
  var r1 = top + 1, r2 = top + 5;
  sheet.getRange(r1, 0, 5, 2).setValues([['区域','金额'],['华东',5200],['华南',4100],['华北',3800],['华东',100]]);
  var tid = 'tid-foot-' + i;
  await sheet.addTable('TblF' + i, sheet.getRange(r1, 0, 5, 2).getRange(), tid, variants[i][1]);
  await new Promise(function (r) { setTimeout(r, 200); });
  var info = sheet.getSubTableInfos().filter(function (x) { return x.id === tid; })[0] || {};
  out.push(variants[i][0] + '→range=' + JSON.stringify(info.range)
    + ' 字段=' + Object.keys(info).join('/'));
}
return out.join('  ||  ');`;

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: process.env.PW_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => console.log('PAGEERR', String(e.message).slice(0, 160)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__cmpRun && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1200);

  const jobs = [
    ['SJS  按颜色筛选', 'sjs', SJS_COLOR],
    ['SJS  Top N 入口', 'sjs', SJS_TOP],
    ['UNI  按颜色筛选', 'uni', UNI_COLOR],
    ['UNI  表内排序命令', 'uni', UNI_SORT],
    ['UNI  文本长度入口', 'uni', UNI_TEXTLEN],
    ['UNI  汇总行字段', 'uni', UNI_FOOTER],
  ];
  const out = [];
  for (const [label, side, code] of jobs) {
    await page.evaluate(() => window.__cmpClear('probe'));
    await page.waitForTimeout(300);
    const r = await page.evaluate(({ s, c }) => window.__cmpRun(s, c), { s: side, c: code });
    out.push({ label, side, ...r });
    console.log('\n### ' + label + '\n  ' + (r.ok ? r.value : 'FAIL: ' + r.err));
  }
  fs.writeFileSync(path.join(OUT, 'table-gaps.json'), JSON.stringify({ at: new Date().toISOString(), results: out }, null, 2));
  await browser.close();
  srv.close();
})();
