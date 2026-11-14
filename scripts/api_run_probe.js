// API 探针跑批器
//   node tools/api_run_probe.js surface   → tools/out/api-surface.json   （两家 API 面枚举）
//   node tools/api_run_probe.js items     → tools/out/api-items.json     （A~K 135 项 set→read 实测）
// 依赖：playwright（NODE_PATH 指向装着它的 node_modules）
// 浏览器：PW_CHANNEL 覆盖，默认 msedge（同 Chromium）；装了 Chrome 可 PW_CHANNEL=chrome。
//
// ⚠️ 放置位置：本脚本假定所在目录是 <项目>/tools/，其父目录含 index.html（即双栏对照台根）。
//    从 skill 的 scripts/ 里拿出来用时，请复制到工程 tools/ 下再跑，否则 ROOT 指错找不到页面。
//    端口可用 PORT 环境变量改（默认 8461，被占用时换一个）。

const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(__dirname, 'out');
const PORT = Number(process.env.PORT || 8461);
const CHANNEL = process.env.PW_CHANNEL || 'msedge';
const MODE = process.argv[2] || 'surface';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const url = decodeURIComponent(req.url.split('?')[0]);
      let p = path.join(ROOT, url === '/' ? 'index.html' : url);
      if (!p.startsWith(ROOT)) { res.writeHead(403).end(); return; }
      fs.readFile(p, (err, buf) => {
        if (err) { res.writeHead(404).end('not found'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(buf);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

// 原型链枚举：收集自有属性里的方法与访问器（必须定义在页面函数内部——page.evaluate 只序列化该函数）
function surfaceFn() {
  const COLLECT = function (obj, maxDepth) {
    const methods = new Set();
    const accessors = new Set();
    const values = {};
    let cur = obj, d = 0;
    while (cur && cur !== Object.prototype && d < (maxDepth || 8)) {
      for (const k of Object.getOwnPropertyNames(cur)) {
        if (k === 'constructor') continue;
        let desc;
        try { desc = Object.getOwnPropertyDescriptor(cur, k); } catch (e) { continue; }
        if (!desc) continue;
        if (typeof desc.value === 'function') methods.add(k);
        else if (desc.get || desc.set) {
          accessors.add(k);
          if (desc.get && !(k in values)) {
            try {
              const v = desc.get.call(obj);
              const t = typeof v;
              if (t !== 'function' && t !== 'object') values[k] = v;
              else if (v === null) values[k] = null;
            } catch (e) { /* 有些 getter 需要参数，忽略 */ }
          }
        }
      }
      cur = Object.getPrototypeOf(cur); d++;
    }
    return { methods: [...methods].sort(), accessors: [...accessors].sort(), values };
  };

  const out = { at: new Date().toISOString(), sjs: {}, uni: {}, enums: {} };

  // ---------- SpreadJS ----------
  try {
    const GCs = window.GC.Spread.Sheets;
    const sheet = window.__spread.getActiveSheet();
    out.sjs.style = COLLECT(new GCs.Style(), 6);
    out.sjs.sheet = COLLECT(sheet, 8);
    out.sjs.range = COLLECT(sheet.getRange(0, 0, 1, 1), 8);
    out.sjs.conditionalFormats = COLLECT(sheet.conditionalFormats, 8);
    out.sjs.namespaces = Object.keys(GCs).sort();

    const enumDump = (e) => {
      const o = {};
      for (const k of Object.keys(e)) if (typeof e[k] === 'number') o[k] = e[k];
      return o;
    };
    for (const nm of ['HorizontalAlign', 'VerticalAlign', 'LineStyle', 'PatternType',
      'TextDecorationType', 'TextDirection', 'CalculationType', 'ClipboardPasteType']) {
      if (!GCs[nm]) continue;
      const e = {};
      for (const k of Object.keys(GCs[nm])) if (typeof GCs[nm][k] === 'number') e[k] = GCs[nm][k];
      out.enums['sjs.' + nm] = e;
    }
    const CF = GCs.ConditionalFormatting;
    if (CF) {
      out.sjs.cfNamespace = Object.keys(CF).sort();
      const pick = ['ScaleValueType', 'IconSetType', 'IconCriterionOperator', 'DataBarAxisPosition',
        'DataBarDirection', 'AverageConditionType', 'Top10ConditionType', 'TextComparisonOperator',
        'DateOccurringType', 'CellValueType', 'RuleType'];
      for (const k of pick) if (CF[k]) {
        const e = {};
        for (const kk of Object.keys(CF[k])) if (typeof CF[k][kk] === 'number') e[k] = e[kk];
        out.enums['sjs.CF.' + k] = e;
      }
      const rules = {};
      for (const k of Object.keys(CF)) {
        const v = CF[k];
        if (typeof v === 'function' && /Rule$/.test(k) && v.prototype) {
          rules[k] = COLLECT(v.prototype, 4).methods.concat(COLLECT(v.prototype, 4).accessors).sort();
        }
      }
      out.sjs.cfRuleClasses = rules;
    }
  } catch (e) { out.sjs.error = String(e && e.message || e); }

  // ---------- Univer ----------
  try {
    const api = window.univerAPI;
    const wb = window.__wb;
    const sheet = wb.getActiveSheet();
    const range = sheet.getRange(0, 0, 1, 1);
    out.uni.facade = COLLECT(api, 4);
    out.uni.workbook = COLLECT(wb, 5);
    out.uni.sheet = COLLECT(sheet, 6);
    out.uni.range = COLLECT(range, 6);

    // 条件格式 builder：链式对象，逐层收集
    if (typeof sheet.newConditionalFormattingRule === 'function') {
      const b = sheet.newConditionalFormattingRule();
      out.uni.cfBuilder = COLLECT(b, 8);
      // setDataBar / setColorScale / setIconSet 等的参数结构，看返回对象的键
      try {
        const probe = {};
        const r1 = b.setDataBar ? b.setDataBar({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6' }) : null;
        probe.dataBarChain = r1 ? COLLECT(r1, 8).methods.length : null;
        out.uni.cfBuilderChainSample = probe;
      } catch (e) { out.uni.cfBuilderChainErr = String(e && e.message || e); }
    }
    if (typeof sheet.getConditionalFormattingRules === 'function') {
      out.uni.hasCFFacade = true;
    }
    // 内部模型是否暴露（判断 S4 用）
    out.uni.hasGetSnapshot = typeof wb.getSnapshot === 'function';
    out.uni.sheetId = sheet.getSheetId ? sheet.getSheetId() : null;
  } catch (e) { out.uni.error = String(e && e.message || e); }

  return out;
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const srv = await serve();
  const { chromium } = require('playwright');
  const browser = await chromium.launch({ channel: CHANNEL, headless: true });
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e.message)));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });

  // 等两家引擎就绪
  await page.waitForFunction(() => window.__spread && window.__wb && window.univerAPI, null, { timeout: 60000 });
  await page.waitForTimeout(1200);

  let result;
  if (MODE === 'surface') {
    result = await page.evaluate(surfaceFn);
  } else {
    result = await page.evaluate(async () => await window.ApiProbe.runAll());
  }

  const file = path.join(OUT, MODE === 'surface' ? 'api-surface.json' : 'api-items.json');
  fs.writeFileSync(file, JSON.stringify(result, null, 2), 'utf8');
  console.log('写出', file);
  if (MODE === 'surface') {
    console.log('SJS style 方法数', result.sjs.style && result.sjs.style.methods.length);
    console.log('SJS sheet 方法数', result.sjs.sheet && result.sjs.sheet.methods.length);
    console.log('SJS CF 方法数', result.sjs.conditionalFormats && result.sjs.conditionalFormats.methods.length);
    console.log('Uni range 方法数', result.uni.range && result.uni.range.methods.length);
    console.log('Uni cfBuilder 方法数', result.uni.cfBuilder && result.uni.cfBuilder.methods.length);
    if (result.sjs.error) console.log('SJS error:', result.sjs.error);
    if (result.uni.error) console.log('Uni error:', result.uni.error);
  } else {
    const items = result.items || [];
    console.log('用例数', items.length);
    const fail = items.filter((i) => i.error);
    if (fail.length) console.log('异常项', fail.length, fail.slice(0, 5).map((f) => f.id + ':' + f.error));
  }
  if (errs.length) console.log('页面错误：', errs.slice(0, 5));

  await browser.close();
  srv.close();
})().catch((e) => { console.error(e); process.exit(1); });
