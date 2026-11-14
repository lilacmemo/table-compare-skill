// Univer 批注（thread comment）能力与导入还原专项
//   背景：bundle 早前没有注册批注插件，报告里「批注导入 0」可能是"没装插件"造成的假阴性。
//   分三步，逐层定位：
//     1) 导入层：加载样本，遍历每张表的 getComments()，看 Excel 传统批注（comments1.xml）是否被解析
//     2) 能力层：在空白单元格 addCommentAsync → getComment() 读回，确认 API 通（插件装对了）
//     3) 落盘层：导出 xlsx，另由 python 解包核对 XML
// 用法：node tools/serve.js 8124（另开）→ NODE_PATH=../poc-compare/node_modules node tools/probe_uni_comment.js
// 产物：out/data/uni-comment.json、out/data/uni-comment-export.xlsx
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DATA = path.join(ROOT, 'out', 'data');
const SAMPLE = 'samples/打印-PDF-多元素样本.xlsx';
const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';

(async () => {
  fs.mkdirSync(DATA, { recursive: true });
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1600, height: 1000 } });
  const p = await ctx.newPage();
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && window.APP.uni && window.univerAPI, null, { timeout: 120000 });
  const plugins = await p.evaluate(() => window.__proExtraPlugins || {});

  // ---- 步骤 1：导入样本后清点批注
  const load = await p.evaluate(async (f) => {
    try { await APP.loadSample(f, '打印-PDF-多元素样本.xlsx'); return 'ok'; } catch (e) { return 'ERR ' + e.message; }
  }, SAMPLE);

  const afterImport = await p.evaluate(() => {
    const out = { sheets: [], total: 0 };
    try {
      const wb = window.univerAPI.getActiveWorkbook();
      for (const sh of wb.getSheets()) {
        let list = [];
        try { list = sh.getComments() || []; } catch (e) { }
        out.sheets.push({
          name: sh.getSheetName(),
          comments: list.length,
          items: list.map((c) => {
            let a1 = null, txt = null;
            try { const r = c.getRange(); a1 = r && r.getA1Notation(); } catch (e) { }
            try { txt = JSON.stringify(c.getCommentData()).slice(0, 160); } catch (e) { }
            return { a1, txt };
          }),
        });
        out.total += list.length;
      }
    } catch (e) { out.error = String(e && e.message || e).slice(0, 300); }
    return out;
  });

  // ---- 步骤 2：能力探针（导入后的工作簿上，选一张表的空白格）
  const capability = await p.evaluate(async () => {
    const out = { apiShape: {}, add: null, readBack: null };
    try {
      const sh = window.univerAPI.getActiveWorkbook().getSheets()[0];
      out.apiShape.sheetGetComments = typeof sh.getComments;
      out.apiShape.rangeAddCommentAsync = typeof sh.getRange('K60').addCommentAsync;
      out.apiShape.rangeGetComment = typeof sh.getRange('K60').getComment;
      const ok = await sh.getRange('K60').addCommentAsync('探针批注-能力验证');
      out.add = ok === true ? 'ok' : String(ok);
      const c = sh.getRange('K60').getComment();
      out.readBack = c ? String(JSON.stringify(c.getCommentData()).slice(0, 200)) : null;
      // 复读工作表级集合，确认新增能进列表
      out.afterAddCount = sh.getComments().length;
    } catch (e) {
      out.error = String(e && e.message || e).slice(0, 300);
    }
    return out;
  });

  // ---- 步骤 3：导出（批次里已含探针批注），交给 python 解包
  let exportBytes = 0;
  try {
    const b64 = await p.evaluate(async () => (APP.exportUni ? await APP.exportUni() : null));
    if (b64) {
      const buf = Buffer.from(String(b64).replace(/^data:[^,]*,/, ''), 'base64');
      fs.writeFileSync(path.join(DATA, 'uni-comment-export.xlsx'), buf);
      exportBytes = buf.length;
    }
  } catch (e) { console.log('导出失败', e.message); }

  // ---- 步骤 4：把刚导出的文件（含 Univer 自己写的 threadedComment + 兼容 comments1）重新导入
  //     目的：区分「Univer 不认任何批注」和「Univer 只认新式 threaded comment、不认 Excel 传统批注（note）」
  let roundtrip = null;
  try {
    roundtrip = await p.evaluate(async () => {
      const out = { sheets: [], total: 0 };
      await APP.importOne('uni', 'out/data/uni-comment-export.xlsx', 'uni-comment-export.xlsx');
      const wb = window.univerAPI.getActiveWorkbook();
      for (const sh of wb.getSheets()) {
        let list = [];
        try { list = sh.getComments() || []; } catch (e) { }
        out.sheets.push({ name: sh.getSheetName(), comments: list.length, items: list.map((c) => { let a1 = null; try { a1 = c.getRange() && c.getRange().getA1Notation(); } catch (e) { } return a1; }) });
        out.total += list.length;
      }
      return out;
    });
  } catch (e) { roundtrip = { error: String(e.message).slice(0, 200) }; }

  const result = { at: new Date().toISOString(), plugins, load, afterImport, capability, exportBytes, roundtrip };
  fs.writeFileSync(path.join(DATA, 'uni-comment.json'), JSON.stringify(result, null, 2));
  console.log(JSON.stringify(result, null, 2));
  await b.close();
})();
