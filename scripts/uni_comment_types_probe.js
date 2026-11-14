// 批注格式覆盖对照：Excel 新式讨论（threadedComment, A1）vs 传统批注（comments1, A2）
//   样本：samples/批注-新式与传统对照.xlsx（Excel 365 COM 造：A1 AddCommentThreaded、A2 AddComment）
//   目的：确认 Univer 导入到底是「不认任何批注」还是「只认新式讨论、不认传统批注」
//   同时跑 SpreadJS 做基准对照（SJS 两种都读）
// 用法：node tools/serve.js 8124（另开）→ NODE_PATH=../poc-compare/node_modules node tools/probe_uni_comment_types.js
// 产物：out/data/comment-types.json
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DATA = path.join(ROOT, 'out', 'data');
const SAMPLE = 'samples/批注-新式与传统对照.xlsx';
const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';

(async () => {
  fs.mkdirSync(DATA, { recursive: true });
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1400, height: 900 } });
  const p = await ctx.newPage();
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && window.APP.ready && window.APP.ready.sjs && window.APP.uni, null, { timeout: 120000 });

  const uni = await p.evaluate(async (f) => {
    await APP.importOne('uni', f, 'c.xlsx');
    const out = [];
    for (const sh of window.univerAPI.getActiveWorkbook().getSheets()) {
      let list = [];
      try { list = sh.getComments() || []; } catch (e) { }
      out.push({
        sheet: sh.getSheetName(),
        n: list.length,
        items: list.map((c) => {
          let a1 = null, t = null;
          try { a1 = c.getRange() && c.getRange().getA1Notation(); } catch (e) { }
          try { const d = c.getCommentData(); t = JSON.stringify(d && d.text ? d.text : d).slice(0, 120); } catch (e) { }
          return { a1, t };
        }),
      });
    }
    return out;
  }, SAMPLE);

  const sjs = await p.evaluate(async (f) => {
    await APP.importOne('sjs', f, 'c.xlsx');
    const sp = APP.sjs;
    const out = [];
    for (let i = 0; i < sp.getSheetCount(); i++) {
      const sh = sp.getSheet(i);
      let notes = [], threads = [];
      try { notes = sh.comments.all() || []; } catch (e) { }
      try { threads = (sh.threadedComments && sh.threadedComments.all()) || []; } catch (e) { }
      out.push({
        sheet: sh.name(),
        notes: notes.length,
        noteTexts: notes.map((c) => (() => { try { return String(c.text()).slice(0, 40); } catch (e) { return null; } })()),
        threads: threads.length,
        threadTexts: threads.map((t) => {
          try {
            const r = t.all()[0];
            return (r && r.message || []).map((m) => m.value || m.text || '').join('').slice(0, 40);
          } catch (e) { return null; }
        }),
      });
    }
    return out;
  }, SAMPLE);

  const result = { at: new Date().toISOString(), sample: SAMPLE, uni, sjs };
  fs.writeFileSync(path.join(DATA, 'comment-types.json'), JSON.stringify(result, null, 2));
  console.log(JSON.stringify(result, null, 2));
  await b.close();
})();
