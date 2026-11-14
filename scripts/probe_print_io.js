// 打印设置导入保真度对账：同一份 xlsx，两侧读回实际生效的页面设置
//   SJS 侧：sheet.printInfo() 逐字段读（margin() 是对象 getter，不是 marginTop()）
//   Univer 侧：univerAPI.addEvent(univerAPI.Event.BeforeSheetPrintConfirm, cb) —— Event.XXX 是事件名常量，
//     注册要用 addEvent；载荷里有 layoutConfig/renderConfig，取完 e.cancel = true 拦掉真打印。
// 用法：NODE_PATH=<poc-compare/node_modules> node tools/probe_print_io.js
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');
const SAMPLE = 'samples/打印-PDF-多元素样本.xlsx';

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 2600, height: 1500 } });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => console.log('PAGEERROR', e.message.slice(0, 140)));
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });
  await p.evaluate(async (f) => APP.loadSample(f, '打印-PDF-多元素样本.xlsx'), SAMPLE);
  await p.waitForTimeout(1200);

  const sjs = await p.evaluate(() => {
    const isEval = (n) => /evaluation/i.test(String(n || ''));
    const sp = APP.sjs;
    let i = 0;
    for (let k = 0; k < sp.getSheetCount(); k++) if (!isEval(sp.getSheet(k).name())) { i = k; break; }
    sp.setActiveSheetIndex(i);
    const sh = sp.getSheet(i);
    const pi = sh.printInfo();
    const ps = pi.paperSize(); const m = pi.margin(); const ce = pi.centering();
    return {
      sheet: sh.name(),
      paperKind: ps && ps.kind, paperW: Math.round(ps.width), paperH: Math.round(ps.height),
      orientation: pi.orientation(), // 1=portrait 2=landscape
      fitPagesWide: pi.fitPagesWide(), fitPagesTall: pi.fitPagesTall(),
      margin: m,
      header: [pi.headerLeft(), pi.headerCenter(), pi.headerRight()],
      footer: [pi.footerLeft(), pi.footerCenter(), pi.footerRight()],
      centering: ce,
      showGridLine: pi.showGridLine(),
      showColumnHeader: pi.showColumnHeader(), showRowHeader: pi.showRowHeader(),
    };
  });
  console.log('SJS', JSON.stringify(sjs, null, 1));

  const uni = await p.evaluate(async () => {
    const U = window.univerAPI;
    const out = { evtType: typeof U.Event.BeforeSheetPrintConfirm, addEvent: typeof U.addEvent };
    let got = null;
    let sub = null;
    try {
      sub = U.addEvent(U.Event.BeforeSheetPrintConfirm, (e) => {
        try { got = JSON.parse(JSON.stringify({ layoutConfig: e.layoutConfig, renderConfig: e.renderConfig })); } catch (x) { got = { err: x.message }; }
        e.cancel = true;
      });
    } catch (e) { out.regErr = e.message.slice(0, 140); }
    try { await U.getActiveWorkbook().print(); } catch (e) { out.printErr = e.message.slice(0, 140); }
    await new Promise((r) => setTimeout(r, 1000));
    try { sub && sub.dispose && sub.dispose(); } catch (e) {}
    out.payload = got;
    return out;
  });
  console.log('UNI', JSON.stringify(uni, null, 1));

  fs.writeFileSync(path.join(ROOT, 'out', 'data', 'print-io-settings.json'), JSON.stringify({ sjs, uni }, null, 1));
  await b.close();
})();
