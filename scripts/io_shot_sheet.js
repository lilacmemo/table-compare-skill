// 切到指定工作表并两侧同框截图（用于"某个不一致点"的局部取证）
//   用法：node tools/shot_sheet.js <样本文件名> <表名关键字> <输出名>
//   例：  node tools/shot_sheet.js 格式与条件格式-全量用例-A至K.xlsx "G 条件格式" cf-format
// 为什么不用 capture_io.js：它只截首表，而格式/条件格式的差异集中在 G/H 表。
const { chromium } = require('playwright');
const path = require('path');

const [, , sample, sheetKey, outName] = process.argv;
if (!sample || !sheetKey || !outName) {
  console.error('用法: node tools/shot_sheet.js <样本> <表名关键字> <输出名>');
  process.exit(1);
}
const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1680, height: 1000 }, deviceScaleFactor: 2 });
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message.slice(0, 140)));
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });

  await p.evaluate((f) => APP.loadSample('samples/' + f, f), sample);
  await p.waitForTimeout(6000);

  const info = await p.evaluate((key) => {
    const out = {};
    // SJS：跳过试用授权表，按名称匹配
    const sp = APP.sjs;
    let sjsIdx = -1, sjsNames = [];
    for (let i = 0; i < sp.getSheetCount(); i++) {
      const n = sp.getSheet(i).name();
      sjsNames.push(n);
      if (sjsIdx < 0 && n.includes(key) && !/evaluation/i.test(n)) sjsIdx = i;
    }
    if (sjsIdx >= 0) {
      sp.setActiveSheetIndex(sjsIdx);
      const s = sp.getSheet(sjsIdx);
      // 条件格式类表通常数据在 4 行以后，滚到顶部并缩小一点，保证一屏看到规则区
      s.setActiveCell(0, 0);
      // showCell 挂在 sheet 上，不在 workbook 上（workbook 上没有这个方法）
      try { s.showCell(0, 0, GC.Spread.Sheets.VerticalPosition.top, GC.Spread.Sheets.HorizontalPosition.left); } catch (e) {}
    }
    out.sjsIdx = sjsIdx; out.sjsNames = sjsNames;

    // Univer：按名称匹配
    const wb = window.univerAPI.getActiveWorkbook();
    const sheets = wb.getSheets();
    let uniIdx = -1, uniNames = [];
    sheets.forEach((s, i) => {
      const n = s.getSheetName();
      uniNames.push(n);
      if (uniIdx < 0 && n.includes(key)) uniIdx = i;
    });
    if (uniIdx >= 0) wb.setActiveSheet(sheets[uniIdx]);
    out.uniIdx = uniIdx; out.uniNames = uniNames;
    return out;
  }, sheetKey);

  console.log(JSON.stringify(info, null, 1));
  await p.waitForTimeout(2000);
  // 可选第 4 个参数：缩放比（如 0.35）。图表排在工作表下方时，缩到能一屏看全比滚动更可靠：
  // 两侧滚轮步长不等，滚完取景位置就对不上了（实测踩过）。
  const zoom = Number(process.argv[5] || 0);
  if (zoom > 0) {
    const z = await p.evaluate((ratio) => {
      const out = {};
      try { const sp = APP.sjs; sp.getActiveSheet().zoom(ratio); out.sjs = sp.getActiveSheet().zoom(); } catch (e) { out.eSjs = String(e.message).slice(0, 80); }
      try { const sh = APP.uni.getActiveWorkbook().getActiveSheet(); sh.zoom(ratio); out.uni = sh.getZoom ? sh.getZoom() : 'set'; } catch (e) { out.eUni = String(e.message).slice(0, 80); }
      return out;
    }, zoom);
    console.log('zoom', JSON.stringify(z));
    await p.waitForTimeout(2000);
  }
  await p.screenshot({ path: path.join(ROOT, 'out/shots', `_${outName}.png`) });
  console.log('shot -> out/shots/_' + outName + '.png', 'errors:', errs.length);
  await b.close();
})();
