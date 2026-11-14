// 导入 / 导出兼容性抓取
//   NODE_PATH=../poc-compare/node_modules node tools/capture_io.js [module...]
// 产物：
//   out/shots/<module>-<side>.png        导入后的渲染截图（SJS / Univer）
//   out/shots/<module>-<side>-rt.png     **导出文件重新打开后**的截图（往返验证）
//   out/export/<module>-<side>.xlsx      各自导出的文件（用于 XML 级差异比对）
//   out/data/<module>-io.json            导入耗时、对象计数、导出字节数、往返结果
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const MODS = process.argv.slice(2).filter((a) => !a.startsWith('--'));
const TARGETS = MODS.length ? MODS : ['charts', 'print', 'worksheet', 'format', 'table', 'formula'];
const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html';
const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(ROOT, 'out', 'shots');
const EXPD = path.join(ROOT, 'out', 'export');
const DATA = path.join(ROOT, 'out', 'data');

const saveDataUrl = (dataUrl, file) => {
  const b64 = String(dataUrl).replace(/^data:[^,]+,/, '');
  fs.writeFileSync(file, Buffer.from(b64, 'base64'));
};

// 让两边都停在**同一张表**的 A1、适应窗口，截图才可比
// 坑 1：setActiveSheet(sheet 对象) 在这版里不刷新画布，要用 setActiveSheetIndex
// 坑 2：图表样本两家的表顺序不同（SJS 首表"组合图"、Univer 首表"新式图表"），
//       必须按**表名**对齐，不能各取第 0 张
// 坑 3：SpreadJS 试用版会把 "Evaluation Version" 授权提示表插在**第一张**（导出文件里也有），
//       直接 setActiveSheetIndex(0) 会截到那张空白授权页 —— 必须跳过它
const isEvalSheet = (n) => /evaluation/i.test(String(n || ''));

const alignViews = (page, wantName) => page.evaluate((want) => {
  const isEval = (n) => /evaluation/i.test(String(n || ''));
  let first = want || null;
  try {
    const sp = APP.sjs;
    const n = sp.getSheetCount();
    let idx = -1;
    if (first) {
      for (let i = 0; i < n; i++) { if (sp.getSheet(i).name() === first) { idx = i; break; } }
    }
    if (idx < 0) {
      for (let i = 0; i < n; i++) {
        const nm = sp.getSheet(i).name();
        if (!isEval(nm)) { idx = i; first = nm; break; }
      }
    }
    if (idx < 0) idx = 0;
    sp.setActiveSheetIndex(idx);
    const sh = sp.getSheet(idx);
    first = first || sh.name();
    sh.setActiveCell(0, 0);
    // 视口复位有两个坑：
    // ① setActiveCell 只移焦点，不回滚视口；
    // ② 有冻结行时（公式样本冻结 4 行），showCell(0,0) 作用在冻结区，滚不动下面的滚动视口，
    //    要把**第一个非冻结行**顶到顶部才行（实测 viewportTopRow(1) 从 243 变 4）。
    // 另外 xlsx 里存了上次的滚动位置，SpreadJS 会还原（Univer 不还原），不复位两边截图就对不上。
    try {
      const fr = sh.frozenRowCount ? sh.frozenRowCount() : 0;
      sh.showRow(fr, GC.Spread.Sheets.VerticalPosition.top);
      sh.showColumn(0, GC.Spread.Sheets.HorizontalPosition.left);
    } catch (e) {}
    sh.zoom(0.8);
  } catch (e) {}
  try {
    const wb = APP.uni.getActiveWorkbook();
    const target = (first && wb.getSheetByName(first)) || wb.getSheets()[0];
    wb.setActiveSheet(target);
    target.getRange('A1').activate();
  } catch (e) {}
  window.__alignedSheet = first;
}, wantName || null);

// 数一遍两侧的工作表名，用来取证「导出的文件里到底多没多表」
const sheetNames = (page, side) => page.evaluate((s) => {
  try {
    if (s === 'sjs') {
      const sp = APP.sjs, out = [];
      for (let i = 0; i < sp.getSheetCount(); i++) out.push(sp.getSheet(i).name());
      return out;
    }
    const wb = APP.uni.getActiveWorkbook();
    return wb.getSheets().map((x) => x.getSheetName());
  } catch (e) { return null; }
}, side);

(async () => {
  [SHOTS, EXPD, DATA].forEach((d) => fs.mkdirSync(d, { recursive: true }));
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1680, height: 1000 } });
  const page = await ctx.newPage();
  page.on('pageerror', (e) => console.log('PAGEERR', String(e).slice(0, 160)));

  const all = {};
  for (const mod of TARGETS) {
    process.stdout.write('\n=== ' + mod + ' ===\n');
    await page.goto(BASE + '?module=' + mod, { waitUntil: 'domcontentloaded', timeout: 120000 });
    for (let i = 0; i < 120; i++) {
      const ok = await page.evaluate(() => window.APP && APP.booted && APP.ready.sjs && APP.ready.uni);
      if (ok) break;
      await page.waitForTimeout(1000);
    }
    for (let i = 0; i < 120; i++) {
      const done = await page.evaluate(() => !!(APP.lastImport && APP.lastImport.sjs && APP.lastImport.uni));
      if (done) break;
      await page.waitForTimeout(1000);
    }
    const info = await page.evaluate(() => ({ sample: APP.sample && APP.sample.name, imp: APP.lastImport }));
    if (!info.imp.uni && !info.imp.sjs) { console.log('  导入失败，跳过'); continue; }
    console.log('  样本', info.sample, 'SJS', info.imp.sjs, 'ms / Univer', info.imp.uni, 'ms');

    // 让两边都停在**同一张表**的 A1、适应窗口，截图才可比（对齐逻辑见文件顶部 alignViews）
    await alignViews(page);
    await page.waitForTimeout(1600);

    await page.locator('#paneSjs').screenshot({ path: path.join(SHOTS, mod + '-sjs.png') });
    await page.locator('#paneUni').screenshot({ path: path.join(SHOTS, mod + '-uni.png') });

    const stats = await page.evaluate(() => ({ sjs: APP.statsSjs(), uni: APP.statsUni() }));

    // 导出
    const exp = {};
    try {
      const u = await page.evaluate(() => APP.exportSjs());
      const f = path.join(EXPD, mod + '-sjs.xlsx');
      saveDataUrl(u, f);
      exp.sjs = { file: f, bytes: fs.statSync(f).size };
    } catch (e) { exp.sjs = { error: e.message.slice(0, 200) }; }
    try {
      const u = await page.evaluate(() => APP.exportUni());
      const f = path.join(EXPD, mod + '-uni.xlsx');
      saveDataUrl(u, f);
      exp.uni = { file: f, bytes: fs.statSync(f).size };
    } catch (e) { exp.uni = { error: e.message.slice(0, 200) }; }
    console.log('  导出 SJS', exp.sjs.bytes || exp.sjs.error, '| Univer', exp.uni.bytes || exp.uni.error);

    // 导出侧取证：把各家刚导出的 xlsx 重新打开，单独截图。
    // 这一步回答的是「导出的文件本身还能不能读、内容还在不在」，
    // 和上面的「导入后视图」合起来才是完整的导入 / 导出两个环节。
    const rt = {};
    const wantSheet = await page.evaluate(() => window.__alignedSheet || null);
    for (const side of ['sjs', 'uni']) {
      if (!(exp[side] && exp[side].file)) { rt[side] = { error: '导出失败，跳过' }; continue; }
      const rel = 'out/export/' + mod + '-' + side + '.xlsx';
      const t0 = Date.now();
      try {
        // 注意：page.evaluate 只接受**一个**参数，多参会抛 "Too many arguments"
        await page.evaluate(({ s, u }) => APP.importOne(s, u, u), { s: side, u: rel });
        // 按第一次导入时对齐的那张表来对，避开试用版插在最前面的授权提示表
        await alignViews(page, wantSheet);
        await page.waitForTimeout(1600);
        const shot = mod + '-' + side + '-rt.png';
        await page.locator(side === 'sjs' ? '#paneSjs' : '#paneUni')
          .screenshot({ path: path.join(SHOTS, shot) });
        const st = await page.evaluate(
          (s) => (s === 'sjs' ? APP.statsSjs() : APP.statsUni()), side);
        const names = await sheetNames(page, side);
        rt[side] = {
          ms: Date.now() - t0, shot,
          totals: (st && st.totals) || null,
          sheets: names ? names.length : null,
          // 导出文件里是否夹带试用版授权表（SpreadJS 独有，正式授权版不会有）
          evalSheet: names ? names.filter(isEvalSheet) : null,
          sheetNames: names ? names.slice(0, 3) : null,
        };
        console.log('  往返 ' + side + ' 重开成功', Date.now() - t0 + 'ms', shot,
          '表数', names ? names.length : '?',
          names && names.filter(isEvalSheet).length ? '（含授权提示表）' : '');
      } catch (e) {
        rt[side] = { error: String(e.message).slice(0, 200) };
        console.log('  往返 ' + side + ' 失败：', String(e.message).slice(0, 160));
      }
    }

    all[mod] = { sample: info.sample, importMs: info.imp, stats, export: exp, roundtrip: rt };
    fs.writeFileSync(path.join(DATA, mod + '-io.json'), JSON.stringify(all[mod], null, 1), 'utf8');
  }
  // 合并写入：单模块重跑不能把别的模块结果冲掉（compose 的页脚统计依赖它）
  const allFile = path.join(DATA, '_io-all.json');
  let merged = {};
  try { merged = JSON.parse(fs.readFileSync(allFile, 'utf8')); } catch (e) {}
  Object.assign(merged, all);
  fs.writeFileSync(allFile, JSON.stringify(merged, null, 1), 'utf8');
  console.log('\n写出 out/data/_io-all.json（本次 ' + Object.keys(all).length + ' 个模块，累计 ' + Object.keys(merged).length + ' 个）');
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
