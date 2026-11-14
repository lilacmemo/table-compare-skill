// 线上 Pro 补测：行/列分组（outline）—— 本地 OSS bundle 无 @univerjs-pro/sheets-outline
//   node probe_pro_outline.js
//
// 【2026-11-06 修正】旧地址 /zh-CN/playground/sheets 已被站点改版下线（401 / "Showcase not found"）。
// 现在每个示例页 = /zh-CN/showcase/sheets/<slug>，编辑器本体在 iframe /zh-CN/playground/sheets/<slug>。
// 直接打开 playground 深链即可拿到 window.univerAPI。
const { chromium } = require('playwright');

const URL = 'https://office.univer.ai/zh-CN/playground/sheets/outline';

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await browser.newContext({ viewport: { width: 1680, height: 1050 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', (e) => errs.push(String(e).slice(0, 160)));

  const waitApi = async (tries) => {
    for (let i = 0; i < tries; i++) {
      if (await page.evaluate(() => !!window.univerAPI)) return true;
      await page.waitForTimeout(1000);
    }
    return false;
  };

  await page.goto(URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  let ok = await waitApi(45);
  if (!ok) {
    console.log('first load stalled, reloading once...');
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    ok = await waitApi(60);
  }
  console.log('univerAPI ready:', ok);
  if (!ok) { console.log(errs.slice(0, 3)); await browser.close(); return; }
  await page.waitForTimeout(2500);

  const out = await page.evaluate(async () => {
    const u = window.univerAPI;
    const wb = u.getActiveWorkbook();
    const E = u.Enum || {};
    const tick = (ms) => new Promise((r) => setTimeout(r, ms || 300));
    const tryIt = async (fn) => { try { return await fn(); } catch (e) { return 'ERR ' + (e && e.message || e); } };

    const R = {
      env: {
        url: location.pathname,
        enumAxis: E.DimensionOutlineAxis || null,
        srcSheet: wb.getActiveSheet().getSheetName(),
        hasAddRow: typeof wb.getActiveSheet().addRowOutline,
      },
    };

    // 独立临时表，避免和示例自带的大纲数据混在一起
    // FWorkbook 没有 createSheet，建表用 insertSheet(name)
    const sh = wb.insertSheet('__pE');
    wb.setActiveSheet(sh);
    await tick();
    const outline = (axis) => sh.getDimensionOutlines(axis) || [];

    // E1 行分组
    R.E1 = await tryIt(async () => {
      const ret = sh.addRowOutline(2, 5);
      await tick();
      const o = outline(E.DimensionOutlineAxis.ROW);
      return { ret: ret === sh ? 'chain-ok' : String(ret), count: o.length, first: o[0] ? JSON.stringify(o[0]) : null };
    });

    // E2 列分组
    R.E2 = await tryIt(async () => {
      sh.addColumnOutline(1, 4);
      await tick();
      const o = outline(E.DimensionOutlineAxis.COLUMN);
      return { count: o.length, first: o[0] ? JSON.stringify(o[0]) : null };
    });

    // E4 折叠 / 展开
    R.E4 = await tryIt(async () => {
      const o = outline(E.DimensionOutlineAxis.ROW);
      if (!o[0]) return 'NO_OUTLINE';
      const before = o[0].collapsed;
      sh.setDimensionOutlineCollapsed(o[0].id, true);
      await tick();
      const mid = (outline(E.DimensionOutlineAxis.ROW)[0] || {}).collapsed;
      sh.setDimensionOutlineCollapsed(o[0].id, false);
      await tick();
      const back = (outline(E.DimensionOutlineAxis.ROW)[0] || {}).collapsed;
      return { before, collapsed: mid, expanded: back };
    });

    // E5 两级嵌套
    R.E5 = await tryIt(async () => {
      sh.addRowOutline(1, 12);
      sh.addRowOutline(3, 4);
      await tick();
      const o = outline(E.DimensionOutlineAxis.ROW);
      return {
        count: o.length,
        list: o.map((x) => ({ s: x.start, e: x.end, level: x.level, collapsed: x.collapsed })),
      };
    });

    // E3 删除单个
    R.E3 = await tryIt(async () => {
      const o = outline(E.DimensionOutlineAxis.ROW);
      if (!o.length) return 'NOTHING';
      const id = o[o.length - 1].id;
      sh.removeDimensionOutline(id);
      await tick();
      return { before: o.length, after: outline(E.DimensionOutlineAxis.ROW).length };
    });

    // E6 批量清（区间）
    R.E6 = await tryIt(async () => {
      sh.clearDimensionOutlines(E.DimensionOutlineAxis.ROW, 0, 200);
      await tick();
      return { rowAfter: outline(E.DimensionOutlineAxis.ROW).length };
    });

    // 顺带确认 Pro 侧其它入口（本地 OSS 为 undefined 的那些）
    R.also = {
      sparkline: typeof sh.addSparkline,
      shape: typeof sh.insertShape,
      chart: typeof sh.insertChart || typeof sh.newChart,
      print: typeof wb.print,
      printCfg: typeof wb.updatePrintConfig,
      pivot: typeof sh.generatePivotTable,
      findReplace: typeof u.createTextFinderAsync,
      crosshair: typeof u.setCrosshairHighlightEnabled,
      permWs: typeof sh.getWorksheetPermission,
      permWb: typeof wb.getWorkbookPermission,
      protect: typeof sh.protect,
      bgImage: typeof sh.setBackgroundImage,
      filter: typeof sh.setFilter,
      validation: typeof sh.setDataValidation,
      slicer: typeof sh.insertSlicer,
    };

    // 顺手把旧环境（charts 示例）也确认一下入口还在，供报告"环境已迁移"备注
    R.alsoChartsPage = { note: 'see run separately' };
    return R;
  });

  console.log(JSON.stringify(out, null, 2));
  if (errs.length) console.log('pageerrors:', errs.slice(0, 3));
  await page.screenshot({ path: '_pro_outline.png' });
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
