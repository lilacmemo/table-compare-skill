// 线上 Pro 逐示例页入口探测：每个 showcase 示例页只装该例所需插件，
// 因此"某方法在对应示例页上存在" = Pro 有该能力；
// 反之"在无关示例页上不存在"不构成缺失证据。
//   node probe_pro_caps.js
const { chromium } = require('playwright');

const BASE = 'https://office.univer.ai/zh-CN/playground/sheets/';

// slug -> 要探的方法（返回 typeof 或值）
const PLAN = [
  { slug: 'outline', label: '行/列分组', probe: (sh, wb, u) => ({ addRowOutline: typeof sh.addRowOutline, addColumnOutline: typeof sh.addColumnOutline }) },
  { slug: 'freeze-panes', label: '冻结', probe: (sh) => ({ setFrozenRows: typeof sh.setFrozenRows, getFreeze: typeof sh.getFreeze }) },
  { slug: 'permission', label: '保护/权限', probe: (sh, wb) => ({ protect: typeof sh.protect, getWorksheetPermission: typeof sh.getWorksheetPermission, getWorkbookPermission: typeof wb.getWorkbookPermission }) },
  { slug: 'sparkline-types', label: '迷你图', probe: (sh) => ({ addSparkline: typeof sh.addSparkline }) },
  { slug: 'shapes', label: '形状', probe: (sh) => ({ insertShape: typeof sh.insertShape }) },
  { slug: 'chart-family-pies-and-donuts', label: '图表', probe: (sh) => ({ insertChart: typeof sh.insertChart, newChart: typeof sh.newChart }) },
  { slug: 'print', label: '打印', probe: (sh, wb) => ({ print: typeof wb.print, updatePrintConfig: typeof wb.updatePrintConfig }) },
  { slug: 'pivot-table-layouts', label: '数据透视', probe: (sh) => ({ generatePivotTable: typeof sh.generatePivotTable, insertPivotTable: typeof sh.insertPivotTable }) },
  { slug: 'clipboard-and-paste-special', label: '剪贴板', probe: (sh, wb, u) => ({ copy: typeof u.copy, paste: typeof u.paste }) },
  { slug: 'images', label: '背景图/图片', probe: (sh) => ({ setBackgroundImage: typeof sh.setBackgroundImage }) },
];

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await browser.newContext({ viewport: { width: 1680, height: 1050 } });
  const results = [];

  for (const p of PLAN) {
    const page = await ctx.newPage();
    const errs = [];
    page.on('pageerror', (e) => errs.push(String(e).slice(0, 120)));
    const url = BASE + p.slug;
    let ready = false;
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      for (let i = 0; i < 40; i++) { if (await page.evaluate(() => !!window.univerAPI)) { ready = true; break; } await page.waitForTimeout(1000); }
      if (!ready) {
        await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
        for (let i = 0; i < 50; i++) { if (await page.evaluate(() => !!window.univerAPI)) { ready = true; break; } await page.waitForTimeout(1000); }
      }
      let data = null;
      if (ready) {
        await page.waitForTimeout(2000);
        data = await page.evaluate((src) => {
          // eslint-disable-next-line no-new-func
          const fn = new Function('sh', 'wb', 'u', 'return (' + src + ')(sh, wb, u)');
          const u = window.univerAPI;
          const wb = u.getActiveWorkbook();
          const sh = wb.getActiveSheet();
          try { return fn(sh, wb, u); } catch (e) { return { ERR: String(e).slice(0, 120) }; }
        }, p.probe.toString());
      }
      results.push({ slug: p.slug, label: p.label, ready, data, errs: errs.slice(0, 2) });
    } catch (e) {
      results.push({ slug: p.slug, label: p.label, ready: false, error: String(e).slice(0, 120) });
    }
    console.log('done', p.slug, ready ? 'OK' : 'NO-API', JSON.stringify(results[results.length - 1].data));
    await page.close();
  }

  console.log('\n=== SUMMARY ===');
  console.log(JSON.stringify(results, null, 2));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
