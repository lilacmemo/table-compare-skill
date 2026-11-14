const { chromium } = require('playwright');
const get = async (page, slug) => {
  await page.goto('https://office.univer.ai/zh-CN/playground/sheets/' + slug, { waitUntil: 'domcontentloaded', timeout: 60000 });
  let ok = false;
  for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) { ok = true; break; } await page.waitForTimeout(1000); }
  if (!ok) { await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 }); for (let i = 0; i < 50; i++) { if (await page.evaluate(() => !!window.univerAPI)) { ok = true; break; } await page.waitForTimeout(1000); } }
  return ok;
};
const allMethods = (obj) => { const s = new Set(); let p = Object.getPrototypeOf(obj); while (p && p !== Object.prototype) { Object.getOwnPropertyNames(p).forEach((n) => s.add(n)); p = Object.getPrototypeOf(p); } return [...s].filter((n) => typeof obj[n] === 'function'); };
(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();

  let ok = await get(page, 'permission');
  console.log('permission page ready', ok);
  if (ok) {
    await page.waitForTimeout(2500);
    const r = await page.evaluate((src) => {
      const allMethods = new Function('return ' + src)();
      const u = window.univerAPI; const wb = u.getActiveWorkbook(); const sh = wb.getActiveSheet();
      const perm = sh.getWorksheetPermission ? sh.getWorksheetPermission() : null;
      const wperm = wb.getWorkbookPermission ? wb.getWorkbookPermission() : null;
      return {
        sheetName: sh.getSheetName(),
        sheetPermMethods: perm ? allMethods(perm).filter((n) => /permi|edit|read|point|role/i.test(n)) : null,
        permProtoKeys: perm ? Object.keys(perm) : null,
        wbPermProtoKeys: wperm ? Object.keys(wperm) : null,
        wbPermMethods: wperm ? allMethods(wperm).filter((n) => /permi|edit|read|point|role/i.test(n)) : null,
        sheetProtectLike: allMethods(sh).filter((n) => /protect|lock|permi|editable/i.test(n)),
        wbProtectLike: allMethods(wb).filter((n) => /protect|lock|permi|editable/i.test(n)),
        enumPoints: { ws: Object.keys(u.Enum.WorksheetPermissionPoint || {}), wb: Object.keys(u.Enum.WorkbookPermissionPoint || {}) },
      };
    }, allMethods.toString());
    console.log(JSON.stringify(r, null, 2).slice(0, 4000));
  }

  ok = await get(page, 'pivot-table-layouts');
  console.log('\npivot page ready', ok);
  if (ok) {
    await page.waitForTimeout(2500);
    const r = await page.evaluate((src) => {
      const allMethods = new Function('return ' + src)();
      const u = window.univerAPI; const wb = u.getActiveWorkbook(); const sh = wb.getActiveSheet();
      return {
        sheetPivotLike: allMethods(sh).filter((n) => /pivot/i.test(n)),
        wbPivotLike: allMethods(wb).filter((n) => /pivot/i.test(n)),
        apiPivotLike: allMethods(u).filter((n) => /pivot/i.test(n)),
        sheetName: sh.getSheetName(),
      };
    }, allMethods.toString());
    console.log(JSON.stringify(r, null, 2).slice(0, 2500));
  }
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
