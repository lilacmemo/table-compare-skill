const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const slugs = ['clipboard-and-paste-special', 'permission'];
  const outAll = {};
  for (const slug of slugs) {
    const go = async () => {
      await page.goto('https://office.univer.ai/zh-CN/playground/sheets/' + slug, { waitUntil: 'domcontentloaded', timeout: 60000 });
      for (let i = 0; i < 40; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
      await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
      for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
      return false;
    };
    if (!await go()) { console.log('fail', slug); continue; }
    await page.waitForTimeout(2500);
    const r = await page.evaluate(() => {
      const allMethods = (o) => { if (!o) return []; const s = new Set(); let p = Object.getPrototypeOf(o); while (p && p !== Object.prototype) { Object.getOwnPropertyNames(p).forEach((n) => s.add(n)); p = Object.getPrototypeOf(p); } return [...s].filter((n) => typeof o[n] === 'function' && n[0] !== '_'); };
      const u = window.univerAPI; const wb = u.getActiveWorkbook(); const sh = wb.getActiveSheet();
      const perm = sh.getWorksheetPermission ? sh.getWorksheetPermission() : null;
      return { sh: allMethods(sh), wb: allMethods(wb), u: allMethods(u), perm: allMethods(perm), range: allMethods(sh.getRange(0, 0, 1, 1)) };
    });
    outAll[slug] = r;
    console.log(slug, 'sh', r.sh.length, 'wb', r.wb.length, 'u', r.u.length, 'range', r.range.length);
  }
  fs.writeFileSync('_pro_surface.json', JSON.stringify(outAll, null, 1), 'utf8');
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
