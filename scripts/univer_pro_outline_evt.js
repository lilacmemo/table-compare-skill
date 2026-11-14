const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await (await browser.newContext({ viewport: { width: 1680, height: 1050 } })).newPage();
  const go = async () => {
    await page.goto('https://office.univer.ai/zh-CN/playground/sheets/outline', { waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 45; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    await page.reload({ waitUntil: 'domcontentloaded', timeout: 60000 });
    for (let i = 0; i < 50; i++) { if (await page.evaluate(() => !!window.univerAPI)) return true; await page.waitForTimeout(1000); }
    return false;
  };
  if (!await go()) { console.log('no api'); await browser.close(); return; }
  await page.waitForTimeout(2500);
  const r = await page.evaluate(async () => {
    const u = window.univerAPI; const wb = u.getActiveWorkbook(); const sh = wb.getActiveSheet();
    const evtKeys = u.Event ? Object.keys(u.Event) : null;
    const outlineEvt = evtKeys ? evtKeys.filter((k) => /outline|group|dimension|collapse|hidden/i.test(k)) : null;

    // 嵌套深度实验
    const ns = wb.insertSheet('__pDepth'); wb.setActiveSheet(ns);
    const log = [];
    let hit = 0;
    for (let i = 0; i < 8; i++) {
      try {
        ns.addRowOutline(i, 20 - i * 2);
        const o = ns.getDimensionOutlines('row') || [];
        log.push({ i, count: o.length, maxNest: o.reduce((a, x) => a + (x.start <= i && x.end >= 20 - i * 2 ? 1 : 0), 0) });
      } catch (e) { log.push({ i, err: String(e && e.message || e).slice(0, 90) }); break; }
    }
    let errHit = null;
    try { ns.addRowOutline(0, 30); } catch (e) { errHit = String(e && e.message || e).slice(0, 120); }

    return { outlineEvt, totalEvtKeys: evtKeys ? evtKeys.length : null, depthLog: log, crossErr: errHit };
  });
  console.log(JSON.stringify(r, null, 2));
  await browser.close();
})().catch((e) => { console.error(e); process.exit(1); });
