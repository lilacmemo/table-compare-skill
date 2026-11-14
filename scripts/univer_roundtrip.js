// 闭环测试 v3：图表锚点全部放在数据区内（避免导入时工作表裁剪影响）
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const DEMO = 'https://office.univer.ai/zh-CN/playground/sheets/charts';
const OUT = path.join(__dirname, 'univer_roundtrip_v3.xlsx');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();
  page.setDefaultTimeout(60000);
  await page.goto(DEMO, { waitUntil: 'networkidle', timeout: 90000 }).catch(() => {});
  await page.waitForTimeout(15000);

  // Stage A：数据 A1:B6 + 3 图锚在 F1 / F10（数据区附近，导入后仍在范围内）
  const stageA = await page.evaluate(async () => {
    const wb = univerAPI.getActiveWorkbook();
    const sheet = wb.insertSheet('RoundTripV3', { sheet: { rowCount: 30, columnCount: 20 } });
    sheet.getRange('A1:B6').setValues([
      ['月份', '销售额'], ['1月', 120], ['2月', 200], ['3月', 160], ['4月', 240], ['5月', 180],
    ]);
    const mk = (type, title, row, col) => sheet
      .newChart(univerAPI.Enum.ChartTypeString[type])
      .setSource({ range: 'A1:B6', orientation: univerAPI.Enum.ChartSourceOrientation.Columns })
      .setPosition(row, col, 0, 0)
      .setSize(420, 240)
      .setTitle(title)
      .build();
    const ids = [];
    for (const [t, n, r, c] of [['Column', 'RT3-柱形图', 0, 4], ['Line', 'RT3-折线图', 0, 9], ['Pie', 'RT3-饼图', 12, 4]]) {
      const ch = await sheet.insertChart(mk(t, n, r, c));
      ids.push(ch && ch.getId());
    }
    await new Promise(r => setTimeout(r, 2500));
    return { ids, unitId: wb.getId() };
  });
  console.log('STAGE_A:', JSON.stringify(stageA));
  await page.evaluate(() => {
    const wb = univerAPI.getActiveWorkbook();
    const rt = wb.getSheets().find(s => /roundtripv3/i.test(s.getSheetName()));
    wb.setActiveSheet(rt);
  });
  await page.waitForTimeout(2000);
  await page.screenshot({ path: 'univer_v3_A.png' });

  // 导出（快照路线）
  const dl = page.waitForEvent('download', { timeout: 120000 });
  await page.evaluate(async (unitId) => {
    const f = await univerAPI.exportSheetBySnapshotAsync(univerAPI.getActiveWorkbook().save(), 'xlsx');
    if (!f) throw new Error('export null');
    univerAPI.downloadFile(f, 'univer_roundtrip_v3', 'xlsx');
  }, stageA.unitId);
  (await dl).saveAs(OUT);
  await new Promise(r => setTimeout(r, 2000));
  console.log('EXPORTED:', fs.statSync(OUT).size, 'bytes');

  // Stage B：再导入，检查渲染
  const b64 = fs.readFileSync(OUT).toString('base64');
  await page.evaluate(async (b64) => {
    const bin = atob(b64); const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const file = new File([bytes], 'v3.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const snap = await univerAPI.importSheetToSnapshotAsync(file);
    univerAPI.createWorkbook(snap);
    await new Promise(r => setTimeout(r, 8000));
    const wb2 = univerAPI.getActiveWorkbook();
    const rt = wb2.getSheets().find(s => /roundtripv3/i.test(s.getSheetName()));
    wb2.setActiveSheet(rt);
  }, b64);
  await page.waitForTimeout(3000);
  await page.screenshot({ path: 'univer_v3_B_reimport.png' });
  await browser.close();
  console.log('DONE');
})().catch(e => { console.error('FAIL:', e.message); process.exit(1); });
