// 把样本导入 Univer 在线（office.univer.ai/zh-CN/exchange），再原样导出 xlsx
// 用法：node tools/uni_roundtrip.js [样本路径]
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const FILE = process.argv[2] || 'D:/workbuddy/spreadjs vs univer/测试工程/samples/格式与条件格式-全量用例-A至K.xlsx';
const OUT = path.resolve(__dirname, 'out', 'univer-reexport.xlsx');
(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ viewport: { width: 1500, height: 950 }, acceptDownloads: true });
  const pg = await ctx.newPage();
  await pg.goto('https://office.univer.ai/zh-CN/exchange', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await pg.waitForTimeout(6000);
  await pg.setInputFiles('input[type=file]', FILE);
  await pg.waitForFunction(() => document.body.innerText.includes('导入完成'), null, { timeout: 180000 });
  await pg.waitForTimeout(4000);
  await pg.getByText('下载文件', { exact: true }).first().click();
  await pg.waitForTimeout(1500);
  const [dl] = await Promise.all([
    pg.waitForEvent('download', { timeout: 60000 }),
    pg.getByText('XLSX', { exact: true }).first().click(),
  ]);
  await dl.saveAs(OUT);
  console.log('已保存', OUT, fs.statSync(OUT).size, 'bytes');
  await b.close();
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
