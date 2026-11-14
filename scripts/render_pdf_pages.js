// 用 Chromium 渲染 PDF 逐页并裁出页面本体（PDF 查看器的灰色背景与工具栏去掉）
//   node tools/render_pdf_pages.js <pdf> <前缀> [最多页数]
// 为什么要 Chromium 而不是 pymupdf：
//   SpreadJS 导出的 PDF 嵌的是 TrueType 子集（CZZZZZ+SimHei），pymupdf/mupdf 渲染时
//   会取错字形、中文全变乱码；而文本层是好的、Chromium 打开完全正常。
//   本轮差点据此写下"SJS 导出 PDF 中文乱码"的错误结论——所以渲染一定要有第二渲染器交叉验证。
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const [, , pdf, prefix, maxPagesArg] = process.argv;
const MAX = Number(maxPagesArg || 6);
const ROOT = path.join(__dirname, '..');
const PDF = path.resolve(pdf);

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ viewport: { width: 1500, height: 1150 }, deviceScaleFactor: 2 });
  const p = await ctx.newPage();
  for (let i = 1; i <= MAX; i++) {
    // PDF 查看器支持 #page=N&zoom=100；用 toolbar=0 去掉顶部工具栏，页面更干净
    const url = 'file:///' + PDF.replace(/\\/g, '/') + `#page=${i}&zoom=100&toolbar=0&navpanes=0`;
    await p.goto(url, { waitUntil: 'networkidle' });
    await p.waitForTimeout(2500);
    const shot = path.join(ROOT, 'out', 'shots', `_${prefix}-raw-p${i}.png`);
    await p.screenshot({ path: shot });
    const done = await p.evaluate(() => document.querySelector('embed'));
    // 判断这一页是否真的存在：PDF 空白页/超出页数时查看器显示灰底
    const buf = fs.readFileSync(shot);
    if (buf.length < 12000) { fs.unlinkSync(shot); break; }
    console.log('page', i, '->', path.relative(ROOT, shot));
  }
  await b.close();

  // 裁白页：查看器背景是灰色，页面是白纸，取最大近白矩形
  const { execFileSync } = require('child_process');
  try {
    execFileSync('python', [path.join(ROOT, 'tools', '_crop_pdf_pages.py'), prefix], { stdio: 'inherit' });
  } catch (e) {
    console.log('crop 失败', e.message);
  }
})();
