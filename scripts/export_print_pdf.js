// 各家打印输出的 PDF（效果层输出对照，基准 = Excel 的 ExportAsFixedFormat）
//   SJS：spread.savePDF（需 gc.spread.sheets.pdf 插件）
//   Univer：Pro 的 print() 走浏览器打印，无 PDF 导出 API → 记为 N/A，只能截对话框预览
// 用法：NODE_PATH=<poc-compare/node_modules> node tools/export_print_pdf.js
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE = 'https://127.0.0.1:8124/' + encodeURIComponent('对比工程') + '/index.html?autoload=0';
const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'out', 'pdf');
const SAMPLE = 'samples/打印-PDF-多元素样本.xlsx';

(async () => {
  const b = await chromium.launch({ channel: 'msedge', headless: true });
  const ctx = await b.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1600, height: 1000 } });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => console.log('PAGEERROR', e.message.slice(0, 140)));
  await p.goto(BASE, { waitUntil: 'networkidle' });
  await p.waitForFunction(() => window.APP && APP.booted === true, null, { timeout: 90000 });
  await p.evaluate(async (f) => APP.loadSample(f, '打印-PDF-多元素样本.xlsx'), SAMPLE);
  await p.waitForTimeout(1500);

  // 先注册中文字体：gc.spread.sheets.pdf 只内置 14 种西文标准字体，不注册中文
  // 就会导出乱码（内嵌字体表里只剩 Times-Roman / Times-Bold，实测过）。
  const font = await p.evaluate(async () => {
    const out = {};
    try {
      const fb = await fetch('vendor/fonts/simhei.ttf').then((r) => r.arrayBuffer());
      out.bytes = fb.byteLength;
      const F = GC.Spread.Sheets.PDF.PDFFontsManager;
      // 工作表里写的是「宋体/微软雅黑/等线」，注册名要和单元格字体名对得上才命中，
      // 所以把同一份字体挂到这几个常用中文名下（skill 里同样的配方）
      ['SimHei', '宋体', '微软雅黑', '等线'].forEach((n) => F.registerFont(n, { normal: fb, bold: fb }));
      // 西文也兜一下，避免 Times 之外的字形缺字
      ['Arial', 'Calibri'].forEach((n) => F.registerFont(n, { normal: fb, bold: fb }));
      out.ok = true;
    } catch (e) { out.err = e.message.slice(0, 140); }
    return out;
  });
  console.log('FONT', JSON.stringify(font));

  const r = await p.evaluate(() => new Promise((resolve) => {
    const out = { hasSavePDF: typeof window.__spread.savePDF };
    const isEval = (n) => /evaluation/i.test(String(n || ''));
    const sp = APP.sjs;
    // 去掉试用版插入的授权提示表，否则 PDF 第一页会是它
    for (let i = sp.getSheetCount() - 1; i >= 0; i--) {
      if (isEval(sp.getSheet(i).name())) { try { sp.removeSheet(i); } catch (e) {} }
    }
    out.sheets = Array.from({ length: sp.getSheetCount() }, (_, i) => sp.getSheet(i).name());
    if (out.hasSavePDF !== 'function') { resolve(out); return; }
    // savePDF 的成功回调给的是 **Blob**（不是 base64），必须用 FileReader 转出来
    sp.savePDF((blob) => {
      out.blobType = Object.prototype.toString.call(blob);
      out.bytes = blob && blob.size;
      const fr = new FileReader();
      fr.onload = () => { out.b64 = String(fr.result); resolve(out); };
      fr.onerror = () => { out.err = 'FileReader failed'; resolve(out); };
      fr.readAsDataURL(blob);
    }, (err) => { out.err = JSON.stringify(err).slice(0, 200); resolve(out); }, {});
    setTimeout(() => { out.timeout = true; resolve(out); }, 60000);
  }));
  if (r.b64) {
    fs.writeFileSync(path.join(OUT, 'sjs-print.pdf'), Buffer.from(String(r.b64).replace(/^data:[^,]+,/, ''), 'base64'));
    delete r.b64;
  }
  console.log(JSON.stringify({ ...r, b64: undefined }));
  await b.close();
})();
