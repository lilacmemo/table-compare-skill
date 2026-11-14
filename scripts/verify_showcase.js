// 成果展示工程冒烟验证：5 个 tab 逐个切换，统计 h2/table/img 数量、pageerror、图片是否真的解码成功。
const { chromium } = require('playwright');
const path = require('path');

const FILE = 'file:///' + path.join(__dirname, '..', '成果展示工程', 'index.html').replace(/\\/g, '/');

(async () => {
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('console', (m) => { if (m.type() === 'error') errors.push('console: ' + m.text()); });

  await page.goto(FILE, { waitUntil: 'load' });
  await page.waitForTimeout(800);

  const keys = await page.$$eval('nav a', (as) => as.map((a) => [a.dataset.target, a.textContent.trim()]));
  console.log('nav:', keys.map((k) => k.join('=')).join(' | '));

  const rows = [];
  for (const [key, label] of keys) {
    await page.click(`nav a[data-target="${key}"]`);
    await page.waitForTimeout(400);
    const stat = await page.evaluate((k) => {
      const d = document.getElementById('doc-' + k);
      if (!d) return { missing: true };
      const imgs = Array.from(d.querySelectorAll('img'));
      return {
        visible: d.classList.contains('visible'),
        h2: d.querySelectorAll('h2').length,
        tables: d.querySelectorAll('table').length,
        imgs: imgs.length,
        imgOk: imgs.filter((i) => i.complete && i.naturalWidth > 0).length,
        imgBad: imgs.filter((i) => !i.complete || i.naturalWidth === 0).map((i) => (i.alt || i.src).slice(0, 60)),
        chars: d.textContent.length,
      };
    }, key);
    rows.push([label, key, stat]);
  }

  for (const [label, key, s] of rows) {
    console.log(
      `${s.missing ? 'MISSING' : s.visible ? 'OK  ' : 'HIDDEN'} ${key.padEnd(8)} ${String(label).padEnd(22)} ` +
      `h2=${s.h2} table=${s.tables} img=${s.imgOk}/${s.imgs} chars=${s.chars}` +
      (s.imgBad && s.imgBad.length ? '  BAD=' + s.imgBad.join(',') : '')
    );
  }

  // 只有第一个 tab 应可见
  const visibleCount = await page.$$eval('.doc.visible', (d) => d.length);
  console.log('visible docs:', visibleCount);
  console.log('pageerrors:', errors.length, errors.slice(0, 5));
  await browser.close();
})();
