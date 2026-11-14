/**
 * probe_number_format.js —— 数字格式「格式串 → 实际显示文本」对照探针
 *
 * 为什么需要它：
 *   E1~E25 那批用例原先只校验 setFormatter/getFormatter 的「往返读回」，
 *   格式串存进去再读出来当然一样，测的是"能不能存字符串"，不是"能不能按它显示"。
 *   真正要判的是 getText / getDisplayValue 出来的文本，以及 SpreadJS 到底认哪些 Excel 原生串。
 *
 * 用法：
 *   NODE_PATH=../poc-compare/node_modules node tools/probe_number_format.js [--out tools/out/number-format.json]
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const PORT = Number(process.env.PORT || 8497);
const OUT = path.join(__dirname, 'out', 'number-format.json');
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.png': 'image/png',
};

// [id, 说明, 格式串, 测试值] —— 格式串按「SpreadJS 官方写法」与「Excel 原生存法」分开列，
// 只有分开才能判断"是产品不支持"还是"我们喂错了串"。
const CASES = [
  ['E1', 'General', 'General', 1234.5678],
  ['E2', '#,##0.00', '#,##0.00', 1234.5678],
  ['E3', '货币', '¥#,##0.00;[Red]¥-#,##0.00', -1234.5678],
  ['E4', '会计专用', '_ ¥* #,##0.00_ ;_ ¥* -#,##0.00_ ;_ ¥* "-"??_ ;_ @_ ', 1234.5678],
  ['E7', '百分比', '0.00%', 0.1234],
  ['E8', '分数', '# ?/?', 1234.5678],
  ['E9', '科学记数', '0.00E+00', 1234.5678],
  ['E10', '文本', '@', 1234.5678],
  // 同一件事的三种写法：Excel 原生存法 / 只留区域码 / SpreadJS 文档写法
  ['E11-x', '中文大写·Excel 原生存法', '[DBNum2][$-804]G/通用格式', 12345],
  ['E11-y', '中文大写·仅区域码', '[$-804]G/通用格式', 12345],
  ['E11-z', '中文大写·SpreadJS 文档写法', '[DBNum2][$-804]General', 12345],
  ['E11-w', '中文大写·不带区域码', '[DBNum2]General', 12345],
  ['E11-v', '中文小写 DBNum1', '[DBNum1][$-804]General', 12345],
  ['E12', '0 占位符', '00000', 1234.5678],
  ['E13', '# 占位符', '###0', 1234.5678],
  ['E14', '? 占位符', '0.0??', 1234.5678],
  ['E15', '文本段 @', '0.00"元";@', 1234.5678],
  ['E16', '千分位缩放', '#,##0,,"百万"', 1234.5678],
  ['E17', '多段+负数括号', '#,##0.00_);[Red](#,##0.00)', -1234.5],
  ['E18', '条件段 [>100]', '[>100][Green]0;[Red]0', 1234.5678],
  ['E19', '颜色段 [Red]', '[Red]0.00;[Blue]-0.00', -1234.5678],
  ['E20', '日期+星期 aaaa', 'yyyy"年"m"月"d"日" aaaa', 1234.5678],
  ['E21', '往返读取', '#,##0.000', 1234.5678],
  ['E22', '区域性代码 [$-411]', '[$-411]#,##0.00', 1234.5678],
  ['E23', '转义符 !', '0!.0,"万"', 1234.5678],
  ['E23-x', '转义符 \\（Excel 标准写法）', '0\\.0,"万"', 1234.5678],
  ['E24', '会计对齐', '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 1234.5678],
  ['E25', '内嵌公式串', 'TEXT(1,"0.00")', 1234.5678],
];

function serve() {
  return new Promise((resolve) => {
    const srv = http.createServer((req, res) => {
      const u = decodeURIComponent(req.url.split('?')[0]);
      const p = path.join(ROOT, u === '/' ? 'index.html' : u);
      fs.readFile(p, (e, b) => {
        if (e) { res.writeHead(404).end('nf'); return; }
        res.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' });
        res.end(b);
      });
    });
    srv.listen(PORT, '127.0.0.1', () => resolve(srv));
  });
}

/** 造一段在对照台里跑的代码：逐条设值 + 设格式 + 读显示文本 */
function buildCode(side) {
  const lines = ['var out = [];'];
  CASES.forEach(([id, , fmt, val], i) => {
    const r = 40 + i;
    if (side === 'sjs') {
      lines.push(`try { var s = spread.getActiveSheet(); s.setValue(${r}, 0, ${val}); s.setFormatter(${r}, 0, ${JSON.stringify(fmt)}); out.push(${JSON.stringify(id)} + '|' + s.getText(${r}, 0)); } catch (e) { out.push(${JSON.stringify(id)} + '|ERR:' + e.message); }`);
    } else {
      lines.push(`try { var rg = sheet.getRange(${r}, 0); rg.setValue(${val}); rg.setNumberFormat(${JSON.stringify(fmt)}); out.push(${JSON.stringify(id)} + '|' + rg.getDisplayValue()); } catch (e) { out.push(${JSON.stringify(id)} + '|ERR:' + e.message); }`);
    }
  });
  lines.push('return out.join("\\u0001");');
  return lines.join('\n');
}

(async () => {
  const { chromium } = require('playwright');
  const srv = await serve();
  const browser = await chromium.launch({ channel: process.env.BROWSER_CHANNEL || 'msedge', headless: true });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => console.log('[pageerror]', e.message));
  await page.goto(`http://127.0.0.1:${PORT}/index.html`, { waitUntil: 'load' });
  await page.waitForFunction(() => window.__cmpRun && window.__wb, null, { timeout: 90000 });
  await page.waitForTimeout(1000);

  const run = async (side) => {
    const res = await page.evaluate(
      ({ side, code }) => window.__cmpRun(side, code),
      { side, code: buildCode(side) }
    );
    if (!res.ok) throw new Error(side + ' 执行失败：' + res.err);
    const map = {};
    String(res.value).split('\u0001').forEach((seg) => {
      const i = seg.indexOf('|');
      map[seg.slice(0, i)] = seg.slice(i + 1);
    });
    return { map, ms: res.ms };
  };

  const sjs = await run('sjs');
  const uni = await run('uni');

  console.log('id      | 说明                        | SpreadJS 显示                | Univer 显示');
  console.log('--------|-----------------------------|------------------------------|------------------------------');
  const rows = CASES.map(([id, name, fmt, val]) => {
    const a = sjs.map[id];
    const b = uni.map[id];
    console.log(
      id.padEnd(7) + ' | ' + name.padEnd(27) + ' | ' + String(a).padEnd(28) + ' | ' + String(b)
    );
    return { id, name, fmt, value: val, sjs: a, uni: b };
  });
  console.log('\nSpreadJS 耗时 ' + sjs.ms + ' ms / Univer 耗时 ' + uni.ms + ' ms');

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify({ at: new Date().toISOString(), rows }, null, 1), 'utf8');
  console.log('写出 ' + OUT);

  await browser.close();
  srv.close();
})();
