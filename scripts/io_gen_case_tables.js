// 把 out/data/<module>-cases.json 的逐条结果生成 md 表格，追加为对应文档的「附录 A」
//   node tools/gen_case_tables.js            六个模块全刷
//   node tools/gen_case_tables.js format     只刷一个
//
// 为什么要有这一步：§4 只给分组通过率，客户/销售追问"哪个 case 没过"时无法当场回答。
// 附录 A 与 §4 同源（都来自 cases.json），不会出现两份数字打架。
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DATA = path.join(ROOT, 'out', 'data');
const DOCS = path.join(ROOT, 'docs');
const MARK = '## 附录 A';

const MAP = {
  // charts 不生成附录：01-图表.md 的 §4 已经按 A~L 分节列了全部 103 项逐条表，重复会打架
  print: {
    doc: '02-打印与PDF.md', title: '打印与 PDF', unit: '项',
    note: '本模块 PASS = 该打印配置项有公开 API 且设置后读回语义一致；Univer 侧没有 getter，读回走 `BeforeSheetPrintConfirm` 事件载荷（§4 有说明）。',
  },
  worksheet: { doc: '03-工作表与工作簿.md', title: '工作表 / 工作簿', unit: '项' },
  format: { doc: '04-格式与条件格式.md', title: '格式与条件格式', unit: '项' },
  table: {
    doc: '05-表-排序-筛选-验证.md', title: '表 / 排序 / 筛选 / 验证', unit: '项',
    note: '本模块部分用例的说明列记录了实测读回值（如表名、区域、样式 ID），便于复核，不代表 FAIL。',
  },
  formula: {
    doc: '06-公式.md', title: '公式', unit: '个函数',
    note: '**本模块 PASS 的含义与其他模块不同**：PASS = 计算结果与 Excel oracle **一致**；FAIL = 结果不一致，或 `#NAME?` 不支持（两类在 §4.1 / §4.2 分开统计，不要混成一个数）。',
  },
};

// 表格单元格里不能有竖线和换行，否则整张表错位
const cell = (s) => String(s == null ? '' : s).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ').trim();

// 探针会把整段 JSON 塞进说明（`sample={"id":"tid-a1","subUnitId":"IZiV18…`）。
// subUnitId 这类字段对读者无意义，且长度会把有用的前半句挤掉，统一收成 {…}。
const tidy = (s) => String(s == null ? '' : s)
  .replace(/\{[^{}]{25,}\}/g, '{…}')
  .replace(/\{[^{}]*$/, '{…}');

// 截断时不能停在半个 JSON / 半个括号里（否则出现「（颜色须 rgb 形」这种断头句）：
// 若尾巴有未闭合的 { 或 （，就回退到它之前再收尾。
const clip = (s, n) => {
  let t = tidy(cell(s));
  if (t.length <= n) return t;
  t = t.slice(0, n);
  const cut = Math.max(t.lastIndexOf('{'), t.lastIndexOf('（'), t.lastIndexOf('“'));
  if (cut > n * 0.45) t = t.slice(0, cut);
  // 不要停在半个英文单词/字段名上（否则出现「…showHeader,sho」）
  const tail = t.match(/[A-Za-z0-9_]{2,}$/);
  if (tail) t = t.slice(0, t.length - tail[0].length).replace(/[,，;；\s]+$/, '');
  return t.replace(/[，、；：,;\s]+$/, '') + '…';
};

function build(mod, d) {
  const meta = MAP[mod];
  const rs = d.results || [];
  const c = (k) => rs.filter((r) => r[k] === 'PASS').length;
  const n = (k, v) => rs.filter((r) => r[k] === v).length;

  const lines = [];
  lines.push('---');
  lines.push('');
  lines.push(MARK + '　逐条用例结果（全量 ' + rs.length + ' ' + meta.unit + '）');
  lines.push('');
  lines.push('> 与 §4 汇总**同源**，数据来自 `out/data/' + mod + '-cases.json`；判定口径见 §4。');
  lines.push('> 说明列主要写 `FAIL` / `NA` 的原因，`PASS` 多数记 `—`；个别模块为便于复核在 `PASS` 行也留了实测读回值，见下方模块注。过长说明截断到 180 字，全文在 JSON 里。');
  if (meta.note) {
    lines.push('>');
    lines.push('> ' + meta.note);
  }
  lines.push('');
  lines.push('**合计 ' + rs.length + ' ' + meta.unit
    + '：SpreadJS PASS ' + c('sjs') + ' / FAIL ' + n('sjs', 'FAIL') + ' / NA ' + n('sjs', 'NA')
    + '　｜　Univer Pro PASS ' + c('uni') + ' / FAIL ' + n('uni', 'FAIL') + ' / NA ' + n('uni', 'NA') + '**');
  lines.push('');
  lines.push('| # | ID | 测试项 | SpreadJS | Univer Pro | 说明 |');
  lines.push('| ---: | --- | --- | :-: | :-: | --- |');
  rs.forEach((r, i) => {
    lines.push('| ' + (i + 1) + ' | `' + cell(r.id) + '` | ' + clip(r.name, 90)
      + ' | ' + r.sjs + ' | ' + r.uni + ' | ' + clip(r.note || '—', 180) + ' |');
  });
  lines.push('');
  return lines.join('\n');
}

function run(mod) {
  const meta = MAP[mod];
  if (!meta) { console.log('未知模块', mod); return; }
  const jf = path.join(DATA, mod + '-cases.json');
  const df = path.join(DOCS, meta.doc);
  if (!fs.existsSync(jf)) { console.log('[skip]', mod, '没有 cases.json'); return; }
  if (!fs.existsSync(df)) { console.log('[skip]', mod, '没有文档', meta.doc); return; }

  const d = JSON.parse(fs.readFileSync(jf, 'utf8'));
  const table = build(mod, d);

  let doc = fs.readFileSync(df, 'utf8');
  const i = doc.indexOf(MARK);
  if (i >= 0) doc = doc.slice(0, i).replace(/\s+$/, '') + '\n\n';
  else doc = doc.replace(/\s+$/, '') + '\n\n';

  fs.writeFileSync(df, doc + table, 'utf8');
  console.log('写出', meta.doc, '附录 A', (d.results || []).length, '行');
}

const args = process.argv.slice(2).filter((a) => !a.startsWith('--'));
( args.length ? args : Object.keys(MAP) ).forEach((m) => run(m));
