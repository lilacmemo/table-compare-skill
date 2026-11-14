// 由 tools/out/api-ws.json 生成《工作表-工作簿 API 对比报告》的分域判定表
//   node tools/gen_report_ws.js > tools/out/ws-tables.md
const r = require('./out/api-ws.json');
const items = r.items || [];

// 线上 Pro 定向补测结论（poc-compare/probe_pro_outline.js、_diag_pro_outline2.js、
// probe_pro_caps.js 实测）：本地 OSS bundle 无 @univerjs-pro/sheets-outline，
// 故 E1~E4 在本地跑批里记为抛错，这里按 Pro 实测结果改判，来源标 [Pro]。
const PRO = {
  E1: { api: 'FWorksheet.addRowOutline(start, num)', got: 'outline 对象 {id,axis,start,end,collapsed}' },
  E2: { api: 'FWorksheet.addColumnOutline(start, num)', got: '同上，axis=column' },
  E3: { api: 'FWorksheet.removeDimensionOutline(id)', got: '3→2' },
  E4: { api: 'FWorksheet.setDimensionOutlineCollapsed(id, bool)', got: 'false→true→false' },
};
// E5~E8 亦已线上复核：Pro 的 FWorksheet 上 outline 相关方法仅 6 个
// （add/remove/setCollapsed/clear/get），outline 对象字段仅
// {id,unitId,subUnitId,axis,start,end,collapsed}，无 direction/level；
// univerAPI.Event 共 52 项，无任何 outline/group/dimension 事件 → 确认无 API。

const GROUPS = {
  A: '工作表管理（增删改名/复制/移动/标签）',
  B: '工作表可见性',
  C: '工作表保护与权限',
  D: '冻结窗格',
  E: '行列分组（大纲）',
  F: '行列操作与尺寸',
  G: '视图与显示',
  H: '查找与替换',
  I: '序列化与文档属性',
  J: '事件 / 命令 / 撤销',
  K: '剪贴板',
  L: '数据区域与选区',
};

const mark = (s, id) => {
  if (!s) return '—';
  if (s.noApi) return `无 API`;
  if (s.setOk === false) {
    if (PRO[id]) return `支持 [Pro]`;
    return `未跑通`;
  }
  if (s.verdict === 'FAIL') return `未达成`;
  if (s.got === 'ENTRY_OK') return `支持（入口）`;
  return `支持`;
};

const cell = (s, id) => {
  const m = mark(s, id);
  if (m === '无 API') return '无 API';
  if (m === '支持 [Pro]') return '支持 [Pro]';
  return m;
};

const apiCell = (s, id, side) => {
  if (side === 'uni' && PRO[id]) return '`' + PRO[id].api + '`';
  if (!s) return '—';
  if (s.noApi) return s.note ? s.note : '—';
  return '`' + (s.api || '') + '`';
};

const tally = (side) => {
  let pass = 0, noApi = 0, err = 0, fail = 0;
  for (const i of items) {
    const s = i[side];
    if (!s) continue;
    if (s.noApi) noApi++;
    else if (s.setOk === false) { if (PRO[i.id]) pass++; else err++; }
    else if (s.verdict === 'FAIL') fail++;
    else pass++;
  }
  return { pass, noApi, err, fail, total: items.length };
};

const out = [];
out.push('## 2. 判定表');
out.push('');
out.push('标记口径：**支持**＝有公开 API 且本次实测调用成功；**支持 [Pro]**＝本地 OSS 包无此插件，线上 Pro 环境实测通过；**无 API**＝文档与运行时方法面上均不存在；**未达成**＝入口存在但读回值不符合语义（仅 strict 项会出此判定）；**未跑通**＝调用了但抛错且无 Pro 备用证据。');
out.push('');
out.push('### 2.1 按组统计');
out.push('');
const sjs = tally('sjs'), uni = tally('uni');
out.push('| 分组 | 项数 | SpreadJS 支持 | Univer 支持 |');
out.push('|---|---|---|---|');
for (const g of Object.keys(GROUPS)) {
  const gi = items.filter((i) => i.g === g);
  const c = (side) => gi.filter((i) => { const s = i[side]; if (!s) return false; if (s.noApi) return false; if (s.setOk === false) return !!PRO[i.id]; return s.verdict !== 'FAIL'; }).length;
  out.push(`| ${g} ${GROUPS[g]} | ${gi.length} | ${c('sjs')} | ${c('uni')} |`);
}
out.push(`| **合计** | **${items.length}** | **${sjs.pass}** | **${uni.pass}** |`);
out.push('');

for (const g of Object.keys(GROUPS)) {
  const gi = items.filter((i) => i.g === g);
  out.push(`### ${g}. ${GROUPS[g]}（${gi.length} 项）`);
  out.push('');
  out.push('| ID | 能力项 | Excel 入口 | SpreadJS | Univer |');
  out.push('|---|---|---|---|---|');
  for (const i of gi) {
    out.push(`| ${i.id} | ${i.name} | ${i.excel || '—'} | ${cell(i.sjs, i.id)} | ${cell(i.uni, i.id)} |`);
  }
  out.push('');
  out.push('<details><summary>API 对照（点开）</summary>');
  out.push('');
  out.push('| ID | SpreadJS API | Univer API | 备注 |');
  out.push('|---|---|---|---|');
  for (const i of gi) {
    const note = [];
    if (i.sjs && i.sjs.note) note.push('SJS: ' + i.sjs.note);
    if (i.uni && i.uni.note) note.push('Uni: ' + i.uni.note);
    if (PRO[i.id]) note.push('Pro 读回: ' + PRO[i.id].got);
    out.push(`| ${i.id} | ${apiCell(i.sjs, i.id, 'sjs')} | ${apiCell(i.uni, i.id, 'uni')} | ${note.join('；').replace(/\|/g, '/') || '—'} |`);
  }
  out.push('');
  out.push('</details>');
  out.push('');
}

console.log(out.join('\n'));
