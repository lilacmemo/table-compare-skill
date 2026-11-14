// 表 Table + 排序 + 筛选 + 数据验证：双侧用例库
//   与 index.html 内联的格式用例合并使用：CASES.push(...window.CASES_TABLE)
//   约定：每条用例返回一句「可核对的结果」，抛错即记 FAIL；
//         excel 字段写 Excel 原生的行为，作为判定基线。
//   排序类用例在 Univer 侧一律先做能力检测再抛错 —— 这是**实测结论**（本版本无任何
//   区域排序 Facade，只有表内 set-table-sort-state 命令），不是漏写代码。
window.CASES_TABLE = [

  // ==================== A 表结构 ====================
  {
    id: 'T-A1',
    title: 'A1 创建表（区域转表，含标题行）',
    excel: 'Excel 基线：选中 A1:C6 → 插入 → 表格，5 行数据自动成为表体，标题行加粗带筛选按钮',
    sjs: `sheet.setArray(0, 0, [['产品','数量','单价'],['苹果',120,3.5],['香蕉',80,2.0],['橙子',200,4.2],['葡萄',60,8.5],['西瓜',30,12]]);
var t = sheet.tables.add('TblA1', 0, 0, 6, 3);
var r = t.range();
return '表数=' + sheet.tables.all().length + ' 名=' + t.name() + ' 区域=' + r.row + ',' + r.col + ' ' + r.rowCount + 'x' + r.colCount;`,
    uni: `sheet.getRange('A1:C6').setValues([['产品','数量','单价'],['苹果',120,3.5],['香蕉',80,2.0],['橙子',200,4.2],['葡萄',60,8.5],['西瓜',30,12]]);
var ok = await sheet.addTable('TblA1', sheet.getRange('A1:C6').getRange(), 'tid-a1', { tableStyleId: 'table-default-1' });
var infos = sheet.getSubTableInfos();
return 'addTable=' + ok + ' 表数=' + infos.length + ' sample=' + JSON.stringify(infos[0] || null).slice(0, 120);`
  },
  {
    id: 'T-A2',
    title: 'A2 改表名',
    excel: 'Excel 基线：表格设计 → 表名称 改为新名，结构化引用会跟着变',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var t = sheet.tables.add('OldName', 0, 0, 3, 2);
var old = t.name();
t.name('NewName');
return '旧=' + old + ' 新=' + sheet.tables.findByName('NewName').name();`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
await sheet.addTable('OldName', sheet.getRange('A1:B3').getRange(), 'tid-a2', {});
var ok = await sheet.setTableName('tid-a2', 'NewName');
return 'setTableName=' + ok + ' 信息=' + JSON.stringify(sheet.getSubTableInfos()[0] || null).slice(0, 140);`
  },
  {
    id: 'T-A3',
    title: 'A3 调整表区域（向下列扩展一行）',
    excel: 'Excel 基线：在表末行下方输入内容按 Tab，表格自动扩行（自动扩展）',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var t = sheet.tables.add('TblA3', 0, 0, 3, 2);
sheet.setValue(3, 0, 5); sheet.setValue(3, 1, 6);
// 注意：resize 在 TableManager 上（sheet.tables.resize(table, range)），不是 table.resize
sheet.tables.resize(t, new GC.Spread.Sheets.Range(0, 0, 4, 2));
return '扩展后=' + t.range().rowCount + ' 行; A4=' + sheet.getValue(3, 0);`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
await sheet.addTable('TblA3', sheet.getRange('A1:B3').getRange(), 'tid-a3', {});
sheet.getRange('A4:B4').setValues([[5, 6]]);
var ok = await sheet.setTableRange('tid-a3', sheet.getRange('A1:B4').getRange());
return 'setTableRange=' + ok + ' 信息=' + JSON.stringify(sheet.getSubTableInfos()[0] || null).slice(0, 140);`
  },
  {
    id: 'T-A4',
    title: 'A4 删除表（移除表对象，保留数据）',
    excel: 'Excel 基线：表格设计 → 转换为区域：表对象消失，值/格式保留，筛选按钮消失',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
sheet.tables.add('TblA4', 0, 0, 3, 2);
// 注意：默认 remove 会连数据一起清掉；要「转普通区域」必须传 keepData
sheet.tables.remove('TblA4', GC.Spread.Sheets.Tables.TableRemoveOptions.keepData);
return '剩余表数=' + sheet.tables.all().length + ' A2仍=' + sheet.getValue(1, 0);`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
await sheet.addTable('TblA4', sheet.getRange('A1:B3').getRange(), 'tid-a4', {});
var ok = await sheet.removeTable('tid-a4');
return 'removeTable=' + ok + ' 剩余=' + sheet.getSubTableInfos().length + ' A2仍=' + sheet.getRange('A2').getValue();`
  },
  {
    id: 'T-A5',
    title: 'A5 按单元格反查所属表',
    excel: 'Excel 基线：Excel 对象模型无直接「按单元格找表」的 UI 操作，属于程序接口（ListObject 遍历）',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var t = sheet.tables.add('TblA5', 0, 0, 3, 2);
var hit = sheet.tables.find(1, 0);
return 'tables.find(1,0)=' + (hit && hit.name());`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
await sheet.addTable('TblA5', sheet.getRange('A1:B3').getRange(), 'tid-a5', {});
var info = sheet.getTableByCell(1, 0);
return 'getTableByCell=' + JSON.stringify(info || null).slice(0, 160);`
  },
  {
    id: 'T-A6',
    title: 'A6 表头行显示 / 隐藏',
    excel: 'Excel 基线：表格设计 → 标头行 勾选切换（隐藏后列名不在表内显示，但结构性引用仍可用）',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var t = sheet.tables.add('TblA6', 0, 0, 3, 2);
t.showHeader(false);
var off = t.showHeader();
t.showHeader(true);
return '关=' + off + ' 开=' + t.showHeader();`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
sheet.getRange('D1:E3').setValues([['A','B'],['1',2],['3',4]]);
// 必须对比默认值，否则「读回 false」可能只是字段默认值
await sheet.addTable('TblDefault', sheet.getRange('A1:B3').getRange(), 'tid-a6a', {});
await sheet.addTable('TblNoHead', sheet.getRange('D1:E3').getRange(), 'tid-a6b', { showHeader: false });
var infos = sheet.getSubTableInfos();
var pick = function (id) { for (var i = 0; i < infos.length; i++) if (infos[i].id === id) return infos[i]; return {}; };
var dft = pick('tid-a6a').showHeader;
var off = pick('tid-a6b').showHeader;
if (dft === true && off === false) return '默认 showHeader=' + dft + '，传 false 后=' + off + ' → 开关生效';
throw new Error('默认=' + dft + '，传 false 后=' + off + ' → 该字段不随参数变化（ITableOptions.showHeader 标 @todo not support yet）');`
  },
  {
    id: 'T-A7',
    title: 'A7 汇总行显示 / 隐藏',
    excel: 'Excel 基线：表格设计 → 汇总行 勾选；最后一行出现「总计」并可切 7 种汇总方式',
    sjs: `sheet.setArray(0, 0, [['区域','金额'],['华东',5200],['华南',4100],['华北',3800]]);
var t = sheet.tables.add('TblA7', 0, 0, 4, 2);
t.showFooter(true);
t.setColumnValue(0, '总计');
t.setColumnFormula(1, '=SUBTOTAL(109,[金额])');
return 'footer=' + t.showFooter() + ' 汇总值=' + sheet.getValue(t.range().row + t.range().rowCount - 1, 1);`,
    uni: `var variants = [['hasTotalRow', { hasTotalRow: true }], ['showFooter', { showFooter: true }]];
var obs = [];
for (var i = 0; i < variants.length; i++) {
  var top = i * 8;
  sheet.getRange(top + 1, 0, 4, 2).setValues([['区域','金额'],['华东',5200],['华南',4100],['华北',3800]]);
  var tid = 'tid-a7-' + i;
  await sheet.addTable('TblA7' + i, sheet.getRange(top + 1, 0, 4, 2).getRange(), tid, variants[i][1]);
  await new Promise(function (r) { setTimeout(r, 200); });
  var info = sheet.getSubTableInfos().filter(function (x) { return x.id === tid; })[0] || {};
  obs.push(variants[i][0] + '→range=' + JSON.stringify(info.range) + ' 字段=[' + Object.keys(info).join(',') + ']');
}
// A6 已经证明 showHeader 的 @todo 注释是过时的，所以这里两个字段都试，不凭注释下结论
throw new Error('ITableOptions 里与汇总行相关的两个字段都试过，表区域都没有多出一行：' + obs.join(' || '));`
  },
  {
    id: 'T-A8',
    title: 'A8 汇总行下拉列表（点单元格选汇总方式）',
    excel: 'Excel 基线：汇总行单元格有下拉箭头，可选 无/平均/计数/最大/最小/求和',
    sjs: `sheet.setArray(0, 0, [['区域','金额'],['华东',5200],['华南',4100]]);
var t = sheet.tables.add('TblA8', 0, 0, 3, 2);
t.showFooter(true);
t.useFooterDropDownList(true);
return '下拉列表=' + t.useFooterDropDownList();`,
    uni: `throw new Error('汇总行本身未实现（hasTotalRow @todo），更无汇总方式下拉');`
  },
  {
    id: 'T-A9',
    title: 'A9 表内计算列（列公式自动填充整列）',
    excel: 'Excel 基线：表内某列输入一次公式，整列自动填充，且新增行自动继承',
    sjs: `sheet.setArray(0, 0, [['数量','单价','金额'],[10,5,null],[20,6,null],[30,7,null]]);
var t = sheet.tables.add('TblA9', 0, 0, 4, 3);
// 列级 API 是 setColumnDataFormula（数据列公式）；setColumnFormula 是页脚汇总公式。
// 两个名字只差一个词，查文档时极易看漏——SpreadJS 有列级计算列 API。
t.setColumnDataFormula(2, '=[@数量]*[@单价]');
var before = [1,2,3].map(function(r){ return sheet.getValue(r,2); }).join(',');
// 扩表并补两行，验证新增行是否自动继承列公式（Excel 行为：继承）
sheet.tables.resize(t, new GC.Spread.Sheets.Range(0,0,6,3));
sheet.setArray(4, 0, [[40,8,null],[50,9,null]]);
var after = [4,5].map(function(r){ return sheet.getValue(r,2); }).join(',');
var f = sheet.getFormula(1,2);
return 'setColumnDataFormula 填充=' + before + '；新增 2 行继承=' + after + '；单元格公式=' + f;`,
    uni: `sheet.getRange('A1:C4').setValues([['数量','单价','金额'],[10,5,null],[20,6,null],[30,7,null]]);
await sheet.addTable('TblA9', sheet.getRange('A1:C4').getRange(), 'tid-a9', {});
var info = sheet.getSubTableInfos()[0] || {};
var cols = info.subTableColumns || info.columns || [];
var colId = cols[2] && (cols[2].id || cols[2].columnId);
// 计算列要用 [@列] 本行引用（整列引用 [列] 在计算列里不成立）
var ok = await sheet.setTableColumnFormula('tid-a9', colId, '=[@数量]*[@单价]');
await new Promise(function(r){ setTimeout(r, 300); });
return 'setTableColumnFormula(列级 API)=' + ok + ' 金额=' + [1,2,3].map(function(r){ return sheet.getRange(r,2).getValue(); }).join(',');`
  },

  // ==================== B 表样式 ====================
  {
    id: 'T-B1',
    title: 'B1 应用内置表样式主题',
    excel: 'Excel 基线：表格样式库，Light/Medium/Dark 各 28/28/11 种',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var T = GC.Spread.Sheets.Tables.TableThemes;
var t = sheet.tables.add('TblB1', 0, 0, 3, 2, T.dark1);
return '主题已应用；可用主题数=' + Object.keys(T).length + '（含 light1-11/medium1-9/dark1-6）';`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
var ok = await sheet.addTable('TblB1', sheet.getRange('A1:B3').getRange(), 'tid-b1', { tableStyleId: 'table-default-5' });
return 'addTable(style=table-default-5)=' + ok;`
  },
  {
    id: 'T-B2',
    title: 'B2 交替行（bandRows）',
    excel: 'Excel 基线：表格样式选项 → 镶边行 勾选，奇数行浅底',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4],['5',6]]);
var t = sheet.tables.add('TblB2', 0, 0, 4, 2);
t.bandRows(false); var off = t.bandRows();
t.bandRows(true);
return '关=' + off + ' 开=' + t.bandRows();`,
    uni: `sheet.getRange('A1:B4').setValues([['A','B'],['1',2],['3',4],['5',6]]);
var ok = await sheet.addTable('TblB2', sheet.getRange('A1:B4').getRange(), 'tid-b2', { tableStyleId: 'table-default-2' });
var info = sheet.getSubTableInfos()[0] || {};
throw new Error('addTable=' + ok + '，但 ITableOptions 无 bandRows 字段；subTable 字段=[' + Object.keys(info).join(',') + '] → 无交替行开关');`
  },
  {
    id: 'T-B3',
    title: 'B3 交替列（bandColumns）',
    excel: 'Excel 基线：表格样式选项 → 镶边列 勾选，偶数列浅底',
    sjs: `sheet.setArray(0, 0, [['A','B','C'],['1',2,3],['4',5,6]]);
var t = sheet.tables.add('TblB3', 0, 0, 3, 3);
t.bandColumns(true);
return 'bandColumns=' + t.bandColumns();`,
    uni: `sheet.getRange('A1:C3').setValues([['A','B','C'],[1,2,3],[4,5,6]]);
await sheet.addTable('TblB3', sheet.getRange('A1:C3').getRange(), 'tid-b3', {});
var info = sheet.getSubTableInfos()[0] || {};
throw new Error('ITableOptions 无 bandColumns 字段；subTable 字段=[' + Object.keys(info).join(',') + '] → 无交替列开关');`
  },
  {
    id: 'T-B4',
    title: 'B4 首列高亮',
    excel: 'Excel 基线：表格样式选项 → 第一列 勾选（整列加粗/深底）',
    sjs: `sheet.setArray(0, 0, [['姓名','部门'],['张三','研发'],['李四','销售']]);
var t = sheet.tables.add('TblB4', 0, 0, 3, 2);
t.highlightFirstColumn(true);
return 'highlightFirstColumn=' + t.highlightFirstColumn();`,
    uni: `sheet.getRange('A1:B3').setValues([['姓名','部门'],['张三','研发'],['李四','销售']]);
await sheet.addTable('TblB4', sheet.getRange('A1:B3').getRange(), 'tid-b4', {});
var info = sheet.getSubTableInfos()[0] || {};
throw new Error('ITableOptions 无 showFirstColumn 字段；subTable 字段=[' + Object.keys(info).join(',') + '] → 无首列高亮开关');`
  },
  {
    id: 'T-B5',
    title: 'B5 末列高亮',
    excel: 'Excel 基线：表格样式选项 → 最后一列 勾选',
    sjs: `sheet.setArray(0, 0, [['A','B','C'],['1',2,3],['4',5,6]]);
var t = sheet.tables.add('TblB5', 0, 0, 3, 3);
t.highlightLastColumn(true);
return 'highlightLastColumn=' + t.highlightLastColumn();`,
    uni: `sheet.getRange('A1:C3').setValues([['A','B','C'],[1,2,3],[4,5,6]]);
await sheet.addTable('TblB5', sheet.getRange('A1:C3').getRange(), 'tid-b5', {});
var info = sheet.getSubTableInfos()[0] || {};
throw new Error('ITableOptions 无 showLastColumn 字段；subTable 字段=[' + Object.keys(info).join(',') + '] → 无末列高亮开关');`
  },
  {
    id: 'T-B6',
    title: 'B6 表内筛选按钮显隐',
    excel: 'Excel 基线：表格设计 → 筛选按钮 勾选，控制列头下拉箭头',
    sjs: `sheet.setArray(0, 0, [['A','B'],['1',2],['3',4]]);
var t = sheet.tables.add('TblB6', 0, 0, 3, 2);
t.filterButtonVisible(true);
return 'filterButtonVisible=' + t.filterButtonVisible();`,
    uni: `sheet.getRange('A1:B3').setValues([['A','B'],['1',2],['3',4]]);
await sheet.addTable('TblB6', sheet.getRange('A1:B3').getRange(), 'tid-b6', { showAutoFilter: true });
var ok = await sheet.setTableFilterButtons('tid-b6', { showFilterButton: false });
return 'setTableFilterButtons=' + ok;`
  },

  // ==================== C 结构化引用 ====================
  {
    id: 'T-C1',
    title: 'C1 整列引用 表[列名]',
    excel: 'Excel 基线：=SUM(TblSales[数量]) → 150',
    sjs: `sheet.setArray(0, 0, [['产品','数量'],['甲',10],['乙',20],['丙',30],['丁',40],['戊',50]]);
sheet.tables.add('TblC1', 0, 0, 6, 2);
sheet.setFormula(0, 3, '=SUM(TblC1[数量])');
return 'SUM(TblC1[数量])=' + sheet.getValue(0, 3);`,
    uni: `sheet.getRange('A1:B6').setValues([['产品','数量'],['甲',10],['乙',20],['丙',30],['丁',40],['戊',50]]);
await sheet.addTable('TblC1', sheet.getRange('A1:B6').getRange(), 'tid-c1', {});
sheet.getRange('D1').setFormula('=SUM(TblC1[数量])');
await new Promise(function(r){ setTimeout(r, 300); });
return 'SUM(TblC1[数量])=' + sheet.getRange('D1').getValue() + ' 显示=' + sheet.getRange('D1').getDisplayValue();`
  },
  {
    id: 'T-C2',
    title: 'C2 [#All] 全区域引用',
    excel: 'Excel 基线：=COUNTA(TblSales[#All]) → 24（6 行 × 4 列非空）',
    sjs: `sheet.setArray(0, 0, [['产品','数量'],['甲',10],['乙',20],['丙',30]]);
sheet.tables.add('TblC2', 0, 0, 4, 2);
sheet.setFormula(0, 3, '=COUNTA(TblC2[#All])');
return 'COUNTA([#All])=' + sheet.getValue(0, 3) + '（Excel 基线 8=4行×2列）';`,
    uni: `sheet.getRange('A1:B4').setValues([['产品','数量'],['甲',10],['乙',20],['丙',30]]);
await sheet.addTable('TblC2', sheet.getRange('A1:B4').getRange(), 'tid-c2', {});
sheet.getRange('D1').setFormula('=COUNTA(TblC2[#All])');
await new Promise(function(r){ setTimeout(r, 300); });
return 'COUNTA([#All])=' + sheet.getRange('D1').getValue() + '（Excel 基线 8）';`
  },
  {
    id: 'T-C3',
    title: 'C3 [#Headers] 标题行引用',
    excel: 'Excel 基线：=COUNTA(TblSales[#Headers]) → 列数 4',
    sjs: `sheet.setArray(0, 0, [['产品','数量'],['甲',10],['乙',20]]);
sheet.tables.add('TblC3', 0, 0, 3, 2);
sheet.setFormula(0, 3, '=COUNTA(TblC3[#Headers])');
return 'COUNTA([#Headers])=' + sheet.getValue(0, 3) + '（Excel 基线 2）';`,
    uni: `sheet.getRange('A1:B3').setValues([['产品','数量'],['甲',10],['乙',20]]);
await sheet.addTable('TblC3', sheet.getRange('A1:B3').getRange(), 'tid-c3', {});
sheet.getRange('D1').setFormula('=COUNTA(TblC3[#Headers])');
await new Promise(function(r){ setTimeout(r, 300); });
return 'COUNTA([#Headers])=' + sheet.getRange('D1').getValue() + '（Excel 基线 2）';`
  },
  {
    id: 'T-C4',
    title: 'C4 [@本行] 表内计算列',
    excel: 'Excel 基线：金额列 =[@数量]*[@单价]，整列自动填充',
    sjs: `sheet.setArray(0, 0, [['数量','单价','金额'],[10,5,null],[20,6,null],[30,7,null]]);
var t = sheet.tables.add('TblC4', 0, 0, 4, 3);
// [@] 是「本行」结构化引用；用列级 API 一次写整列，值能算出来即语法被识别
t.setColumnDataFormula(2, '=[@数量]*[@单价]');
return '金额=' + [1,2,3].map(function(r){ return sheet.getValue(r,2); }).join(',') + '；表内存储公式=' + sheet.getFormula(1,2);`,
    uni: `sheet.getRange('A1:C4').setValues([['数量','单价','金额'],[10,5,null],[20,6,null],[30,7,null]]);
await sheet.addTable('TblC4', sheet.getRange('A1:C4').getRange(), 'tid-c4', {});
var colId = (function(){ var i = sheet.getSubTableInfos()[0] || {}; var cols = i.subTableColumns || i.columns || []; return cols[2] && (cols[2].id || cols[2].columnId); })();
var ok = await sheet.setTableColumnFormula('tid-c4', colId, '=[@数量]*[@单价]');
await new Promise(function(r){ setTimeout(r, 300); });
return 'setTableColumnFormula=' + ok + ' 值=' + [1,2,3].map(function(r){ return sheet.getRange(r,2).getValue(); }).join(',');`
  },
  {
    id: 'T-C5',
    title: 'C5 汇总行里的结构化引用 SUBTOTAL',
    excel: 'Excel 基线：汇总行自动写 =SUBTOTAL(109,表[金额])，随筛选变化',
    sjs: `sheet.setArray(0, 0, [['区域','金额'],['华东',5200],['华南',4100],['华北',3800]]);
var t = sheet.tables.add('TblC5', 0, 0, 4, 2);
t.showFooter(true);
t.setColumnFormula(1, '=SUBTOTAL(109,[金额])');
var fr = t.range().row + t.range().rowCount - 1;
return '汇总值=' + sheet.getValue(fr, 1) + '（Excel 基线 13100）';`,
    uni: `sheet.getRange('A1:B4').setValues([['区域','金额'],['华东',5200],['华南',4100],['华北',3800]]);
await sheet.addTable('TblC5', sheet.getRange('A1:B4').getRange(), 'tid-c5', {});
sheet.getRange('D1').setFormula('=SUBTOTAL(109,TblC5[金额])');
await new Promise(function(r){ setTimeout(r, 300); });
var v = sheet.getRange('D1').getValue();
throw new Error('SUBTOTAL + 结构化引用本身可算（=' + v + '）；缺的是「汇总行」这个载体（ITableOptions.hasTotalRow 标 @todo）');`
  },

  // ==================== D 排序 ====================
  {
    id: 'T-D1',
    title: 'D1 单列升序',
    excel: 'Excel 基线：选金额列 → 升序 7600,8800,9500,10300,12000',
    sjs: `sheet.setArray(0, 0, [['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
var ok = sheet.sortRange(1, 0, 5, 2, true, [{ index: 1, ascending: true }]);
return '返回=' + ok + ' 升序后金额=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B6').setValues([['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
await sheet.addTable('TblD1', sheet.getRange('A1:B6').getRange(), 'tid-d1', {});
var before = [1,2,3,4,5].map(function (r) { return sheet.getRange(r, 1).getValue(); }).join(',');
// 没有 sheet.sortRange，但表内排序有一条命令；实测它只写 sortInfo 状态，不移动数据
// （sheets-table 源码里 sortInfo 只有 setSortInfo/toJSON，没有任何消费方去重排行）
var ok = await univerAPI.executeCommand('sheet.command.set-table-sort-state',
  { unitId: wb.getId(), tableId: 'tid-d1', sortInfo: { columnIndex: 1, sortState: 'asc' } });
await new Promise(function (r) { setTimeout(r, 400); });
var after = [1,2,3,4,5].map(function (r) { return sheet.getRange(r, 1).getValue(); }).join(',');
throw new Error('无 sheet.sortRange；表内排序命令返回 ' + ok + '，但数据未动：前=' + before + ' 后=' + after);`
  },
  {
    id: 'T-D2',
    title: 'D2 单列降序',
    excel: 'Excel 基线：金额降序 12000,10300,9500,8800,7600',
    sjs: `sheet.setArray(0, 0, [['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
sheet.sortRange(1, 0, 5, 2, true, [{ index: 1, ascending: false }]);
return '降序后金额=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B6').setValues([['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
if (typeof sheet.sortRange !== 'function') throw new Error('无 sheet.sortRange');`
  },
  {
    id: 'T-D3',
    title: 'D3 多列排序（部门↑ 金额↓）',
    excel: 'Excel 基线：先按部门 A→Z，同部门内按金额 高→低',
    sjs: `sheet.setArray(0, 0, [['姓名','部门','金额'],['张三','研发',8800],['李四','销售',12000],['王五','研发',9500],['赵六','销售',7600],['钱七','财务',10300]]);
var ok = sheet.sortRange(1, 0, 5, 3, true, [{ index: 1, ascending: true }, { index: 2, ascending: false }]);
return '返回=' + ok + ' 结果=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1) + '/' + sheet.getValue(r,2); }).join(' ');`,
    uni: `sheet.getRange('A1:C6').setValues([['姓名','部门','金额'],['张三','研发',8800],['李四','销售',12000],['王五','研发',9500],['赵六','销售',7600],['钱七','财务',10300]]);
if (typeof sheet.sortRange !== 'function') throw new Error('无多列排序入口');`
  },
  {
    id: 'T-D4',
    title: 'D4 自定义序列排序（compareFunction）',
    excel: 'Excel 基线：先定义自定义列表「高,中,低」，再按该序列排序',
    sjs: `sheet.setArray(0, 0, [['项目','等级'],['a','中'],['b','低'],['c','高'],['d','中'],['e','高']]);
var ORDER = ['高','中','低'];
function cmp(v1, v2) { return ORDER.indexOf(v1) - ORDER.indexOf(v2); }
var ok = sheet.sortRange(1, 0, 5, 2, true, [{ index: 1, ascending: true, compareFunction: cmp }]);
return '返回=' + ok + ' 结果=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B6').setValues([['项目','等级'],['a','中'],['b','低'],['c','高'],['d','中'],['e','高']]);
if (typeof sheet.sortRange !== 'function') throw new Error('无排序入口，自定义序列更无从谈起');`
  },
  {
    id: 'T-D5',
    title: 'D5 按单元格背景色排序',
    excel: 'Excel 基线：排序 → 按单元格颜色 红→黄→绿（Excel 独有）',
    sjs: `sheet.setArray(0, 0, [['任务','优先级','色'],['甲','绿','#C6EFCE'],['乙','红','#FFC7CE'],['丙','黄','#FFEB9C'],['丁','红','#FFC7CE'],['戊','绿','#C6EFCE']]);
for (var r = 1; r <= 5; r++) {
  var st = new GC.Spread.Sheets.Style();
  st.backColor = sheet.getValue(r, 2);
  sheet.setStyle(r, 1, st);
}
sheet.sortRange(1, 0, 5, 3, true, [{ index: 1, backColor: '#FFC7CE', order: 'top' }]);
return '红色置顶后优先级=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:C6').setValues([['任务','优先级','色'],['甲','绿','#C6EFCE'],['乙','红','#FFC7CE'],['丙','黄','#FFEB9C'],['丁','红','#FFC7CE'],['戊','绿','#C6EFCE']]);
if (typeof sheet.sortRange !== 'function') throw new Error('无排序入口，按颜色排序不支持');`
  },
  {
    id: 'T-D6',
    title: 'D6 按字体颜色排序',
    excel: 'Excel 基线：排序 → 按字体颜色',
    sjs: `sheet.setArray(0, 0, [['任务','状态'],['甲','绿'],['乙','红'],['丙','黄'],['丁','红']]);
var COL = { '绿': '#548235', '红': '#C65911', '黄': '#BF8F00' };
for (var r = 1; r <= 4; r++) { sheet.getCell(r, 1).foreColor(COL[sheet.getValue(r, 1)]); }
sheet.sortRange(1, 0, 4, 2, true, [{ index: 1, fontColor: '#C65911', order: 'top' }]);
return '红色字体置顶后=' + [1,2,3,4].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B5').setValues([['任务','状态'],['甲','绿'],['乙','红'],['丙','黄'],['丁','红']]);
if (typeof sheet.sortRange !== 'function') throw new Error('无排序入口，按字体颜色排序不支持');`
  },
  {
    id: 'T-D7',
    title: 'D7 排序时忽略隐藏行',
    excel: 'Excel 基线：筛选后再排序，隐藏行默认不参与（排序范围只作用于可见行）',
    sjs: `sheet.setArray(0, 0, [['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
sheet.setRowVisible(2, false);
var ok = sheet.sortRange(1, 0, 5, 2, true, [{ index: 1, ascending: true }], { ignoreHidden: true });
sheet.setRowVisible(2, true);
return '返回=' + ok + ' 结果=' + [1,2,3,4,5].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B6').setValues([['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600],['钱七',10300]]);
if (typeof sheet.sortRange !== 'function') throw new Error('无排序入口，ignoreHidden 选项不存在');`
  },
  {
    id: 'T-D8',
    title: 'D8 获取 / 复用排序状态',
    excel: 'Excel 基线：Excel 无「读回排序状态」的 UI 操作；这是程序接口（Sort.SortFields）',
    sjs: `sheet.setArray(0, 0, [['姓名','金额'],['张三',8800],['李四',12000],['王五',9500]]);
sheet.sortRange(1, 0, 3, 2, true, [{ index: 1, ascending: false }]);
var st = sheet.getSortState();
sheet.setValue(3, 1, 100); sheet.setValue(3, 0, '新');
sheet.sortRange();
return '状态=' + JSON.stringify(st && st.sortConditions || null) + ' 复用后=' + [1,2,3].map(function(r){ return sheet.getValue(r,1); }).join(',');`,
    uni: `sheet.getRange('A1:B4').setValues([['姓名','金额'],['张三',8800],['李四',12000],['王五',9500]]);
if (typeof sheet.sortRange !== 'function') throw new Error('无排序入口，排序状态无从获取');`
  },

  // ==================== E 筛选 ====================
  {
    id: 'T-E1',
    title: 'E1 创建自动筛选（区域加筛选按钮）',
    excel: 'Excel 基线：数据 → 筛选，列头出现下拉按钮',
    sjs: `sheet.setArray(0, 0, [['部门','姓名','金额'],['研发','张三',8800],['销售','李四',12000],['研发','王五',9500]]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 3));
sheet.rowFilter(rf);
return '筛选已挂载=' + !!sheet.rowFilter() + ' 按键数=' + rf.range.columnCount;`,
    uni: `sheet.getRange('A1:C4').setValues([['部门','姓名','金额'],['研发','张三',8800],['销售','李四',12000],['研发','王五',9500]]);
var f = sheet.getRange('A1:C4').createFilter();
var rng = f && f.getRange && f.getRange();
return 'createFilter=' + !!f + ' 范围=' + (rng ? (rng.startRow + ',' + rng.startColumn + ' ' + rng.rowCount + 'x' + rng.columnCount) : 'n/a');`
  },
  {
    id: 'T-E2',
    title: 'E2 文本筛选（等于某值）',
    excel: 'Excel 基线：部门筛选 = 研发，只剩研发行',
    sjs: `sheet.setArray(0, 0, [['部门','姓名'],['研发','张三'],['销售','李四'],['研发','王五']]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 2));
sheet.rowFilter(rf);
rf.addFilterItem(0, '研发');
rf.filter(0, ['研发']);
// 可见行统计：筛选后隐藏行 isVisible=false
var vis = [];
for (var r = 1; r <= 3; r++) if (sheet.getRowVisible(r)) vis.push(sheet.getValue(r, 0));
return '筛选后可见=' + vis.join(',');`,
    uni: `sheet.getRange('A1:B4').setValues([['部门','姓名'],['研发','张三'],['销售','李四'],['研发','王五']]);
var f = sheet.getRange('A1:B4').createFilter();
f.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['研发'] } });
await new Promise(function(r){ setTimeout(r, 200); });
return '隐藏行=' + JSON.stringify(f.getFilteredOutRows());`
  },
  {
    id: 'T-E3',
    title: 'E3 数字筛选（大于 1000）',
    excel: 'Excel 基线：金额 → 数字筛选 → 大于 1000',
    sjs: `sheet.setArray(0, 0, [['部门','金额'],['研发',8800],['销售',12000],['财务',300],['研发',1500]]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 5, 2));
sheet.rowFilter(rf);
rf.addFilterItem(1, '>1000');
rf.filter(1, ['>1000']);
return '已加自定义条件 金额>1000';`,
    uni: `sheet.getRange('A1:B5').setValues([['部门','金额'],['研发',8800],['销售',12000],['财务',300],['研发',1500]]);
var f = sheet.getRange('A1:B5').createFilter();
f.setColumnFilterCriteria(1, { colId: 1, filters: { customFilters: [{ operator: 'greaterThan', val: '1000' }] } });
await new Promise(function(r){ setTimeout(r, 200); });
return '隐藏行=' + JSON.stringify(f.getFilteredOutRows());`
  },
  {
    id: 'T-E4',
    title: 'E4 日期筛选（大于某日）',
    excel: 'Excel 基线：日期 → 日期筛选 → 之后',
    sjs: `sheet.setArray(0, 0, [['姓名','日期'],['张三','2025-01-05'],['李四','2025-02-11'],['王五','2025-03-20']]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 2));
sheet.rowFilter(rf);
rf.addFilterItem(1, '2025-02-11');
rf.filter(1, ['2025-02-11']);
return '日期条件已应用';`,
    uni: `sheet.getRange('A1:B4').setValues([['姓名','日期'],['张三','2025-01-05'],['李四','2025-02-11'],['王五','2025-03-20']]);
var f = sheet.getRange('A1:B4').createFilter();
f.setColumnFilterCriteria(1, { colId: 1, filters: { customFilters: [{ operator: 'greaterThan', val: '2025-02-01' }] } });
await new Promise(function(r){ setTimeout(r, 200); });
return '隐藏行=' + JSON.stringify(f.getFilteredOutRows());`
  },
  {
    id: 'T-E5',
    title: 'E5 按颜色筛选',
    excel: 'Excel 基线：筛选 → 按颜色筛选（Excel 独有）',
    sjs: `sheet.setArray(0, 0, [['任务','优先级'],['甲','绿'],['乙','红'],['丙','黄']]);
var MAP = { '绿': '#C6EFCE', '红': '#FFC7CE', '黄': '#FFEB9C' };
for (var r = 1; r <= 3; r++) {
  var st = new GC.Spread.Sheets.Style();
  st.backColor = MAP[sheet.getValue(r, 1)];
  sheet.setStyle(r, 1, st);
}
var C = GC.Spread.Sheets.ConditionalFormatting;
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 2));
sheet.rowFilter(rf);
// 颜色筛选不是独立方法，而是 ConditionType.colorCondition 的一个条件项
var cond = new C.Condition(C.ConditionType.colorCondition,
  { compareType: C.ColorCompareType.backgroundColor, expected: '#FFC7CE' });
rf.addFilterItem(1, cond);
rf.filter(1);
var vis = [];
for (var i = 1; i <= 3; i++) if (sheet.getRowVisible(i)) vis.push(sheet.getValue(i, 0));
if (vis.join(',') === '乙') return '按底色 #FFC7CE 筛选后可见=' + vis.join(',') + '（按值对照：乙）';
throw new Error('按底色筛选后可见=' + vis.join(','));`,
    uni: `sheet.getRange('A1:B4').setValues([['任务','优先级'],['甲','绿'],['乙','红'],['丙','黄']]);
sheet.getRange('B2').setBackgroundColor('#c6efce');
sheet.getRange('B3').setBackgroundColor('#ffc7ce');
sheet.getRange('B4').setBackgroundColor('#ffeb9c');
var f = sheet.getRange('A1:B4').createFilter();
// 关键：填入的是 IFilterColumn.colorFilters，颜色必须是 rgb(r,g,b) 无空格形式，
// 内部用 new ColorKit(bg.rgb).toRgbString() 比对；传 #ffc7ce 不报错但一行都匹配不上
f.setColumnFilterCriteria(1, { colId: 1, colorFilters: { cellFillColors: ['rgb(255,199,206)'] } });
await new Promise(function (r) { setTimeout(r, 200); });
var byColor = JSON.stringify(f.getFilteredOutRows());
f.setColumnFilterCriteria(1, { colId: 1, filters: { filters: ['红'] } });
await new Promise(function (r) { setTimeout(r, 200); });
var byValue = JSON.stringify(f.getFilteredOutRows());
if (byColor === byValue) return '按底色 rgb(255,199,206) 隐藏行=' + byColor + '，与按值筛选(红)一致 → 颜色筛选可用（颜色须 rgb 形式）';
throw new Error('按颜色隐藏行=' + byColor + '，按值=' + byValue + ' → 不一致');`
  },
  {
    id: 'T-E6',
    title: 'E6 前 N 项筛选（Top10）',
    excel: 'Excel 基线：数字筛选 → 前 10 项',
    sjs: `sheet.setArray(0, 0, [['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600]]);
var C = GC.Spread.Sheets.ConditionalFormatting;
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 5, 2));
sheet.rowFilter(rf);
// 正确参数出自官方 Demo：{ type: Top10ConditionType, expected: N }；
// 传 compareType 不报错也不生效（结果永远是后 N 项），踩过这个坑
var cond = new C.Condition(C.ConditionType.top10Condition,
  { type: C.Top10ConditionType.top, expected: 2 });
rf.addFilterItem(1, cond);
rf.filter(1);
var vis = [];
for (var r = 1; r <= 4; r++) if (sheet.getRowVisible(r)) vis.push(sheet.getValue(r, 1));
if (vis.join(',') === '12000,9500') return '前 2 项（按金额）=12000,9500';
throw new Error('前 2 项结果=' + vis.join(','));`,
    uni: `sheet.getRange('A1:B5').setValues([['姓名','金额'],['张三',8800],['李四',12000],['王五',9500],['赵六',7600]]);
var f = sheet.getRange('A1:B5').createFilter();
// ITableFilterItem 只有 manual / condition / color / record 四类，没有「前 N 项」；
// 下面用「金额>9000」的数值条件做近似，说明能做到什么、做不到什么
f.setColumnFilterCriteria(1, {
  colId: 1,
  customFilters: { customFilters: [{ val: 9000, operator: 'greaterThan' }] }
});
await new Promise(function (r) { setTimeout(r, 200); });
throw new Error('无前 N 项语义（ITableFilterItem 只有 manual/condition/color/record）；'
  + '用「金额>9000」近似后隐藏行=' + JSON.stringify(f.getFilteredOutRows())
  + '，与 Top2 结果相同纯属这份数据凑巧，不是 Top N');`
  },
  {
    id: 'T-E7',
    title: 'E7 多列组合筛选',
    excel: 'Excel 基线：部门=研发 且 金额>1000',
    sjs: `sheet.setArray(0, 0, [['部门','姓名','金额'],['研发','张三',8800],['销售','李四',12000],['研发','王五',500]]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 3));
sheet.rowFilter(rf);
rf.addFilterItem(0, '研发');
rf.addFilterItem(2, '>1000');
rf.filter(0, ['研发']);
rf.filter(2, ['>1000']);
return '两列条件均已加入';`,
    uni: `sheet.getRange('A1:C4').setValues([['部门','姓名','金额'],['研发','张三',8800],['销售','李四',12000],['研发','王五',500]]);
var f = sheet.getRange('A1:C4').createFilter();
f.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['研发'] } });
f.setColumnFilterCriteria(2, { colId: 2, filters: { customFilters: [{ operator: 'greaterThan', val: '1000' }] } });
await new Promise(function(r){ setTimeout(r, 200); });
return '两列条件后隐藏行=' + JSON.stringify(f.getFilteredOutRows());`
  },
  {
    id: 'T-E8',
    title: 'E8 清除 / 取消筛选',
    excel: 'Excel 基线：数据 → 清除（恢复全部行）',
    sjs: `sheet.setArray(0, 0, [['部门','姓名'],['研发','张三'],['销售','李四'],['研发','王五']]);
var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 4, 2));
sheet.rowFilter(rf);
rf.filter(0, ['研发']);
rf.reset();
return 'reset 后筛选器仍在=' + !!sheet.rowFilter();`,
    uni: `sheet.getRange('A1:B4').setValues([['部门','姓名'],['研发','张三'],['销售','李四'],['研发','王五']]);
var f = sheet.getRange('A1:B4').createFilter();
f.setColumnFilterCriteria(0, { colId: 0, filters: { filters: ['研发'] } });
await new Promise(function(r){ setTimeout(r, 200); });
f.removeFilterCriteria(0);
await new Promise(function(r){ setTimeout(r, 200); });
var removed = f.remove();
return 'removeFilterCriteria 后隐藏行=' + JSON.stringify(f.getFilteredOutRows()) + ' handler.remove=' + removed;`
  },
  {
    id: 'T-E9',
    title: 'E9 表内筛选（表自带筛选与区域筛选的关系）',
    excel: 'Excel 基线：表自带筛选按钮；同一区域不能再叠加区域筛选',
    sjs: `sheet.setArray(0, 0, [['部门','姓名'],['研发','张三'],['销售','李四']]);
var t = sheet.tables.add('TblE9', 0, 0, 3, 2);
var hasBtn = t.filterButtonVisible();
var err = '';
try {
  var rf = new GC.Spread.Sheets.Filter.HideRowFilter(new GC.Spread.Sheets.Range(0, 0, 3, 2));
  sheet.rowFilter(rf);
} catch (e) { err = e.message; }
return '表筛选按钮=' + hasBtn + ' 再挂区域筛选=' + (err ? ('抛错:' + err) : '未抛错');`,
    uni: `sheet.getRange('A1:B3').setValues([['部门','姓名'],['研发','张三'],['销售','李四']]);
await sheet.addTable('TblE9', sheet.getRange('A1:B3').getRange(), 'tid-e9', { showAutoFilter: true });
var f = sheet.getRange('A1:B3').createFilter();
return '表与区域筛选可并存？createFilter=' + !!f;`
  },

  // ==================== F 数据验证 ====================
  {
    id: 'T-F1',
    title: 'F1 序列验证（字面量）',
    excel: 'Excel 基线：数据验证 → 序列 → 甲,乙,丙（下拉选）',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createListValidator('甲,乙,丙');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 列表=' + back.value1();`,
    uni: `var rule = univerAPI.newDataValidation().requireValueInList(['甲','乙','丙'], false).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
return '类型=' + (back && back.getCriteriaType && back.getCriteriaType()) + ' 值=' + JSON.stringify(back && back.getCriteriaValues && back.getCriteriaValues());`
  },
  {
    id: 'T-F2',
    title: 'F2 序列验证（引用区域）',
    excel: 'Excel 基线：数据验证 → 序列 → =$H$5:$H$8',
    sjs: `sheet.setArray(0, 7, [['甲'],['乙'],['丙'],['丁']]);
var dv = GC.Spread.Sheets.DataValidation.createFormulaListValidator('=$H$1:$H$4');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 来源=' + back.value1();`,
    uni: `sheet.getRange('H1:H4').setValues([['甲'],['乙'],['丙'],['丁']]);
// 注意：requireValueInRange 直接吃 FRange 对象（不能传 .getRange()），也不需要 allowBlank 参数
var rule = univerAPI.newDataValidation().requireValueInRange(sheet.getRange('H1:H4')).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
return '类型=' + (back && back.getCriteriaType && back.getCriteriaType()) + ' 值=' + JSON.stringify(back && back.getCriteriaValues && back.getCriteriaValues());`
  },
  {
    id: 'T-F3',
    title: 'F3 整数范围 1~100',
    excel: 'Excel 基线：数据验证 → 整数 → 介于 1 和 100',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createNumberValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.between, '1', '100', true);
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 区间=' + back.value1() + '~' + back.value2();`,
    uni: `var rule = univerAPI.newDataValidation().requireNumberBetween(1, 100).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
var t1 = back && back.getCriteriaType && back.getCriteriaType();
// builder 没有「整数」专用方法，试 setCriteria 指到 WHOLE（注意：改完必须 setDataValidation 写回，
// 只改 getDataValidation() 拿到的副本不会落到单元格上）
var t2 = 'n/a';
try {
  var dv = sheet.getRange('A1').getDataValidation();
  dv.setCriteria(univerAPI.Enum.DataValidationType.WHOLE, [univerAPI.Enum.DataValidationOperator.BETWEEN, '1', '100']);
  sheet.getRange('A1').setDataValidation(dv);
  await new Promise(function(r){ setTimeout(r, 200); });
  t2 = sheet.getRange('A1').getDataValidation().getCriteriaType();
} catch (e) { t2 = 'ERR ' + e.message; }
return 'builder 默认类型=' + t1 + '；setCriteria(WHOLE)+写回后=' + t2;`
  },
  {
    id: 'T-F4',
    title: 'F4 小数范围 0~1',
    excel: 'Excel 基线：数据验证 → 小数 → 介于 0 和 1',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createDecimalValidator
  ? GC.Spread.Sheets.DataValidation.createDecimalValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.between, '0', '1')
  : GC.Spread.Sheets.DataValidation.createNumberValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.between, '0', '1', false);
sheet.setDataValidator(0, 0, dv);
return '类型=' + sheet.getDataValidator(0, 0).type();`,
    uni: `// Univer 的 builder 不区分「整数 / 小数」，只给 requireNumberBetween（产出 decimal）；
// 这里显式试 DECIMAL 类型，确认小数范围能建出来
var rule = univerAPI.newDataValidation().requireNumberBetween(0, 1).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
var t1 = back && back.getCriteriaType && back.getCriteriaType();
var dv = sheet.getRange('A1').getDataValidation();
dv.setCriteria(univerAPI.Enum.DataValidationType.DECIMAL, [univerAPI.Enum.DataValidationOperator.BETWEEN, '0', '1']);
await new Promise(function(r){ setTimeout(r, 200); });
return 'builder 默认=' + t1 + ' setCriteria(DECIMAL) 后=' + sheet.getRange('A1').getDataValidation().getCriteriaType();`
  },
  {
    id: 'T-F5',
    title: 'F5 日期验证（≥2025-01-01）',
    excel: 'Excel 基线：数据验证 → 日期 → 大于或等于',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createDateValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.greaterThanOrEqualsTo, new Date(2025, 0, 1));
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 起始=' + back.value1();`,
    uni: `// 注意：requireDateOnOrAfter 吃 Date 对象，传字符串会走 toLocaleDateString 报错
var rule = univerAPI.newDataValidation().requireDateOnOrAfter(new Date('2025-01-01')).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
return '类型=' + (back && back.getCriteriaType && back.getCriteriaType()) + ' 值=' + JSON.stringify(back && back.getCriteriaValues && back.getCriteriaValues());`
  },
  {
    id: 'T-F6',
    title: 'F6 文本长度 = 5',
    excel: 'Excel 基线：数据验证 → 文本长度 → 等于 5',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createTextLengthValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.equalsTo, '5');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 长度=' + back.value1();`,
    uni: `// builder 上没有文本长度方法，但 Enum.DataValidationType 里有 TEXT_LENGTH，
// 试 FDataValidation.setCriteria 能否建出来 —— 这才是「有没有公开入口」的判据
var rule = univerAPI.newDataValidation().requireValueInList(['x'], false).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var dv = sheet.getRange('A1').getDataValidation();
var builderText = typeof univerAPI.newDataValidation().setTextLength;
var after = 'n/a';
try {
  var dv = sheet.getRange('A1').getDataValidation();
  dv.setCriteria(univerAPI.Enum.DataValidationType.TEXT_LENGTH,
    [univerAPI.Enum.DataValidationOperator.EQUAL, '5']);
  sheet.getRange('A1').setDataValidation(dv);
  await new Promise(function(r){ setTimeout(r, 200); });
  after = sheet.getRange('A1').getDataValidation().getCriteriaType();
} catch (e) { after = 'ERR ' + e.message; }
if (after === 'textLength') return 'builder.setTextLength=' + builderText + '，改用 setCriteria 可建 TEXT_LENGTH（类型读回=' + after + '）';
throw new Error('builder.setTextLength=' + builderText + '，setCriteria 后类型=' + after + ' → 文本长度无可用入口');`
  },
  {
    id: 'T-F7',
    title: 'F7 自定义公式验证（只接受偶数）',
    excel: 'Excel 基线：数据验证 → 自定义 → =MOD(A1,2)=0',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createFormulaValidator('=MOD(A1,2)=0');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '类型=' + back.type() + ' 公式=' + back.value1();`,
    uni: `var rule = univerAPI.newDataValidation().requireFormulaSatisfied('=MOD(A1,2)=0').build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
return '类型=' + (back && back.getCriteriaType && back.getCriteriaType()) + ' 值=' + JSON.stringify(back && back.getCriteriaValues && back.getCriteriaValues());`
  },
  {
    id: 'T-F8',
    title: 'F8 输入提示（选中单元格弹提示）',
    excel: 'Excel 基线：数据验证 → 输入信息 → 标题+内容，选中时显示',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createListValidator('甲,乙,丙');
dv.showInputMessage(true);
dv.inputTitle('请选择');
dv.inputMessage('只能选甲/乙/丙');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '提示=' + back.showInputMessage() + ' 标题=' + back.inputTitle() + ' 内容=' + back.inputMessage();`,
    uni: `var b = univerAPI.newDataValidation().requireValueInList(['甲','乙','丙'], false);
var hasHelp = typeof b.setHelpText === 'function';
var rule = b.build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
var readBack = back && back.getHelpText ? back.getHelpText() : undefined;
if (hasHelp) return 'setHelpText 可用，读回=' + JSON.stringify(readBack);
throw new Error('builder.setHelpText=' + hasHelp + '（无输入提示入口），读回=' + JSON.stringify(readBack));`
  },
  {
    id: 'T-F9',
    title: 'F9 出错警告（拒绝非法输入）',
    excel: 'Excel 基线：数据验证 → 出错警告 → 停止/警告/信息 + 标题+内容',
    sjs: `var dv = GC.Spread.Sheets.DataValidation.createNumberValidator(GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators.between, '1', '100', true);
dv.showErrorMessage(true);
dv.errorTitle('输入有误');
dv.errorMessage('必须是 1~100 的整数');
sheet.setDataValidator(0, 0, dv);
var back = sheet.getDataValidator(0, 0);
return '出错警告=' + back.showErrorMessage() + ' 标题=' + back.errorTitle() + ' 内容=' + back.errorMessage();`,
    uni: `var b = univerAPI.newDataValidation().requireNumberBetween(1, 100);
var allowInvalidBefore = b.getAllowInvalid ? b.getAllowInvalid() : 'n/a';
var rule = b.setAllowInvalid(false).build();
sheet.getRange('A1').setDataValidation(rule);
await new Promise(function(r){ setTimeout(r, 200); });
var back = sheet.getRange('A1').getDataValidation();
return 'setAllowInvalid 存在=' + (typeof b.setAllowInvalid) + ' 之前=' + allowInvalidBefore + ' 读回=' + (back && back.getAllowInvalid ? back.getAllowInvalid() : 'n/a');`
  }
];
