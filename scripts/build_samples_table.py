# -*- coding: utf-8 -*-
"""
造「表 Table + 排序 + 筛选 + 数据验证」四组基线样本。
输出：samples/表-排序-筛选-验证.xlsx

设计原则（沿用公式/格式两轮的约定）：
  1. 以 Excel 原生能力为基准线，每块上方写一行「Excel 行为」说明，便于人工比对；
  2. 每块只放一个变量，便于判定是"结构丢失"还是"样式丢失"；
  3. 排序块放的是「Excel 排完的结果」静态快照 —— API 用例测的是"能否用代码触发排序"，
     不是"能不能显示已排好的顺序"，两者要分开记；
  4. 全部纵向单列排布 + 横向页面自适应，保证 COM 导出 PDF 一页装得下（横向多表会被截断）。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn, TableFormula
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   'samples', '表-排序-筛选-验证.xlsx')

HDR_FONT = Font(bold=True, color='FFFFFF')
HDR_FILL = PatternFill('solid', fgColor='4472C4')
NOTE_FONT = Font(italic=True, color='7F7F7F', size=9)
BLOCK_FONT = Font(bold=True, size=11, color='1F3864')


def block(ws, row, text):
    """块标题（写在该块表头上方一行）"""
    c = ws.cell(row=row, column=1, value=text)
    c.font = BLOCK_FONT


def head(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='center')


def table(ws, name, ref, style='TableStyleMedium2', banded=True,
          first=False, last=False, banded_col=False, totals=0, totals_def=None,
          columns=None, calc_cols=None):
    """建表。

    columns   : 列名列表（顺序与 ref 的列一致）。
    calc_cols : {1-based 列号: 公式} —— 计算列（Excel 的 calculatedColumnFormula）。
                公式必须写**全限定结构化引用**，如
                'TblX[[#This Row],[数量]]*TblX[[#This Row],[单价]]'，
                不能写简写 '[@数量]*[@单价]'。
                坑（实测）：单元格里直接塞 '=[@数量]*[@单价]'，Excel 打开后整列 #REF! ——
                xlsx 存储用的是全限定形式，简写只存在于 Excel 的编辑栏里。
    """
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(
        name=style, showRowStripes=banded, showColumnStripes=banded_col,
        showFirstColumn=first, showLastColumn=last,
    )
    if columns or calc_cols:
        # _initialise_columns 只会填 Column1/Column2…，列名得自己给，
        # 否则表头文字和 tableColumn 名字对不上，Excel 判定结构不一致。
        t._initialise_columns()
        for tc, nm in zip(t.tableColumns, columns or []):
            tc.name = nm
        for idx, f in (calc_cols or {}).items():
            t.tableColumns[idx - 1].calculatedColumnFormula = TableFormula(attr_text=f)
    if totals:
        t.totalsRowCount = totals
        # 坑（实测）：只写 totalsRowCount 而 tableColumn 里没有 totalsRowFunction /
        # totalsRowLabel 时，Excel 判定结构不一致，直接拒绝打开（COM 报「不能取得 Workbooks
        # 的 Open 属性」，无任何提示）。汇总行只要有内容，必须逐列声明汇总方式。
        t._initialise_columns()
        cols = []
        for i, item in enumerate(totals_def or [], 1):
            nm, fn, lbl = (tuple(item) + (None, None))[:3]
            tc = TableColumn(id=i, name=nm)
            if fn:
                tc.totalsRowFunction = fn
            if lbl:
                tc.totalsRowLabel = lbl
            cols.append(tc)
        if cols:
            t.tableColumns = cols
    ws.add_table(t)
    return t


def page(ws, landscape=True):
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


wb = Workbook()

# ============================================================
# A-表结构
# ============================================================
ws = wb.active
ws.title = 'A-表结构'
ws['A1'] = 'A 组：表结构 —— 四张表参数各不相同'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

block(ws, 2, '① 标准表 TblStd：默认样式 + 交替行 + 自动筛选按钮')
head(ws, 3, ['产品', '数量', '单价', '金额'])
rows = [('苹果', 120, 3.5, 420), ('香蕉', 80, 2.0, 160), ('橙子', 200, 4.2, 840),
        ('葡萄', 60, 8.5, 510), ('西瓜', 30, 12.0, 360)]
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        ws.cell(row=4 + i, column=1 + j, value=v)
table(ws, 'TblStd', 'A3:D8', 'TableStyleMedium2')

block(ws, 10, '② 带汇总行 TblTotal：末尾「总计」行，可切换 7 种汇总方式')
head(ws, 11, ['区域', '销售额', '订单数', '客单价'])
rows2 = [('华东', 5200, 130, 40.0), ('华南', 4100, 105, 39.0), ('华北', 3800, 95, 40.0),
         ('西南', 2600, 68, 38.2), ('东北', 1900, 52, 36.5)]
for i, r in enumerate(rows2):
    for j, v in enumerate(r):
        ws.cell(row=12 + i, column=1 + j, value=v)
ws['A17'] = '总计'
ws['B17'] = '=SUBTOTAL(109,B12:B16)'
ws['C17'] = '=SUBTOTAL(109,C12:C16)'
ws['D17'] = '=SUBTOTAL(101,D12:D16)'
for c in range(1, 5):
    ws.cell(row=17, column=c).font = Font(bold=True)
table(ws, 'TblTotal', 'A11:D17', 'TableStyleLight9', totals=1,
      totals_def=[('区域', None, '总计'), ('销售额', 'sum', None),
                  ('订单数', 'sum', None), ('客单价', 'average', None)])

block(ws, 19, '③ 首末列高亮 + 无交替行：showFirstColumn / showLastColumn = 1，showRowStripes = 0')
head(ws, 20, ['姓名', '部门', '工号', '职级'])
rows3 = [('张三', '研发', 'E001', 'P6'), ('李四', '销售', 'E002', 'P5'),
         ('王五', '财务', 'E003', 'P7'), ('赵六', '研发', 'E004', 'P6'),
         ('钱七', '客服', 'E005', 'P4')]
for i, r in enumerate(rows3):
    for j, v in enumerate(r):
        ws.cell(row=21 + i, column=1 + j, value=v)
table(ws, 'TblFirstLast', 'A20:D25', 'TableStyleMedium9', banded=False, first=True, last=True)

block(ws, 27, '④ 曾为表、已「转换为区域」：只剩数值与边框，无表对象、无筛选按钮')
head(ws, 28, ['城市', '人口', '面积', '密度'])
rows4 = [('北京', 2189, 16410, 1334), ('上海', 2487, 6340, 3923),
         ('广州', 1868, 7434, 2513), ('深圳', 1756, 1997, 8793),
         ('成都', 2094, 14335, 1461)]
for i, r in enumerate(rows4):
    for j, v in enumerate(r):
        ws.cell(row=29 + i, column=1 + j, value=v)

for col, w in zip('ABCD', [14, 12, 12, 14]):
    ws.column_dimensions[col].width = w
page(ws)

# ============================================================
# B-表样式
# ============================================================
ws = wb.create_sheet('B-表样式')
ws['A1'] = 'B 组：表样式 —— 同一份数据、四种 TableStyle 与条纹开关'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

DATA = [('A', 1, 'x'), ('B', 2, 'y'), ('C', 3, 'x'), ('D', 4, 'y'), ('E', 5, 'x')]
styles = [
    ('TblLight',  'TableStyleLight9',  True,  False, '① Light9 + 交替行（浅蓝细条纹）'),
    ('TblMedium', 'TableStyleMedium2', True,  True,  '② Medium2 + 行条纹 + 列条纹（列也深浅相间）'),
    ('TblDark',   'TableStyleDark1',   True,  False, '③ Dark1 深色皮肤（深底白字，对比最强）'),
    ('TblNoBand', 'TableStyleMedium4', False, False, '④ Medium4 关闭全部条纹（纯色底，无交替）'),
]
r = 2
for name, style, banded, banded_col, note in styles:
    block(ws, r, note)
    head(ws, r + 1, ['列1', '列2', '列3'])
    for i, d in enumerate(DATA):
        for j, v in enumerate(d):
            ws.cell(row=r + 2 + i, column=1 + j, value=v)
    table(ws, name, f'A{r+1}:C{r+6}', style, banded=banded, banded_col=banded_col)
    r += 8

for col, w in zip('ABC', [16, 10, 10]):
    ws.column_dimensions[col].width = w
page(ws)

# ============================================================
# C-结构化引用
# ============================================================
ws = wb.create_sheet('C-结构化引用')
ws['A1'] = 'C 组：结构化引用 —— 表名[列名] / [#All] / [#Headers] / [@本行]'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

block(ws, 3, '① 表 TblSales，D 列是表内计算列（Excel 的 calculatedColumnFormula，自动填充整列）')
head(ws, 4, ['产品', '数量', '单价', '金额'])
CALC_D = 'TblSales[[#This Row],[数量]]*TblSales[[#This Row],[单价]]'
for i, r0 in enumerate([('甲', 10, 5.0), ('乙', 20, 6.0), ('丙', 30, 7.0),
                        ('丁', 40, 8.0), ('戊', 50, 9.0)]):
    ws.cell(row=5 + i, column=1, value=r0[0])
    ws.cell(row=5 + i, column=2, value=r0[1])
    ws.cell(row=5 + i, column=3, value=r0[2])
    # 必须写全限定形式。写 '=[@数量]*[@单价]' 时 Excel 打开即整列 #REF!（xlsx 里没有
    # “@”简写这种存法，简写只是 Excel 编辑栏的显示形式）。
    ws.cell(row=5 + i, column=4, value='=' + CALC_D)
table(ws, 'TblSales', 'A4:D9', 'TableStyleMedium7',
      columns=['产品', '数量', '单价', '金额'], calc_cols={4: CALC_D})

block(ws, 11, '② 表外引用公式（A 列是活公式，显示的就是 Excel 实算结果；B 列是人工填的期望值）')
head(ws, 12, ['公式（活，显示结果）', '期望结果', '公式原文 ｜ 说明'], col=1)
refs = [
    ('=SUM(TblSales[数量])',              150,   '整列引用求和'),
    ('=COUNTA(TblSales[#All])',           24,    '[#All] 全区域非空单元格数（6 行 × 4 列）'),
    ('=COUNTA(TblSales[#Headers])',       4,     '[#Headers] 标题行列数'),
    ('=TblSales[[#Headers],[数量]]',      '数量', '按「列名+标题」定位单个单元格'),
    ('=AVERAGE(TblSales[单价])',          7,     '整列求平均'),
]
for i, (f, v, note) in enumerate(refs):
    # A 列放活公式：Excel 里显示的是计算结果，B 列是人工期望值，两者并排肉眼可对。
    # 公式原文没法显示在 A 列（它显示结果），所以拼进 C 列说明里，读者不用点单元格。
    ws.cell(row=13 + i, column=1, value=f)
    ws.cell(row=13 + i, column=2, value=v)
    ws.cell(row=13 + i, column=3,
            value='{} ｜ {}'.format(f.lstrip('='), note)).font = NOTE_FONT

c = ws.cell(row=20, column=1, value='判读：D 列若变成静态值、0 或 #REF!，即「表内计算列 + [#This Row] 引用」未被识别；B 列 5 项全对才算结构化引用完整。')
c.font = NOTE_FONT
for col, w in zip('ABCD', [22, 12, 62, 12]):
    ws.column_dimensions[col].width = w
page(ws)

# ============================================================
# D-排序
# ============================================================
ws = wb.create_sheet('D-排序')
ws['A1'] = 'D 组：排序 —— 左侧原始数据，右侧 Excel 排完的结果（API 用例另测「能否用代码触发排序」）'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

RAW = [('张三', '研发', 8800), ('李四', '销售', 12000), ('王五', '研发', 9500),
       ('赵六', '销售', 7600), ('钱七', '财务', 10300)]


def pair_block(row, title, left, right):
    block(ws, row, title)
    head(ws, row + 1, ['姓名', '部门', '金额'], col=1)
    head(ws, row + 1, ['姓名', '部门', '金额'], col=6)
    for i, r0 in enumerate(left):
        for j, v in enumerate(r0):
            ws.cell(row=row + 2 + i, column=1 + j, value=v)
    for i, r0 in enumerate(right):
        for j, v in enumerate(r0):
            ws.cell(row=row + 2 + i, column=6 + j, value=v)


pair_block(3, '① 单列升序：按金额 小→大', RAW, sorted(RAW, key=lambda x: x[2]))
pair_block(11, '② 多列排序：部门 A→Z，同部门内金额 高→低', RAW,
           sorted(RAW, key=lambda x: (x[1], -x[2])))

# ③ 自定义序列
block(ws, 19, '③ 自定义序列排序：高→中→低（Excel 需先定义自定义列表，否则按拼音排）')
head(ws, 20, ['项目', '等级'], col=1)
head(ws, 20, ['项目', '等级'], col=6)
LV = [('a', '中'), ('b', '低'), ('c', '高'), ('d', '中'), ('e', '高')]
for i, r0 in enumerate(LV):
    for j, v in enumerate(r0):
        ws.cell(row=21 + i, column=1 + j, value=v)
ORD = {'高': 0, '中': 1, '低': 2}
for i, r0 in enumerate(sorted(LV, key=lambda x: ORD[x[1]])):
    for j, v in enumerate(r0):
        ws.cell(row=21 + i, column=6 + j, value=v)

# ④ 按单元格颜色
block(ws, 27, '④ 按单元格颜色排序：红→黄→绿（Excel 支持「按颜色排序」，多数引擎无此能力）')
head(ws, 28, ['任务', '优先级'], col=1)
head(ws, 28, ['任务', '优先级'], col=6)
CLR = [('甲', '绿', 'C6EFCE'), ('乙', '红', 'FFC7CE'), ('丙', '黄', 'FFEB9C'),
       ('丁', '红', 'FFC7CE'), ('戊', '绿', 'C6EFCE')]
for i, (n, lv, color) in enumerate(CLR):
    ws.cell(row=29 + i, column=1, value=n)
    ws.cell(row=29 + i, column=2, value=lv).fill = PatternFill('solid', fgColor=color)
CO = {'红': 0, '黄': 1, '绿': 2}
for i, (n, lv, color) in enumerate(sorted(CLR, key=lambda x: CO[x[1]])):
    ws.cell(row=29 + i, column=6, value=n)
    ws.cell(row=29 + i, column=7, value=lv).fill = PatternFill('solid', fgColor=color)

for col, w in zip('ABCDEFG', [12, 12, 12, 10, 10, 12, 12]):
    ws.column_dimensions[col].width = w
page(ws)

# ============================================================
# E-筛选
# ============================================================
ws = wb.create_sheet('E-筛选')
ws['A1'] = 'E 组：筛选 —— 表内自动筛选按钮 + 预置条件（筛选是隐藏行，不是删除）'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

block(ws, 3, '① 表 TblFilter：带自动筛选按钮，Excel 点列头下拉可选文本/数字/日期/颜色/前N')
head(ws, 4, ['部门', '姓名', '金额', '日期', '状态'])
ROWS = [('研发', '张三', 8800, '2025-01-05', '在职'),
        ('销售', '李四', 12000, '2025-02-11', '在职'),
        ('研发', '王五', 9500, '2025-03-20', '离职'),
        ('销售', '赵六', 7600, '2025-04-02', '在职'),
        ('财务', '钱七', 10300, '2025-05-18', '在职'),
        ('研发', '孙八', 1500, '2025-06-25', '在职'),
        ('销售', '周九', 300, '2025-07-30', '离职')]
for i, r0 in enumerate(ROWS):
    for j, v in enumerate(r0):
        ws.cell(row=5 + i, column=1 + j, value=v)
table(ws, 'TblFilter', 'A4:E11', 'TableStyleMedium1')

block(ws, 13, '② 预置筛选示例：金额 > 1000（工作表级 AutoFilter，条件写在 filterColumn 里）')
head(ws, 14, ['部门', '姓名', '金额'])
for i, r0 in enumerate(ROWS):
    for j, v in enumerate(r0[:3]):
        ws.cell(row=15 + i, column=1 + j, value=v)
ws.auto_filter.ref = 'A14:C21'

block(ws, 23, '判读：导入方若只读 table 的 autoFilter 而忽略 worksheet 级 autoFilter，'
              '表现为「筛选按钮在、预置条件丢了」')
for col, w in zip('ABCDE', [10, 10, 10, 14, 10]):
    ws.column_dimensions[col].width = w
page(ws)

# ============================================================
# F-数据验证
# ============================================================
ws = wb.create_sheet('F-数据验证')
ws['A1'] = 'F 组：数据验证 —— 7 类规则 + 输入提示 + 出错警告（「圈释无效」是 UI 状态，xlsx 不存）'
ws['A1'].font = Font(bold=True, size=14, color='1F3864')

block(ws, 3, '① 下方 7 列各带一种验证规则；再下面一行是「序列来源区域」的候选值')
head(ws, 4, ['序列(字面量)', '序列(区域)', '整数1~100', '小数0~1',
             '日期≥2025/1/1', '长度=5', '偶数'])
for col in range(1, 8):
    for r0 in range(5, 13):
        ws.cell(row=r0, column=col).value = None

# 示例值（放在规则行下方，便于看效果）
SAMPLES = ['甲', '乙', 42, 0.35, '2025-03-01', 'ABCDE', 8]
for i, v in enumerate(SAMPLES):
    ws.cell(row=5, column=1 + i, value=v)

# 序列来源区域（H 列）
ws.cell(row=4, column=8, value='序列来源').font = HDR_FONT
ws.cell(row=4, column=8).fill = HDR_FILL
for i, v in enumerate(['甲', '乙', '丙', '丁']):
    ws.cell(row=5 + i, column=8, value=v)

RUNS = [
    ('A5:A12', 'list',        '"甲,乙,丙"',        None,                        '请选择',    '只能选甲/乙/丙'),
    ('B5:B12', 'list',        '$H$5:$H$8',         None,                        '请选择',    '只能从区域里选'),
    ('C5:C12', 'whole',       '1',                 '100',                       '输入 1~100 整数', '超出范围'),
    ('D5:D12', 'decimal',     '0',                 '1',                         '输入 0~1 小数',   '超出范围'),
    ('E5:E12', 'date',        '2025-01-01',        None,                        '日期需 ≥2025/1/1', '日期太早'),
    ('F5:F12', 'textLength',  '5',                 None,                        '长度必须为 5',    '长度不对'),
    ('G5:G12', 'custom',      'MOD(G5,2)=0',       None,                        '只接受偶数',      '不是偶数'),
]
for rng, typ, f1, f2, prompt, err in RUNS:
    op = None
    if typ == 'whole' or typ == 'decimal':
        op = 'between'
    elif typ == 'date':
        op = 'greaterThanOrEqual'
    elif typ == 'textLength':
        op = 'equal'
    dv = DataValidation(type=typ, formula1=f1, formula2=f2, operator=op, allow_blank=True,
                        showInputMessage=True, showErrorMessage=True,
                        promptTitle='输入提示', prompt=prompt,
                        errorTitle='输入有误', error=err)
    ws.add_data_validation(dv)
    dv.add(rng)

block(ws, 14, 'Excel 独有：①「圈释无效数据」是视图状态、不写入 xlsx；②「粘贴绕过验证」需工作表保护才拦得住')
block(ws, 15, '判读：判的是「能否用 API 建出这 7 类规则」，不是「导入后下拉框还能不能点」')
for col, w in zip('ABCDEFGH', [14, 14, 14, 12, 18, 12, 12, 12]):
    ws.column_dimensions[col].width = w
page(ws)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print('写出', OUT)
print('工作表：', wb.sheetnames)


# ============================================================
# 收尾：让 Excel 真算一遍再存
# ============================================================
# 坑（实测）：openpyxl 只写 <f>公式</f>，不写 <v>缓存值</v>。文件用 Excel 打开没问题
# （workbook.xml 带 fullCalcOnLoad="1"，Excel 会重算），但**任何读缓存值的引擎**
# （SpreadJS / Univer 导入、各种在线预览）看到的都是 0 或空。基线样本必须自带缓存值。
# 做法：把文件拷到纯 ASCII 临时路径 → 交给 Excel COM 重算并保存 → 拷回。
FINALIZE_PS1 = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'finalize_sample.ps1')
TMP = r'C:\temp\_sample_finalize.xlsx'


def finalize_with_excel():
    import shutil
    import subprocess
    if not os.path.exists(FINALIZE_PS1):
        print('未找到 finalize_sample.ps1，跳过重算')
        return False
    os.makedirs(r'C:\temp', exist_ok=True)
    shutil.copyfile(OUT, TMP)
    r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass',
                        '-File', FINALIZE_PS1], capture_output=True, text=True)
    if r.returncode != 0:
        print('Excel 重算失败：', (r.stderr or r.stdout).strip()[:400])
        return False
    shutil.copyfile(TMP, OUT)
    print('已用 Excel 重算并写回缓存值')
    return True


if '--no-excel' not in os.sys.argv:
    finalize_with_excel()
