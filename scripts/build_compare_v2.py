# -*- coding: utf-8 -*-
"""生成《公式对比-Excel365基准-三方-v2.xlsx》

结构：
  对比表          —— 519 行全量；E 列活公式引「Excel实测」，G/J 列先判「是否支持」，I/L 列再判「结果是否一致」
  结论            —— 全部活公式统计
  SpreadJS差异明细 —— 支持但结果与 Excel 不同的函数（人工确认用）
  Univer差异明细   —— 同上
  分族统计
  口径说明
  Excel实测        —— 数据区 + 每条用例的活公式（行间隔 12 行，容纳动态数组溢出）

判定链（与 judge_v2.py 完全一致）：
  #NAME? → 不支持（函数未实现），不比较值
  其他错误 → 产品报错
  双方都支持 → 比较值：数值 1e-9 相对容差 / 布尔 vs 0-1 记类型不同 / 文本按单元格显示文本精确比对（EXACT）
"""
import json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import judge_v2 as J

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, '..', '公式对比-Excel365基准519函数-三方-v2.xlsx')

EXCEL_VER = '16.0.20527（Microsoft 365 企业应用版 16.0）'
SJS_VER = '19.2.5'
UNIVER_VER = 'Pro 1.3.1'

rows, excluded = J.load()
data_block = json.load(open(os.path.join(HERE, 'data_block_v2.json'), encoding='utf-8'))
cases = json.load(open(os.path.join(HERE, 'cases_v2.json'), encoding='utf-8'))

GAP = 12
HDR = 4

# ---------------- 说明列 ----------------
VOLATILE = {'RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY'}
PRECISION = {'BESSELI': 'Bessel 函数精度差异（约 1e-8）', 'BESSELJ': 'Bessel 函数精度差异（约 1e-8）',
             'BESSELK': 'Bessel 函数精度差异（约 1e-8）', 'BESSELY': 'Bessel 函数精度差异（约 1e-8）',
             'T.INV.2T': 'T 分布分位数精度差异（约 3e-9）', 'TINV': 'T 分布分位数精度差异（约 3e-9）',
             'GAMMA': 'Gamma 函数精度差异（约 3e-9）'}


def note_for(r):
    n = r['name']
    out = []
    if r['sjs_verdict'] == '不支持（#NAME?）':
        out.append('SpreadJS 未实现该函数（写入后返回 #NAME?）')
    if r['univer_verdict'] == '不支持（#NAME?）':
        out.append('Univer 未实现该函数（写入后返回 #NAME?）')
    if n in VOLATILE:
        out.append('易失函数：每次重算/打开结果都变，只判定是否支持，不对比具体结果值')
    if n in PRECISION and any('数值不同' == v for v in (r['sjs_verdict'], r['univer_verdict'])):
        out.append(PRECISION[n])
    if n.startswith('IM') and any('文本不同' == v for v in (r['sjs_verdict'], r['univer_verdict'])):
        out.append('复数文本仅末位有效数字不同（Excel 显示 15 位，产品输出 17 位）')
    if n == 'BASE':
        out.append('进制转换输出字母大小写不同（Excel 大写 / 产品小写）')
    if n == 'DOLLAR':
        out.append('货币符号随区域设置（本机 Excel 显示 ¥，产品固定 $）')
    if n == 'HYPERLINK':
        out.append('Univer 不返回显示文本（返回空）')
    if n == 'VALUETOTEXT':
        out.append('Univer 返回数值 123，Excel 返回文本 "123"')
    if any('布尔/数值' in v for v in (r['sjs_verdict'], r['univer_verdict'])):
        out.append('布尔结果以 1/0 返回：参与算术可用，但 ISLOGICAL/TYPE/EXACT 等会不同')
    if r.get('arrayNote'):
        out.append(r['arrayNote'])
    return '；'.join(dict.fromkeys(out))


def category(r, key):
    v, n = r[key], r['name']
    if n in VOLATILE:
        return '随机/时间'
    if n in ('DOLLAR',):
        return '区域设置'
    if n.startswith('IM') or n in PRECISION:
        return '精度差异'
    if n.startswith('IM') and any('文本不同' == v for v in (r['sjs_verdict'], r['univer_verdict'])):
        out.append('复数文本仅末位有效数字不同（Excel 显示 15 位，产品输出 17 位）')
    if n == 'BASE':
        return '格式差异'
    if n == 'HYPERLINK':
        return '功能差异'
    if '布尔/数值' in v:
        return '类型语义'
    if v == '产品报错':
        return '报错'
    return '功能差异'


# ---------------- 样式 ----------------
HEAD_FILL = PatternFill('solid', fgColor='1F4E79')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=10)
BORDER = Border(*[Side(style='thin', color='BFBFBF')] * 4)
GREEN = PatternFill('solid', fgColor='E2EFDA')
RED = PatternFill('solid', fgColor='FCE4EC')
YELLOW = PatternFill('solid', fgColor='FFF2CC')
GREY = PatternFill('solid', fgColor='EDEDED')

wb = Workbook()
ws = wb.active
ws.title = '对比表'
ws_sum = wb.create_sheet('结论')
ws_diff_s = wb.create_sheet('SpreadJS差异明细')
ws_diff_u = wb.create_sheet('Univer差异明细')
ws_fam = wb.create_sheet('分族统计')
ws_note = wb.create_sheet('口径说明')
ws2 = wb.create_sheet('Excel实测')

# ---------------- 对比表 ----------------
HEADERS = ['#', '函数名', '函数族', '用例公式（Excel 已验证可出值）',
           'Excel 365 结果（活公式）', 'Excel 结果文本',
           'SpreadJS\n是否支持', 'SpreadJS 结果（快照）', 'SpreadJS 判定',
           'Univer\n是否支持', 'Univer 结果（快照）', 'Univer 判定', '说明']

ws['A1'] = 'Excel 365 原生函数 × SpreadJS / Univer Pro 三方对比（v2：用例先在 Excel 中验证可出值）'
ws['A1'].font = Font(bold=True, size=14, color='1F4E79')
ws['A2'] = ('基线 Excel %s ｜ SpreadJS %s ｜ Univer %s　　'
            'E 列是活公式（换 Excel 版本打开会重算）；G/J 列先判「该产品是否支持」，I/L 列只在双方都支持时才比较结果。'
            % (EXCEL_VER, SJS_VER, UNIVER_VER))
ws['A2'].font = Font(size=9, color='595959')
for i, h in enumerate(HEADERS, 1):
    c = ws.cell(row=HDR, column=i, value=h)
    c.fill = HEAD_FILL; c.font = HEAD_FONT
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = BORDER
ws.row_dimensions[HDR].height = 34


def judge_formula(res_col, row):
    """生成判定活公式：先看是否支持，易失函数豁免，再比较结果。res_col 取 'H'（SJS）或 'K'（Univer）。"""
    X = res_col
    return (
        '=IF(LEFT(${X}{r},10)="ERR:#NAME?","不支持（#NAME?）",'
        'IF(LEFT(${X}{r},7)="ERR:SET","公式写入被拒",'
        'IF(OR($B{r}="NOW",$B{r}="TODAY",$B{r}="RAND",$B{r}="RANDBETWEEN",$B{r}="RANDARRAY"),'
        '"支持（易失函数，结果不对比）",'
        'IF(LEFT(${X}{r},4)="ERR:",IF(ISERROR($E{r}),"一致（双方同为错误值）","产品报错"),'
        'IF(ISERROR($E{r}),"Excel 报错",'
        'IF(AND(ISNUMBER($E{r}),ISNUMBER(${X}{r})),'
        'IF(ABS($E{r}-${X}{r})<=MAX(0.000000001,ABS($E{r})*0.000000001),"一致","数值不同"),'
        'IF(AND(ISLOGICAL($E{r}),ISNUMBER(${X}{r})),"类型不同（布尔/数值）",'
        'IF(AND(ISNUMBER($E{r}),NOT(ISNUMBER(${X}{r}))),"类型不同（数值/文本）",'
        'IF(EXACT($F{r},${X}{r}&""),"一致","文本不同")))))))))'
    ).format(X=X, r=row)



def put(sheet, row, col, val):
    """写单元格：以 '=' 开头的纯文本必须强制成字符串，否则 openpyxl 会当成活公式"""
    c = sheet.cell(row=row, column=col, value=val)
    if isinstance(val, str) and val.startswith('='):
        c.data_type = 's'
    return c


TOTAL = len(rows) + len(excluded)
last = HDR + TOTAL

for i, r in enumerate(rows):
    row = HDR + 1 + i
    er = 2 + i * GAP
    r['excel_row'] = er
    ws.cell(row=row, column=1, value=i + 1)
    ws.cell(row=row, column=2, value=r['name'])
    ws.cell(row=row, column=3, value=r['family'])
    c4 = ws.cell(row=row, column=4, value=r['formula']); c4.data_type = 's'
    ws.cell(row=row, column=5, value="='Excel实测'!Z%d" % er)
    ws.cell(row=row, column=6, value='=IF(ISERROR(E%d),"ERR",IF(ISLOGICAL(E%d),IF(E%d,"TRUE","FALSE"),E%d&""))' % (row, row, row, row))
    # G: 是否支持
    ws.cell(row=row, column=7, value=(
        '=IF(LEFT(H%d,10)="ERR:#NAME?","不支持",IF(LEFT(H%d,7)="ERR:SET","无法写入","支持"))' % (row, row)))
    put(ws, row, 8, r['sjs'][0])
    # I: 判定
    ws.cell(row=row, column=9, value=judge_formula('H', row))
    ws.cell(row=row, column=10, value=(
        '=IF(LEFT(K%d,10)="ERR:#NAME?","不支持",IF(LEFT(K%d,7)="ERR:SET","无法写入","支持"))' % (row, row)))
    # K: Univer 快照值；布尔行按显示文本写 TRUE/FALSE（t=3，显示层与 Excel 一致）
    if r['univer'][1] == 'bool':
        put(ws, row, 11, 'TRUE' if r['univer'][0] else 'FALSE')
    else:
        put(ws, row, 11, r['univer'][0])
    ws.cell(row=row, column=12, value=judge_formula('K', row))
    ws.cell(row=row, column=13, value=note_for(r))
    for col in range(1, 14):
        cc = ws.cell(row=row, column=col)
        cc.border = BORDER
        cc.font = Font(size=9)
        cc.alignment = Alignment(horizontal='center' if col in (1, 3, 7, 9, 10, 12) else 'left',
                                 vertical='center', wrap_text=(col in (4, 13)))
    ws.cell(row=row, column=4).font = Font(size=9, color='404040')

# 环境依赖行（19 个）：2026-11-05 起用双端代码实测判支持性（写入公式看是否 #NAME?），值不对比
base = HDR + 1 + len(rows)
for j, e in enumerate(excluded):
    row = base + j
    sv, sk = e.get('sjs_env', ('未测', 'text'))
    uv, uk = e.get('uni_env', ('未测', 'text'))
    ws.cell(row=row, column=1, value=len(rows) + j + 1)
    ws.cell(row=row, column=2, value=e['name'])
    ws.cell(row=row, column=3, value=e['family'])
    c4 = ws.cell(row=row, column=4, value='（环境依赖，无法构造 Excel 可出值用例）'); c4.data_type = 's'
    ws.cell(row=row, column=5, value='—')
    ws.cell(row=row, column=6, value='—')
    # G/J 与主行同口径活公式：#NAME? → 不支持
    ws.cell(row=row, column=7, value=(
        '=IF(LEFT(H%d,10)="ERR:#NAME?","不支持",IF(LEFT(H%d,7)="ERR:SET","无法写入","支持"))' % (row, row)))
    put(ws, row, 8, sv)
    # I: 环境依赖只判支持，值不对比
    ws.cell(row=row, column=9, value='=IF(LEFT(H%d,10)="ERR:#NAME?","不支持（#NAME?）","支持（环境依赖，值不对比）")' % row)
    ws.cell(row=row, column=10, value=(
        '=IF(LEFT(K%d,10)="ERR:#NAME?","不支持",IF(LEFT(K%d,7)="ERR:SET","无法写入","支持"))' % (row, row)))
    put(ws, row, 11, uv)
    ws.cell(row=row, column=12, value='=IF(LEFT(K%d,10)="ERR:#NAME?","不支持（#NAME?）","支持（环境依赖，值不对比）")' % row)
    extra = '代码实测（写入公式看返回）：SJS 结果 %s、Univer 结果 %s。' % (
        sv if sk != 'err' else str(sv)[4:], uv if uk != 'err' else str(uv)[4:])
    ws.cell(row=row, column=13, value='需外部数据源/服务，无法在本机单表构造「Excel 能返回正确值」的用例（' + e['reason'] + '），故只判支持、值不对比。' + extra)
    for col in range(1, 14):
        cc = ws.cell(row=row, column=col); cc.border = BORDER; cc.font = Font(size=9, italic=True, color='808080')
        cc.alignment = Alignment(horizontal='center' if col in (1, 3, 7, 9, 10, 12) else 'left', vertical='center', wrap_text=(col in (4, 13)))
        cc.fill = GREY

widths = {1: 5, 2: 20, 3: 9, 4: 52, 5: 20, 6: 16, 7: 9, 8: 20, 9: 19, 10: 8, 11: 20, 12: 19, 13: 40}
for c, w in widths.items():
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'D5'
ws.auto_filter.ref = 'A%d:M%d' % (HDR, last)

# 条件格式
from openpyxl.formatting.rule import FormulaRule
for col in ('I', 'L'):
    rng = '%s%d:%s%d' % (col, HDR + 1, col, last)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("不支持",%s%d))' % (col, HDR + 1)], fill=GREY))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("易失函数",%s%d))' % (col, HDR + 1)], fill=YELLOW))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['ISNUMBER(SEARCH("一致",%s%d))' % (col, HDR + 1)], fill=GREEN))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['NOT(ISNUMBER(SEARCH("一致",%s%d)))*NOT(ISNUMBER(SEARCH("不支持",%s%d)))*NOT(ISNUMBER(SEARCH("易失函数",%s%d)))' % (col, HDR + 1, col, HDR + 1, col, HDR + 1)], fill=RED))

# ---------------- Excel实测（数据区 + 占位） ----------------
for cell in data_block['cells']:
    v = cell['value']
    ws2[cell['cell']] = v
ws2['Y1'] = '函数名'; ws2['Z1'] = '用例公式（活公式）'
ws2['Y1'].font = HEAD_FONT; ws2['Z1'].font = HEAD_FONT
ws2['Y1'].fill = HEAD_FILL; ws2['Z1'].fill = HEAD_FILL
for r in rows:
    ws2.cell(row=r['excel_row'], column=25, value=r['name']).font = Font(size=9)
ws2.column_dimensions['Y'].width = 18
ws2.column_dimensions['Z'].width = 80
ws2.freeze_panes = 'A2'
# 数据区说明
ws2['A14'] = '数据区（与 SpreadJS / Univer 探测页完全一致，定义见 data_block_v2.json）：'
ws2['A14'].font = Font(size=9, italic=True, color='808080')
for k, g in enumerate(data_block['groups']):
    ws2.cell(row=15 + k, column=1, value=g['range']).font = Font(size=9, bold=True, color='808080')
    ws2.cell(row=15 + k, column=2, value=g['note']).font = Font(size=9, color='808080')

# ---------------- 差异明细 ----------------
def fill_diff(sheet, key, who, col_idx):
    sheet['A1'] = '%s：与 Excel 365 结果不一致的函数（仅列「该产品支持该函数」的用例，供人工逐条确认）' % who
    sheet['A1'].font = Font(bold=True, size=12, color='C00000')
    sheet['A2'] = ('筛选口径：产品返回 #NAME? 的行属于「不支持」，不在本表；本表只放「支持但结果/类型与 Excel 不同」和「产品报错」的行。'
                   '顺序按函数族排列。')
    sheet['A2'].font = Font(size=9, color='595959')
    heads = ['#', '函数名', '函数族', '用例公式', 'Excel 365 结果', '%s 结果' % who, '差异类型', '类别', '可能原因']
    for i, h in enumerate(heads, 1):
        c = sheet.cell(row=4, column=i, value=h)
        c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    diffs = [r for r in rows if r[key] in J.DIFF_LABELS]
    diffs.sort(key=lambda r: (r['family'], r['name']))
    for i, r in enumerate(diffs):
        row = 5 + i
        sheet.cell(row=row, column=1, value=i + 1)
        sheet.cell(row=row, column=2, value=r['name'])
        sheet.cell(row=row, column=3, value=r['family'])
        c = sheet.cell(row=row, column=4, value=r['formula']); c.data_type = 's'
        sheet.cell(row=row, column=5, value="='对比表'!E%d" % (HDR + 1 + rows.index(r)))
        sheet.cell(row=row, column=5).number_format = 'General'
        put(sheet, row, 6, r[col_idx][0])
        sheet.cell(row=row, column=7, value=r[key])
        sheet.cell(row=row, column=8, value=category(r, key))
        sheet.cell(row=row, column=9, value=note_for(r))
        for col in range(1, 10):
            cc = sheet.cell(row=row, column=col); cc.border = BORDER; cc.font = Font(size=9)
            cc.alignment = Alignment(horizontal='center' if col in (1, 3, 7, 8) else 'left', vertical='center', wrap_text=(col in (4, 9)))
    for c, w in {1: 5, 2: 20, 3: 9, 4: 50, 5: 22, 6: 22, 7: 20, 8: 13, 9: 42}.items():
        sheet.column_dimensions[get_column_letter(c)].width = w
    sheet.freeze_panes = 'A5'
    sheet.auto_filter.ref = 'A4:I%d' % (4 + len(diffs))
    return len(diffs)


n_s = fill_diff(ws_diff_s, 'sjs_verdict', 'SpreadJS', 'sjs')
n_u = fill_diff(ws_diff_u, 'univer_verdict', 'Univer', 'univer')

# ---------------- 结论 ----------------
ws_sum['A1'] = '结论（全部为活公式，随「对比表」重算自动更新）'
ws_sum['A1'].font = Font(bold=True, size=14, color='1F4E79')
ws_sum['A2'] = '基线 Excel %s ｜ SpreadJS %s ｜ Univer %s' % (EXCEL_VER, SJS_VER, UNIVER_VER)
ws_sum['A2'].font = Font(size=9, color='595959')

hdr = ['指标', 'SpreadJS %s' % SJS_VER, 'Univer %s' % UNIVER_VER, '说明']
for i, h in enumerate(hdr, 1):
    c = ws_sum.cell(row=4, column=i, value=h)
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
    c.alignment = Alignment(horizontal='center', wrap_text=True)

R1 = '%d' % (HDR + 1)
R2 = '%d' % last
T = "'对比表'"
ENVS = HDR + 1 + len(rows)   # 环境依赖行起始

metrics = [
    ('Excel 365 原生函数基线', '=%d' % TOTAL, '=%d' % TOTAL, 'Excel 519 个原生函数全量（含 19 个环境依赖项）'),
    ('可构造有效用例（Excel 能出值）', '=%d' % len(rows), '=%d' % len(rows), '用例先在 Excel 中验证过，返回的不是错误值'),
    ('环境依赖项（无法构造用例，已代码实测支持性）', '=%d' % len(excluded), '=%d' % len(excluded), 'CUBE*/CALL/RTD/WEBSERVICE/STOCKHISTORY 等，需外部数据源或服务；写入公式看是否 #NAME?'),
    ('① 支持——可构造用例', '=COUNTIF(%s!G%d:G%d,"支持")' % (T, HDR + 1, HDR + len(rows)), '=COUNTIF(%s!J%d:J%d,"支持")' % (T, HDR + 1, HDR + len(rows)), '能写入并算出结果（返回 #NAME? 视为不支持）'),
    ('① 支持——环境依赖（代码实测）', '=COUNTIF(%s!G%d:G%d,"支持")' % (T, ENVS, last), '=COUNTIF(%s!J%d:J%d,"支持")' % (T, ENVS, last), '引擎认识该函数即可（GETPIVOTDATA 返回 #REF!、WEBSERVICE 返回空都算认识），值不对比'),
    ('① 支持合计（占 519）', '=B%d+B%d' % (8, 9), '=C%d+C%d' % (8, 9), '产品层面「这个函数有没有」的总口径'),
    ('支持率（占 519）', '=IFERROR(B%d/B%d,"")' % (10, 5), '=IFERROR(C%d/C%d,"")' % (10, 5), '支持合计 ÷ 基线'),
    ('② 不支持（#NAME?）', '=COUNTIF(%s!G%d:G%d,"不支持")' % (T, HDR + 1, last), '=COUNTIF(%s!J%d:J%d,"不支持")' % (T, HDR + 1, last), '产品未实现该函数（含环境依赖行）'),
    ('③ 与 Excel 结果一致（可构造用例内）', '=COUNTIF(%s!I%d:I%d,"一致*")' % (T, HDR + 1, last), '=COUNTIF(%s!L%d:L%d,"一致*")' % (T, HDR + 1, last), '数值容差 1e-9；文本按单元格显示文本精确比对；布尔按显示 TRUE/FALSE 比对'),
    ('一致率（占可构造用例支持数）', '=IFERROR(B%d/B%d,"")' % (13, 8), '=IFERROR(C%d/C%d,"")' % (13, 8), '③ ÷ ①（可构造用例部分）'),
    ('④ 支持但结果不一致', '=COUNTIF(%s!I%d:I%d,"数值不同")+COUNTIF(%s!I%d:I%d,"文本不同")+COUNTIF(%s!I%d:I%d,"类型不同*")+COUNTIF(%s!I%d:I%d,"产品报错")' % ((T, HDR + 1, last) * 4), '=COUNTIF(%s!L%d:L%d,"数值不同")+COUNTIF(%s!L%d:L%d,"文本不同")+COUNTIF(%s!L%d:L%d,"类型不同*")+COUNTIF(%s!L%d:L%d,"产品报错")' % ((T, HDR + 1, last) * 4), '见对应的「差异明细」页'),
    ('其中：易失函数（支持但结果不对比）', '=COUNTIF(%s!I%d:I%d,"支持（易失函数，结果不对比）")' % (T, HDR + 1, last), '=COUNTIF(%s!L%d:L%d,"支持（易失函数，结果不对比）")' % (T, HDR + 1, last), 'NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY：每次重算结果都变，只判支持不比结果'),
    ('其中：产品报错', '=COUNTIF(%s!I%d:I%d,"产品报错")' % (T, HDR + 1, last), '=COUNTIF(%s!L%d:L%d,"产品报错")' % (T, HDR + 1, last), '支持该函数，但这条用例上报了错'),
]
for i, (k, b, c_, note) in enumerate(metrics):
    row = 5 + i
    ws_sum.cell(row=row, column=1, value=k).font = Font(size=10, bold=(i in (5, 8, 10)))
    ws_sum.cell(row=row, column=2, value=b)
    ws_sum.cell(row=row, column=3, value=c_)
    ws_sum.cell(row=row, column=4, value=note).font = Font(size=9, color='595959')
    for col in range(1, 5):
        cc = ws_sum.cell(row=row, column=col); cc.border = BORDER
        cc.alignment = Alignment(horizontal='left' if col in (1, 4) else 'center', vertical='center', wrap_text=(col == 4))
        if col > 1:
            cc.font = Font(size=11, bold=True) if i in (5, 8, 10) else Font(size=10)
for row in (11, 14):   # 百分比行：支持率 / 一致率
    ws_sum.cell(row=row, column=2).number_format = '0.0%'
    ws_sum.cell(row=row, column=3).number_format = '0.0%'

# 差异构成
ws_sum['A20'] = '差异构成（同一函数是否两家都不一致）'
ws_sum['A20'].font = Font(bold=True, size=11, color='1F4E79')
for i, h in enumerate(['', '数量', '说明'], 1):
    c = ws_sum.cell(row=21, column=i, value=h)
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
    c.alignment = Alignment(horizontal='center')
I = "%s!I%d:I%d" % (T, HDR + 1, last)
L = "%s!L%d:L%d" % (T, HDR + 1, last)
comp = [
    ('两家都不一致', '=COUNTIFS(%s,"数值不同",%s,"数值不同")+COUNTIFS(%s,"文本不同",%s,"文本不同")+COUNTIFS(%s,"类型不同*",%s,"类型不同*")' % (I, L, I, L, I, L), '同一条用例两家都没对上'),
    ('仅 SpreadJS 不一致', '=COUNTIFS(%s,"数值不同",%s,"一致*")+COUNTIFS(%s,"文本不同",%s,"一致*")' % (I, L, I, L), 'SpreadJS 侧独有问题'),
    ('仅 Univer 不一致', '=COUNTIFS(%s,"一致*",%s,"数值不同")+COUNTIFS(%s,"一致*",%s,"文本不同")+COUNTIFS(%s,"一致*",%s,"类型不同*")' % (I, L, I, L, I, L), 'Univer 侧独有问题'),
]
for i, (k, f, note) in enumerate(comp):
    row = 22 + i
    ws_sum.cell(row=row, column=1, value=k).font = Font(size=10)
    ws_sum.cell(row=row, column=2, value=f)
    ws_sum.cell(row=row, column=3, value=note).font = Font(size=9, color='595959')
    for col in range(1, 4):
        cc = ws_sum.cell(row=row, column=col); cc.border = BORDER
        cc.alignment = Alignment(horizontal='left' if col != 2 else 'center', wrap_text=(col == 3))

# 差异性质分布（静态，按本次实测）
from collections import Counter
cat_s = Counter(category(r, 'sjs_verdict') for r in rows if r['sjs_verdict'] in J.DIFF_LABELS)
cat_u = Counter(category(r, 'univer_verdict') for r in rows if r['univer_verdict'] in J.DIFF_LABELS)
ws_sum['A29'] = '差异性质分布（静态统计，按本次实测；活公式统计见上表）'
ws_sum['A29'].font = Font(bold=True, size=11, color='1F4E79')
for i, h in enumerate(['类别', 'SpreadJS', 'Univer', '是否功能性差异'], 1):
    c = ws_sum.cell(row=30, column=i, value=h)
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
    c.alignment = Alignment(horizontal='center')
CAT_NOTE = {'精度差异': '否（末位有效数字）', '随机/时间': '已豁免（易失函数只判支持不比结果）',
            '区域设置': '否（货币符号）', '格式差异': '否（大小写）',
            '类型语义': '是（值等价、类型不同）', '功能差异': '是', '报错': '是'}
cats = ['精度差异', '随机/时间', '区域设置', '格式差异', '类型语义', '功能差异', '报错']
for i, cname in enumerate(cats):
    row = 31 + i
    ws_sum.cell(row=row, column=1, value=cname).font = Font(size=10)
    ws_sum.cell(row=row, column=2, value=cat_s.get(cname, 0))
    ws_sum.cell(row=row, column=3, value=cat_u.get(cname, 0))
    ws_sum.cell(row=row, column=4, value=CAT_NOTE.get(cname, '')).font = Font(size=9, color='595959')
    for col in range(1, 5):
        cc = ws_sum.cell(row=row, column=col); cc.border = BORDER
        cc.alignment = Alignment(horizontal='left' if col in (1, 4) else 'center')
trow2 = 31 + len(cats)
ws_sum.cell(row=trow2, column=1, value='合计').font = Font(bold=True)
ws_sum.cell(row=trow2, column=2, value='=SUM(B31:B%d)' % (trow2 - 1)).font = Font(bold=True)
ws_sum.cell(row=trow2, column=3, value='=SUM(C31:C%d)' % (trow2 - 1)).font = Font(bold=True)
for col in range(1, 5):
    cc = ws_sum.cell(row=trow2, column=col); cc.border = BORDER
    cc.fill = PatternFill('solid', fgColor='DDEBF7')
    cc.alignment = Alignment(horizontal='left' if col == 1 else 'center')

ws_sum['A26'] = '提示：差异按性质分「精度差异 / 区域设置 / 类型语义 / 功能差异」，可在两张差异明细页用「类别」列筛选；'
ws_sum['A27'] = '　　　易失函数（NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY）只判支持、结果不对比，不计入一致也不计入不一致。'
ws_sum['A28'] = '提示：环境依赖 19 项只判支持（写入公式看是否 #NAME?）、值不对比；③「一致」不包含「不支持」的行。'
for r_ in (26, 27, 28):
    ws_sum.cell(row=r_, column=1).font = Font(size=9, color='C00000')

for c, w in {1: 30, 2: 18, 3: 18, 4: 70}.items():
    ws_sum.column_dimensions[get_column_letter(c)].width = w

# ---------------- 分族统计 ----------------
ws_fam['A1'] = '分族支持情况（活公式，随「对比表」更新）'
ws_fam['A1'].font = Font(bold=True, size=12, color='1F4E79')
fh = ['函数族', '用例数', 'SJS 支持', 'SJS 不支持', 'SJS 一致', 'SJS 一致率', 'Univer 支持', 'Univer 不支持', 'Univer 一致', 'Univer 一致率']
for i, h in enumerate(fh, 1):
    c = ws_fam.cell(row=3, column=i, value=h)
    c.fill = HEAD_FILL; c.font = HEAD_FONT; c.border = BORDER
    c.alignment = Alignment(horizontal='center', wrap_text=True)
fams = sorted({r['family'] for r in rows})
LASTV = HDR + len(rows)   # 有效用例最后一行：环境依赖行不参与分族统计
for i, fam in enumerate(fams):
    row = 4 + i
    ws_fam.cell(row=row, column=1, value=fam)
    ws_fam.cell(row=row, column=2, value='=COUNTIF(%s!$C$5:$C$%d,A%d)' % (T, LASTV, row))
    ws_fam.cell(row=row, column=3, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$G$5:$G$%d,"支持")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=4, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$G$5:$G$%d,"不支持")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=5, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$I$5:$I$%d,"一致*")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=6, value='=IFERROR(E%d/C%d,"")' % (row, row))
    ws_fam.cell(row=row, column=7, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$J$5:$J$%d,"支持")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=8, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$J$5:$J$%d,"不支持")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=9, value='=COUNTIFS(%s!$C$5:$C$%d,A%d,%s!$L$5:$L$%d,"一致*")' % (T, LASTV, row, T, LASTV))
    ws_fam.cell(row=row, column=10, value='=IFERROR(I%d/C%d,"")' % (row, row))
    ws_fam.cell(row=row, column=6).number_format = '0.0%'
    ws_fam.cell(row=row, column=10).number_format = '0.0%'
    for col in range(1, 11):
        cc = ws_fam.cell(row=row, column=col); cc.border = BORDER; cc.font = Font(size=10)
        if col > 1:
            cc.alignment = Alignment(horizontal='center')
trow = 4 + len(fams)
ws_fam.cell(row=trow, column=1, value='合计').font = Font(bold=True)
for col in range(2, 10):
    if col in (6, 10):
        continue
    L_ = get_column_letter(col)
    ws_fam.cell(row=trow, column=col, value='=SUM(%s4:%s%d)' % (L_, L_, trow - 1)).font = Font(bold=True)
ws_fam.cell(row=trow, column=6, value='=IFERROR(E%d/C%d,"")' % (trow, trow)).number_format = '0.0%'
ws_fam.cell(row=trow, column=10, value='=IFERROR(I%d/C%d,"")' % (trow, trow)).number_format = '0.0%'
for col in range(1, 11):
    cc = ws_fam.cell(row=trow, column=col); cc.border = BORDER
    cc.fill = PatternFill('solid', fgColor='DDEBF7')
    if col > 1:
        cc.alignment = Alignment(horizontal='center')
for c, w in {1: 14, 2: 9, 3: 9, 4: 11, 5: 9, 6: 10, 7: 11, 8: 12, 9: 10, 10: 11}.items():
    ws_fam.column_dimensions[get_column_letter(c)].width = w

# ---------------- 口径说明 ----------------
notes = [
    ('测试目标', '回答两个问题：①这个函数产品支不支持？②支持的话，算得对不对。两件事分开判定，不混成一个「不一致」。'),
    ('基线', 'Excel 365（本机 %s）。Excel 的返回结果是唯一基准。' % EXCEL_VER),
    ('对比对象', 'SpreadJS %s（商业版）｜ Univer %s（商用版）。' % (SJS_VER, UNIVER_VER)),
    ('用例设计原则（v2 重点）', '每条用例都必须先在 Excel 里跑出真实值，禁止拿「双方都报错」当对比结果。'
                        '参数按语义构造：统计用有离散度的数列（CORREL 不恒为 1）、财务用真实票息与期限、'
                        '工程用合法单位、数据库用带表头的数据区和条件区。Excel 侧实测：500 条用例全部返回非错误值（唯一例外 NA 本身就该返回 #N/A）。'),
    ('环境依赖项（2026-11-05 修订）', '19 个函数无法在单表构造出「Excel 能返回正确值」的用例（CUBE* 需 OLAP、CALL/REGISTER.ID 需宏表、'
                          'EUROCONVERT 需加载项、JIS 需日文语言包等），但**支持性仍做了双端代码实测**：把公式写进产品看是否返回 #NAME?。'
                          '结果：SpreadJS 支持 6 个（EUROCONVERT 甚至能算出 196、GETPIVOTDATA/IMAGE/WEBSERVICE/SHEET/SHEETS）、不支持 13 个；'
                          'Univer 支持 3 个（IMAGE/SHEET/SHEETS）、不支持 16 个。环境依赖行只判支持、值不对比。'),
    ('易失函数（2026-11-05 定版）', 'NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY 每次重算结果都不同，做值对比无意义。'
                          '口径：只要支持且返回非错误值，一律记「支持（易失函数，结果不对比）」，不计入一致也不计入不一致。'),
    ('支持判定', '产品写入同一公式后返回 #NAME? → 判「不支持」（函数未实现）；写入被拒 → 「无法写入」；其他错误（含 #REF! 等）→ 引擎认识该函数，算「支持」。'),
    ('一致判定', '只在双方都支持时比较：数值按相对误差 1e-9；文本按单元格显示文本精确比对（EXACT）；'
             '布尔按显示文本 TRUE/FALSE 比对（Univer 快照 v 字段序列化为 1/0，但 t=3 类型标记为布尔、'
             '显示层 getDisplayValue() 返回 TRUE/FALSE，下游语义与 Excel 一致——2026-11-05 实测复核后修正，此前误记「类型不同」）。'),
    ('为什么先判支持', '上一版把 #NAME? 也计入「不一致」，导致「未实现」和「算错了」混在一起。v2 把两者拆开：'
                  '先看支持率，再看一致率（一致率的分母是该产品的支持数）。'),
    ('动态判定', '「对比表」E 列是活公式，换 Excel 版本打开会重算；G/J/I/L 列都是活公式，结果随之更新。'
             'SpreadJS / Univer 两列是实测快照（写死），不随打开变化。'),
    ('快照约定', '快照里的 "ERR:#NAME?" 表示该产品返回了 #NAME? 错误值，不是文本。'),
    ('Excel实测页', '数据区与两个探测页完全一致（见 data_block_v2.json）；用例公式按 12 行间隔存放，容纳动态数组溢出。'),
    ('可复现', 'poc-compare/ 下：data_block_v2.py（数据区）、semantic_cases_v2.py（用例）、excel_case_eval.ps1（Excel oracle 校验）、'
            'sjs_v2.js / univer_v3.js（两产品探测，Univer 探针带 t 类型标记 + 显示文本）、env_probe.js（环境依赖 19 函数双端支持性）、'
            'judge_v2.py（判定）、build_compare_v2.py（本工作簿）。'),
]
for i, (k, v) in enumerate(notes, 1):
    ws_note.cell(row=i, column=1, value=k).font = Font(bold=True, size=10)
    ws_note.cell(row=i, column=1).alignment = Alignment(vertical='top', wrap_text=True)
    ws_note.cell(row=i, column=2, value=v).font = Font(size=10)
    ws_note.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical='top')
ws_note.column_dimensions['A'].width = 24
ws_note.column_dimensions['B'].width = 120

wb.save(OUT)
print('saved', os.path.abspath(OUT))
print('rows', len(rows), '+ env', len(excluded), '=', TOTAL)
print('SJS diff', n_s, '| Univer diff', n_u)
