# -*- coding: utf-8 -*-
"""
工作表 / 工作簿层 能力域样本（A~L 12 个能力域）

样本先行原则（2026-11-06 用户定版）：
  1. 一个模块 = 一份主样本，测试项必须在样本里能指认到具体载体
  2. 元素刻意铺满：可见性三态、保护矩阵、冻结、三级分组、隐藏行列、
     网格线/缩放/RTL/背景图、搜索靶子、文档属性、UsedRange 四档
  3. 造完必须三重核验（verify_sample_worksheet.py 解包数 XML 节点；
     annotate_worksheet_sample.ps1 用 Excel 打开并重存 = Excel 本体确认）

输出：C:/temp/ws_sample.xlsx（ASCII 路径，供 PowerShell COM 阶段读取）
"""
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = r"C:/temp/ws_sample.xlsx"
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="FFFFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="FF1F4E79")
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True)
HEAD_FILL = PatternFill("solid", fgColor="FFDCE6F1")
BODY_FONT = Font(name="微软雅黑", size=10)
NOTE_FONT = Font(name="微软雅黑", size=9, italic=True, color="FF7F7F7F")
THIN = Side(style="thin", color="FFBFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def banner(ws, text, sub="", span=8):
    """每个 sheet 顶部统一的标题带（同时也是 A1 载体）。"""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(row=1, column=1, value=text)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26
    if sub:
        ws.cell(row=2, column=1, value=sub).font = NOTE_FONT
    return 3 if sub else 2


def table(ws, top, headers, rows, widths=None):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=top, column=j, value=h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[top].height = 20
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=top + i, column=j, value=v)
            c.font = BODY_FONT
            c.border = BOX
            c.alignment = Alignment(vertical="center", wrap_text=True)
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + j)].width = w


wb = Workbook()

# ---------------------------------------------------------------- 00 封面
cv = wb.active
cv.title = "00-封面"
cv.sheet_properties.tabColor = "FF808080"
for col, w in zip("ABCD", (34, 22, 46, 14)):
    cv.column_dimensions[col].width = w
r = banner(cv, "工作表 / 工作簿层 能力域样本",
           "SpreadJS 19.2.2  vs  Univer Pro 1.3.1    |    Excel 365 为基线    |    生成于 %s"
           % datetime.now().strftime("%Y-%m-%d %H:%M"), span=4)
cv.cell(row=r, column=1, value="本样本的用法").font = HEAD_FONT
cv.cell(row=r + 1, column=1,
        value="① 直接拖进 SpreadJS 双栏台 / Univer 官方 playground 打开；"
              "② 每个 sheet 对应一个能力域，A 列为可指认载体；"
              "③ 条目编号与《测试计划》一一对应，指不到的项不入矩阵。").font = NOTE_FONT
cv.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=4)
cv.cell(row=r + 1, column=1).alignment = Alignment(wrap_text=True, vertical="top")
cv.row_dimensions[r + 1].height = 34

SHEET_MAP = [
    ("00-封面", "样本说明", "—"),
    ("01-表管理", "A 工作表管理", "表序 / 表名 / 标签色 / 活动表"),
    ("02-可见性", "B 工作表可见性", "visible / hidden / veryHidden 三态"),
    ("03-保护", "C 工作表保护", "密码 + 锁定/解锁区 + 12 项保护开关矩阵"),
    ("04-冻结", "D 冻结与拆分", "5 行 + 3 列冻结，60×12 数据区可滚动"),
    ("05-分组", "E 分级显示 outline", "行 3 级 / 列 2 级 + 汇总行在上"),
    ("06-行列", "F 行列操作", "隐藏行列 / 行高列宽 / 换行待自适应 / 合并表头"),
    ("07-视图", "G 视图", "网格线关闭 + 缩放 80% + 标签色"),
    ("08-搜索", "H 搜索", "大小写 / 公式 / 批注 / 通配符 靶子"),
    ("09-序列化", "I 序列化", "docProps 8 项 + 2 个定义名称 + 自定义属性"),
    ("10-区域", "L 数据区域与选择", "UsedRange 四档：值 / 仅样式 / 仅批注 / 仅验证"),
    ("11-RTL", "G 视图（RTL）", "表级 rightToLeft"),
    ("12-背景图", "G 视图（背景图）", "工作表背景图，Excel COM 补写"),
    ("隐藏表-hidden", "B 可见性", "普通隐藏"),
    ("深度隐藏-veryHidden", "B 可见性", "深度隐藏"),
]
table(cv, r + 3, ["Sheet", "能力域", "承载内容"], SHEET_MAP, widths=(34, 22, 46))
cv.cell(row=r + 3 + len(SHEET_MAP) + 2, column=1,
        value="注意：J（事件/命令/撤销）与 K（剪贴板/拖拽）属交互类，"
              "载体为 04 与 06 的数据区，不单独建 sheet。").font = NOTE_FONT

# ---------------------------------------------------------------- 01 A 表管理
s1 = wb.create_sheet("01-表管理")
s1.sheet_properties.tabColor = "FF2E75B6"
r1 = banner(s1, "A 工作表管理", "载体：本簿 15 个 sheet 的存在、顺序、名称、标签色、活动表")
s1.column_dimensions["A"].width = 8
s1.column_dimensions["B"].width = 30
s1.column_dimensions["C"].width = 26
s1.column_dimensions["D"].width = 30
s1.cell(row=r1, column=1, value="A 组测试项（每项的载体就是本簿的表结构本身）").font = HEAD_FONT
A_ITEMS = [
    ("A1", "新建 / 插入工作表", "本簿共 15 个 sheet，可再插入"),
    ("A2", "删除工作表", "隐藏表可被删（未删，作为 B 组载体保留）"),
    ("A3", "重命名", "每个 sheet 名都是可改载体"),
    ("A4", "移动 / 重排", "见下表「预期表序」，移动后应变化"),
    ("A5", "复制工作表", "01 表含表对象与标签色，复制后应保持一致"),
    ("A6", "取表数量 / 按索引取表", "15"),
    ("A7", "按名取表", "含中文名与带连字符的名字"),
    ("A8", "活动表切换", "当前活动表 = 00-封面"),
    ("A9", "标签色", "6 个 sheet 显式设色（灰/蓝/绿/红/橙/紫）"),
    ("A10", "多选工作表（成组）", "无 Excel 载体（视图态），仅 API 判定"),
]
table(s1, r1 + 1, ["ID", "能力", "载体"], A_ITEMS, widths=(8, 30, 46))
s1.cell(row=r1 + 1 + len(A_ITEMS) + 2, column=1,
        value="预期表序（自上而下即标签顺序）").font = HEAD_FONT
order_rows = [(i + 1, n) for i, n in enumerate(wb.sheetnames)]
table(s1, r1 + 1 + len(A_ITEMS) + 3, ["#", "Sheet 名"], order_rows, widths=(8, 30))

# ---------------------------------------------------------------- 02 B 可见性
s2 = wb.create_sheet("02-可见性")
r2 = banner(s2, "B 工作表可见性（三态）",
            "载体：本表可见；「隐藏表-hidden」普通隐藏；「深度隐藏-veryHidden」深度隐藏")
s2.column_dimensions["A"].width = 8
s2.column_dimensions["B"].width = 30
s2.column_dimensions["C"].width = 34
B_ITEMS = [
    ("B1", "visible（默认）", "本表 + 其余 12 个可见表"),
    ("B2", "hidden（普通隐藏）", "sheet「隐藏表-hidden」，用户可从右键菜单取消隐藏"),
    ("B3", "veryHidden（深度隐藏）", "sheet「深度隐藏-veryHidden」，不出现在取消隐藏列表"),
    ("B4", "读三态", "SJS 返回 0/1/2；Univer getHiddenState 返回 0/1/2"),
    ("B5", "veryHidden 行为约束", "不能设为活动表 / 不能被多选 / 全表查找检索不到 / 不能打印"),
    ("B6", "唯一可见表不可隐藏", "Excel 与 Univer 都会报错；SJS 需实测"),
]
table(s2, r2, ["ID", "能力", "载体 / 期望"], B_ITEMS, widths=(8, 30, 56))
s2.cell(row=r2 + len(B_ITEMS) + 2, column=1,
        value="验证动作：在 Excel 中右键任一标签 →「取消隐藏」，应只列出 hidden 表，不列出 veryHidden 表。").font = NOTE_FONT

# ---------------------------------------------------------------- 03 C 保护
s3 = wb.create_sheet("03-保护")
s3.sheet_properties.tabColor = "FF548235"
r3 = banner(s3, "C 工作表保护",
            "密码 sjs2026 ｜ B17:C24 为解锁输入区（黄底）｜ 其余单元格锁定")
for col, w in zip("ABCD", (10, 18, 18, 40)):
    s3.column_dimensions[col].width = w
s3.cell(row=r3, column=1, value="C 组：保护开关矩阵（设 → 读回 → 行为三验）").font = HEAD_FONT
C_ITEMS = [
    ("C1", "开启/关闭保护", "isProtected / protect() / unprotect(pwd)"),
    ("C2", "密码", "sjs2026；错误密码不应抛异常而是返回 false"),
    ("C3", "解锁单元格", "B17:C24 locked=False（黄底可见）"),
    ("C4", "允许插入行/列", "Excel 勾选；SJS allowInsertRows/Columns"),
    ("C5", "禁止删除行/列", "Excel 未勾选；SJS allowDeleteRows/Columns=false"),
    ("C6", "允许排序 / 筛选", "Excel 勾选"),
    ("C7", "允许设置单元格格式", "Excel 勾选"),
    ("C8", "禁止设置列/行格式", "Excel 未勾选"),
    ("C9", "禁止插入超链接 / 透视表", "Excel 未勾选"),
    ("C10", "受保护后仍可 API 插行", "SJS 文档明示：只限 UI，代码不受约束"),
]
table(s3, 4, ["ID", "能力", "载体 / 期望"], C_ITEMS, widths=(10, 26, 52))

s3.cell(row=16, column=1, value="输入区（解锁，locked=False）").font = HEAD_FONT
input_fill = PatternFill("solid", fgColor="FFFFF2CC")
for rr in range(17, 25):
    s3.cell(row=rr, column=1, value="第 %d 行" % (rr - 16)).font = BODY_FONT
    for cc in (2, 3):
        c = s3.cell(row=rr, column=cc, value="可编辑" if cc == 2 else (rr - 16) * 100)
        c.fill = input_fill
        c.border = BOX
        c.protection = Protection(locked=False)

p = s3.protection
p.sheet = True
p.password = "sjs2026"
p.insertRows = False        # False = 允许（OOXML 属性 0）
p.insertColumns = False
p.sort = False
p.autoFilter = False
p.formatCells = False
p.deleteRows = True         # True = 禁止
p.deleteColumns = True
p.formatColumns = True
p.formatRows = True
p.insertHyperlinks = True
p.pivotTables = True
p.selectLockedCells = False
p.selectUnlockedCells = False
p.objects = False
p.scenarios = False

# ---------------------------------------------------------------- 04 D 冻结
s4 = wb.create_sheet("04-冻结")
s4.sheet_properties.tabColor = "FFC00000"
r4 = banner(s4, "D 冻结与拆分", "冻结 5 行 + 3 列（freeze_panes = D6），数据区 60 行 × 12 列")
s4.cell(row=2, column=1,
        value="验证：向下滚动到 30 行附近，第 1~5 行应固定；向右滚动，A~C 列应固定。").font = NOTE_FONT
s4.freeze_panes = "D6"
heads4 = ["区域", "产品", "季度", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月"]
for j, h in enumerate(heads4, start=1):
    c = s4.cell(row=5, column=j, value=h)
    c.font = HEAD_FONT
    c.fill = HEAD_FILL
    c.border = BOX
for i in range(60):
    rr = 6 + i
    s4.cell(row=rr, column=1, value="R%02d" % (i + 1)).font = BODY_FONT
    s4.cell(row=rr, column=2, value="产品-%02d" % ((i % 12) + 1)).font = BODY_FONT
    s4.cell(row=rr, column=3, value="Q%d" % ((i % 4) + 1)).font = BODY_FONT
    for j in range(4, 13):
        c = s4.cell(row=rr, column=j, value=float((i + 1) * j % 97))
        c.font = BODY_FONT
for col, w in zip("ABC", (8, 12, 8)):
    s4.column_dimensions[col].width = w

# ---------------------------------------------------------------- 05 E 分组
s5 = wb.create_sheet("05-分组")
r5 = banner(s5, "E 分级显示（outline）",
            "行：3-5 组1、6-9 组1（内含 7-8 组2 嵌套）｜列：B-D 组1、C-D 组2｜汇总行在上")
for col, w in zip("ABCDE", (22, 12, 12, 12, 12)):
    s5.column_dimensions[col].width = w
s5.cell(row=2, column=1, value="区域").font = HEAD_FONT
s5.cell(row=2, column=1).fill = HEAD_FILL
for j, h in enumerate(["Q1", "Q2", "Q3", "Q4"], start=2):
    c = s5.cell(row=2, column=j, value=h)
    c.font = HEAD_FONT
    c.fill = HEAD_FILL
BLOCKS = [
    ("华东", [120, 135, 128, 150]),
    ("华南", [98, 110, 105, 121]),
    ("华北", [76, 88, 92, 101]),
    ("西南", [54, 61, 58, 66]),
    ("东北", [43, 47, 45, 52]),
    ("西北", [31, 36, 34, 39]),
    ("合计", [422, 477, 462, 529]),
]
for i, (name, vals) in enumerate(BLOCKS):
    rr = 3 + i
    s5.cell(row=rr, column=1, value=name).font = BODY_FONT
    for j, v in enumerate(vals, start=2):
        s5.cell(row=rr, column=j, value=float(v)).font = BODY_FONT
s5.sheet_properties.outlinePr.summaryBelow = False   # 汇总行在上（Excel: xlSummaryAbove）
s5.row_dimensions.group(3, 5, outline_level=1, hidden=False)
s5.row_dimensions.group(6, 9, outline_level=1, hidden=False)
s5.row_dimensions.group(7, 8, outline_level=2, hidden=False)
s5.column_dimensions.group("B", "D", outline_level=1, hidden=False)
s5.column_dimensions.group("C", "D", outline_level=2, hidden=False)
s5.cell(row=12, column=1,
        value="验证：左侧应出现 1/2 两级「+」按钮，点 1 级可整块折叠；上方列标出现列分组按钮。").font = NOTE_FONT

# ---------------------------------------------------------------- 06 F 行列
s6 = wb.create_sheet("06-行列")
r6 = banner(s6, "F 行列操作",
            "隐藏 D、F 列与 12、13 行｜行 4 高 36｜B 列宽 26｜B8 长文本换行待自适应｜合并表头")
for col, w in zip("ABCDEF", (14, 26, 14, 14, 14, 14)):
    s6.column_dimensions[col].width = w
s6.merge_cells("B3:D3")
s6.cell(row=3, column=2, value="上半年度").font = HEAD_FONT
s6.cell(row=3, column=2).fill = HEAD_FILL
for j, h in enumerate(["项目", "说明", "Q1", "Q2", "Q3", "Q4"], start=1):
    c = s6.cell(row=4, column=j, value=h)
    c.font = HEAD_FONT
    c.fill = HEAD_FILL
    c.border = BOX
s6.row_dimensions[4].height = 36
s6.column_dimensions["B"].width = 26
F_ROWS = [
    ("行高", "本行高被设为 36 像素点（24pt）", 1, 2, 3, 4),
    ("列宽", "B 列宽 26，明显宽于相邻列", 5, 6, 7, 8),
    ("隐藏列", "D 列与 F 列被隐藏，A-E 之间应断开", 9, 10, 11, 12),
    ("隐藏行", "第 12、13 行被隐藏，跳号可见", 13, 14, 15, 16),
    ("自动换行", "这一列文字刻意很长，用来验证 autoFitRow 是否会把行高撑开；"
                "若行高保持不变，说明自适应未生效。", 17, 18, 19, 20),
    ("合并表头", "B3:D3 已合并，承载多级表头语义", 21, 22, 23, 24),
]
for i, row in enumerate(F_ROWS):
    for j, v in enumerate(row, start=1):
        c = s6.cell(row=5 + i, column=j, value=v)
        c.font = BODY_FONT
        c.border = BOX
        if j == 2:
            c.alignment = Alignment(wrap_text=True, vertical="top")
s6.row_dimensions[12].hidden = True
s6.row_dimensions[13].hidden = True
s6.column_dimensions["D"].hidden = True
s6.column_dimensions["F"].hidden = True
s6.cell(row=15, column=1,
        value="验证：D 列标题应缺失（B→C→E 相连）；行号应跳过 12、13；"
              "「自动换行」那行若只显示一行被截断，说明行高自适应未生效。").font = NOTE_FONT

# ---------------------------------------------------------------- 07 G 视图
s7 = wb.create_sheet("07-视图")
s7.sheet_properties.tabColor = "FFED7D31"
r7 = banner(s7, "G 视图：网格线 / 缩放",
            "本表网格线已关闭（showGridLines=0）｜缩放 80%｜标签色橙")
s7.sheet_view.showGridLines = False
s7.sheet_view.zoomScale = 80
s7.sheet_view.zoomScaleNormal = 80
for col, w in zip("ABCD", (22, 16, 16, 30)):
    s7.column_dimensions[col].width = w
G_ITEMS = [
    ("G1", "网格线显隐", "本表无网格线；其他表有"),
    ("G2", "网格线颜色", "无 Excel 载体（Excel 不存单表网格线色），仅 API 判定"),
    ("G3", "缩放比例", "本表 80%；切回其他表应为 100%"),
    ("G4", "标签色", "本表橙色；01 蓝 / 04 绿 / 05 红 / 12 紫"),
    ("G5", "RTL", "见 sheet「11-RTL」"),
    ("G6", "背景图", "见 sheet「12-背景图」"),
    ("G7", "焦点单元格高亮", "无 Excel 载体（视图态），仅 API 判定"),
    ("G8", "显示公式 / 显示零值（视图态）", "仅 API 判定，不入导入矩阵"),
]
table(s7, r7, ["ID", "能力", "载体 / 期望"], G_ITEMS, widths=(8, 30, 52))

# ---------------------------------------------------------------- 08 H 搜索
s8 = wb.create_sheet("08-搜索")
r8 = banner(s8, "H 搜索 / 查找替换",
            "靶子：大小写三态、公式文本、批注文本、通配符字面量")
for col, w in zip("ABCDE", (10, 26, 26, 16, 34)):
    s8.column_dimensions[col].width = w
s8.cell(row=3, column=1, value="大小写").font = HEAD_FONT
for i, (txt, note) in enumerate([
    ("ALPHA", "全大写"),
    ("alpha", "全小写"),
    ("Alpha test", "首字母大写 + 后续单词"),
    ("ALPHABET", "包含关系，测整词匹配应为不命中"),
]):
    s8.cell(row=4 + i, column=2, value=txt).font = BODY_FONT
    s8.cell(row=4 + i, column=3, value=note).font = NOTE_FONT
s8.cell(row=9, column=1, value="通配符").font = HEAD_FONT
for i, (txt, note) in enumerate([
    ("a*b", "字面量星号，非通配符时应只命中自身"),
    ("axb", "可被 a*b 通配符命中"),
    ("a1b", "可被 a?b / a*b 命中"),
]):
    s8.cell(row=10 + i, column=2, value=txt).font = BODY_FONT
    s8.cell(row=10 + i, column=3, value=note).font = NOTE_FONT
s8.cell(row=14, column=1, value="公式").font = HEAD_FONT
for i, f in enumerate(["=SUM(D4:D6)", "=1+2", "=CONCAT(\"al\",\"pha\")"]):
    s8.cell(row=15 + i, column=2, value=f).font = BODY_FONT
for rr in (4, 5, 6):
    s8.cell(row=rr, column=4, value=float(rr * 10)).font = BODY_FONT
s8.cell(row=19, column=1, value="批注").font = HEAD_FONT
s8.cell(row=20, column=2, value="批注由 COM 阶段写入")
s8.cell(row=20, column=3, value="B20 批注含「促销」").font = NOTE_FONT
# 注意：openpyxl 写的批注 Excel 只读出 1 个字符、重存时会整条丢失，
# 所以批注文本落在 C:/temp/ws_cmt*.txt，由 annotate_worksheet_sample.ps1
# 用 Range.AddComment 写（见该脚本）。
s8.cell(row=22, column=1, value="超链接").font = HEAD_FONT
lh = s8.cell(row=23, column=2, value="SpreadJS 官网")
lh.hyperlink = "https://demo.grapecity.com.cn/spreadjs/"
lh.font = Font(name="微软雅黑", size=10, color="FF0563C1", underline="single")

# ---------------------------------------------------------------- 09 I 序列化
s9 = wb.create_sheet("09-序列化")
r9 = banner(s9, "I 序列化 / 文档属性",
            "docProps 8 项 + 2 个定义名称 + 1 个自定义文档属性（COM 补写）")
for col, w in zip("ABCD", (22, 20, 20, 40)):
    s9.column_dimensions[col].width = w
s9.cell(row=3, column=1, value="文档属性（core.xml）").font = HEAD_FONT
DOC_ROWS = [
    ("title", "工作表与工作簿能力域样本"),
    ("subject", "SpreadJS vs Univer 工作簿/工作表层对比"),
    ("creator", "对比测试工程"),
    ("lastModifiedBy", "table-sdk-compare"),
    ("category", "SDK 对比"),
    ("keywords", "SpreadJS;Univer;Worksheet;Workbook"),
    ("description", "15 个 sheet 覆盖 A~L 12 个能力域"),
    ("自定义属性 CompareItem", "COM 阶段写入，见 custom.xml"),
]
table(s9, 4, ["字段", "值"], DOC_ROWS, widths=(26, 44))
s9.cell(row=14, column=1, value="定义名称").font = HEAD_FONT
table(s9, 15, ["名称", "指向", "用途"], [
    ("TaxRate", "'09-序列化'!$B$20", "单值常量名"),
    ("DataBlock", "'10-区域'!$A$1:$D$6", "区域名，跨表引用"),
], widths=(26, 30, 30))
s9.cell(row=19, column=1, value="计算区（依赖 TaxRate 名称）").font = HEAD_FONT
s9.cell(row=20, column=1, value="税率").font = BODY_FONT
s9.cell(row=20, column=2, value=0.13).font = BODY_FONT
s9.cell(row=21, column=1, value="金额").font = BODY_FONT
s9.cell(row=21, column=2, value=1000).font = BODY_FONT
s9.cell(row=22, column=1, value="含税").font = BODY_FONT
s9.cell(row=22, column=2, value="=B21*(1+TaxRate)").font = BODY_FONT
wb.defined_names["TaxRate"] = DefinedName("TaxRate", attr_text="'09-序列化'!$B$20")
wb.defined_names["DataBlock"] = DefinedName("DataBlock", attr_text="'10-区域'!$A$1:$D$6")

# ---------------------------------------------------------------- 10 L 区域
s10 = wb.create_sheet("10-区域")
r10 = banner(s10, "L 数据区域与选择",
             "UsedRange 四档：A1:D6 有值｜H3 仅样式｜K20 仅批注｜M30 仅数据验证")
for col, w in zip("ABCDEFGHIJKLM", [16, 14, 14, 14, 3, 3, 3, 12, 3, 3, 10, 3, 12]):
    s10.column_dimensions[col].width = w
s10.cell(row=3, column=1, value="区域").font = HEAD_FONT
for j, h in enumerate(["区域", "Q1", "Q2", "Q3"], start=1):
    c = s10.cell(row=3, column=j, value=h)
    c.font = HEAD_FONT
    c.fill = HEAD_FILL
    c.border = BOX
for i, (nm, a, b, c_) in enumerate([
    ("华东", 120, 135, 128), ("华南", 98, 110, 105),
    ("华北", 76, 88, 92), ("西南", 54, 61, 58),
]):
    rr = 4 + i
    s10.cell(row=rr, column=1, value=nm).font = BODY_FONT
    for j, v in enumerate([a, b, c_], start=2):
        s10.cell(row=rr, column=j, value=float(v)).font = BODY_FONT
# 仅样式无值
sty = s10.cell(row=3, column=8, value=None)
sty.fill = PatternFill("solid", fgColor="FFFFC7CE")
sty.border = BOX
s10.cell(row=2, column=8, value="← H3 仅样式无值").font = NOTE_FONT
# 仅批注（文本由 COM 阶段写入，见上）
kc = s10.cell(row=20, column=11, value=None)
s10.cell(row=19, column=11, value="K20 仅批注").font = NOTE_FONT
# 仅数据验证
dv = DataValidation(type="list", formula1='"甲,乙,丙"', allow_blank=True)
s10.add_data_validation(dv)
dv.add("M30")
s10.cell(row=29, column=13, value="M30 仅数据验证").font = NOTE_FONT
s10.cell(row=32, column=1,
         value="期望：getUsedRange(data)=A1:D6；含 style 应扩到 H3；"
               "含 comment 应含 K20；含 dataValidation 应含 M30。").font = NOTE_FONT

# ---------------------------------------------------------------- 11 RTL
s11 = wb.create_sheet("11-RTL")
r11 = banner(s11, "G 视图：RTL", "本表 DisplayRightToLeft = True，列顺序从右往左")
s11.sheet_view.rightToLeft = True
s11.column_dimensions["A"].width = 20
s11.column_dimensions["B"].width = 20
s11.cell(row=3, column=1, value="第一列（RTL 下显示在右侧）").font = BODY_FONT
s11.cell(row=3, column=2, value=100).font = BODY_FONT
s11.cell(row=4, column=1, value="第二列").font = BODY_FONT
s11.cell(row=4, column=2, value=200).font = BODY_FONT
s11.cell(row=6, column=1,
         value="注意：Univer 官方明示 RTL 仅 UI 层，渲染层不支持；"
               "本表用于验证其是否真的不镜像。").font = NOTE_FONT

# ---------------------------------------------------------------- 12 背景图
s12 = wb.create_sheet("12-背景图")
s12.sheet_properties.tabColor = "FF7030A0"
r12 = banner(s12, "G 视图：工作表背景图", "背景图由 COM 阶段 SetBackgroundPicture 写入")
for col, w in zip("ABC", (26, 26, 26)):
    s12.column_dimensions[col].width = w
s12.cell(row=3, column=1, value="背景图验证点").font = HEAD_FONT
for i, t in enumerate([
    "① 背景图是否随 xlsx 一起被导入",
    "② 是否平铺（Excel 默认平铺）",
    "③ 与单元格填充色的层次关系（单元格色应盖在背景图之上）",
    "④ Univer 明示：背景图不参与打印",
]):
    s12.cell(row=4 + i, column=1, value=t).font = BODY_FONT

# ---------------------------------------------------------------- 13/14 隐藏表
sh = wb.create_sheet("隐藏表-hidden")
banner(sh, "隐藏表（hidden）", "普通隐藏：用户可从右键菜单「取消隐藏」找回")
sh.cell(row=3, column=1, value="本表 Visible = xlSheetHidden (0)").font = BODY_FONT
sh.cell(row=4, column=1, value="关键字 alpha 也在这里出现一次，用于验证隐藏表是否被全表查找命中").font = NOTE_FONT

vh = wb.create_sheet("深度隐藏-veryHidden")
banner(vh, "深度隐藏表（veryHidden）", "深度隐藏：不出现在「取消隐藏」对话框")
vh.cell(row=3, column=1, value="本表 Visible = xlSheetVeryHidden (2)").font = BODY_FONT

sh.sheet_state = "hidden"
vh.sheet_state = "veryHidden"

# ---------------------------------------------------------------- 文档属性
pr = wb.properties
pr.title = "工作表与工作簿能力域样本"
pr.subject = "SpreadJS vs Univer 工作簿/工作表层对比"
pr.creator = "对比测试工程"
pr.lastModifiedBy = "table-sdk-compare"
pr.category = "SDK 对比"
pr.keywords = "SpreadJS;Univer;Worksheet;Workbook"
pr.description = "15 个 sheet 覆盖 A~L 12 个能力域"
pr.created = datetime(2026, 11, 6, 10, 0, 0)
pr.modified = datetime(2026, 11, 6, 10, 0, 0)

# 批注文本：UTF-8 无 BOM，由 COM 阶段用 Range.AddComment 写入
for p, t in [
    (r"C:/temp/ws_cmt1.txt", "促销活动复核：本批注包含关键字 促销，用于测试按批注查找"),
    (r"C:/temp/ws_cmt2.txt", "K20 只有批注，没有值和样式"),
]:
    with open(p, "w", encoding="utf-8") as f:
        f.write(t)

wb.active = 0
os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)

# ---------------------------------------------------------------- 自定义文档属性
# Excel COM 的 CustomDocumentProperties.Add() 在“包里还没有 custom.xml”时会抛
# NullReferenceException（COM 已知坑），所以这里直接注入 OOXML 部件，再交给
# Excel 重存。三个部件缺一不可：docProps/custom.xml 本体、[Content_Types].xml
# 的 Override、_rels/.rels 的关系项。
import zipfile

CUSTOM_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"'
    ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="CompareItem">'
    '<vt:lpwstr>table-sdk-compare/worksheet</vt:lpwstr></property></Properties>'
)
CUSTOM_REL = (
    '<Relationship Id="rIdCustomProps" Type="http://schemas.openxmlformats.org/'
    'officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>'
)
CUSTOM_OVERRIDE = (
    '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.'
    'openxmlformats-officedocument.custom-properties+xml"/>'
)


def inject_custom_props(path):
    with zipfile.ZipFile(path) as zin:
        data = {n: zin.read(n) for n in zin.namelist()}
    ct = data["[Content_Types].xml"].decode("utf-8")
    if "custom-properties" not in ct:
        ct = ct.replace("</Types>", CUSTOM_OVERRIDE + "</Types>")
        data["[Content_Types].xml"] = ct.encode("utf-8")
    rels = data["_rels/.rels"].decode("utf-8")
    if "docProps/custom.xml" not in rels:
        rels = rels.replace("</Relationships>", CUSTOM_REL + "</Relationships>")
        data["_rels/.rels"] = rels.encode("utf-8")
    data["docProps/custom.xml"] = CUSTOM_XML.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items():
            zout.writestr(n, b)
    os.replace(tmp, path)


inject_custom_props(OUT)
print("custom property injected: CompareItem")

print("saved:", OUT)
print("sheets:", len(wb.sheetnames))
for i, n in enumerate(wb.sheetnames):
    print("  %02d %s" % (i + 1, n))
