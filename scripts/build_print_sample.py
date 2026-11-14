# -*- coding: utf-8 -*-
"""构造打印 / PDF 保真测试样本（Excel 基线）。

为什么用 openpyxl 而不是全 COM：COM 的 Cells.Item().Value2 在 PS 5.1 下对 int/DateTime
会抛 InvalidCastException（同一条语句换成 double 又能过），批量写值极不稳定。
openpyxl 负责数据/样式/条件格式/批注/超链接/合并/分组/打印设置，
COM 只补 openpyxl 不支持的形状与迷你图（见 tools/annotate_print_sample.ps1）。

用法: python tools/build_print_sample.py
产出: samples/打印-PDF-多元素样本.xlsx
"""
import os
import random

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "samples", "打印-PDF-多元素样本.xlsx")
LOGO = os.path.join(HERE, "out", "print-logo.png")

THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DCE6F1")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")


def make_logo():
    """用 PIL 画一个小 PNG，避免依赖外部图片。"""
    from PIL import Image, ImageDraw
    os.makedirs(os.path.dirname(LOGO), exist_ok=True)
    img = Image.new("RGB", (200, 60), (19, 70, 164))
    d = ImageDraw.Draw(img)
    d.text((46, 20), "LOGO", fill=(255, 255, 255))
    img.save(LOGO)
    return LOGO


def sheet_data(wb):
    ws = wb.active
    ws.title = "Data"

    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "季度销售报表 Quarterly Sales Report"
    c.font = Font(size=20, bold=True, color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = HDR_FILL
    ws.row_dimensions[1].height = 34

    headers = ["区域", "Q1", "Q2", "Q3", "Q4", "合计", "占比", "日期", "趋势"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    regions = ["华东", "华南", "华北", "西南", "华中", "东北", "西北", "东南", "港澳", "海外"]
    rnd = random.Random(7)
    for i in range(20):
        r = 4 + i
        ws.cell(row=r, column=1, value=f"{regions[i % 10]} {i // 10 + 1}")
        for col in range(2, 6):
            ws.cell(row=r, column=col, value=rnd.randint(120, 980))
        ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
        ws.cell(row=r, column=7, value=f"=F{r}/F$25")
        ws.cell(row=r, column=8, value=f"2025-01-{i*11 % 28 + 1:02d}")
        for col in range(1, 9):
            ws.cell(row=r, column=col).border = BORDER

    ws.cell(row=25, column=1, value="合计 TOTAL")
    for col in range(2, 7):
        letter = chr(64 + col)
        cell = ws.cell(row=25, column=col, value=f"=SUBTOTAL(109,{letter}4:{letter}23)")
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = BORDER
    for col in range(1, 9):
        ws.cell(row=25, column=col).font = Font(bold=True)
        ws.cell(row=25, column=col).fill = TOTAL_FILL
        ws.cell(row=25, column=col).border = BORDER

    for r in range(4, 26):
        for col in range(2, 7):
            ws.cell(row=r, column=col).number_format = "#,##0"
    for r in range(4, 24):
        ws.cell(row=r, column=7).number_format = "0.0%"
        ws.cell(row=r, column=7).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=8).number_format = "yyyy-mm-dd"

    widths = {"A": 16, "B": 10, "C": 10, "D": 10, "E": 10, "F": 11, "G": 9, "H": 12, "I": 14}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    # 条件格式三件套
    ws.conditional_formatting.add("B4:B23", DataBarRule(start_type="min", end_type="max", color="638EC6"))
    ws.conditional_formatting.add("C4:C23", ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add("D4:D23", IconSetRule("3TrafficLights1", "percent", [0, 33, 67]))

    # 批注 + 超链接
    ws["A3"].comment = Comment("表头行，勿改名", "tester", width=160, height=60)
    ws["A27"] = "SpreadJS 官网"
    ws["A27"].hyperlink = "https://www.grapecity.com.cn"
    ws["A27"].font = Font(color="0563C1", underline="single")

    # 分组（大纲）
    for r in range(10, 15):
        ws.row_dimensions[r].outlineLevel = 1
    ws.sheet_properties.outlinePr.summaryBelow = True

    ws.freeze_panes = "A4"

    # 打印设置
    ws.print_area = "$A$1:$I$27"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9           # A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.gridLines = False
    ws.print_options.horizontalCentered = True
    ws.oddHeader.left.text = "机密 · 内部使用"
    ws.oddHeader.left.size = 8
    ws.oddHeader.center.text = "季度销售报表"
    ws.oddHeader.center.size = 11
    ws.oddHeader.center.font = "Arial,Bold"
    ws.oddHeader.right.text = "&D"
    ws.oddFooter.left.text = "由 Excel 生成"
    ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    ws.oddFooter.right.text = "&F"
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.7
    ws.page_margins.bottom = 0.7
    return ws


def sheet_visuals(wb):
    ws = wb.create_sheet("Visuals")
    ws["A1"] = "图表 / 形状 / 图片 保真"
    ws["A1"].font = Font(size=16, bold=True)

    rows = [["月份", "销售额", "成本"],
            ["1月", 320, 210], ["2月", 410, 260], ["3月", 380, 240],
            ["4月", 520, 300], ["5月", 610, 350], ["6月", 580, 330]]
    for i, row in enumerate(rows, start=3):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=v).border = BORDER
    for i in range(3, 10):
        ws.cell(row=i, column=1).font = Font(bold=(i == 3))

    data = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=9)
    cats = Reference(ws, min_col=1, min_row=4, max_row=9)

    bar = BarChart()
    bar.type = "col"
    bar.title = "销售额 vs 成本"
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height, bar.width = 7.4, 12.0
    ws.add_chart(bar, "A12")

    line = LineChart()
    line.title = "趋势"
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.height, line.width = 7.4, 11.0
    ws.add_chart(line, "G12")

    pie = PieChart()
    pie.title = "占比"
    pie.add_data(Reference(ws, min_col=2, min_row=3, max_row=9), titles_from_data=True)
    pie.set_categories(cats)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.height, pie.width = 7.4, 9.5
    ws.add_chart(pie, "A28")

    if os.path.exists(LOGO):
        img = XLImage(LOGO)
        img.width, img.height = 160, 48
        ws.add_image(img, "I28")

    ws.column_dimensions["A"].width = 12
    ws.print_area = "$A$1:$R$44"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_options.gridLines = False
    ws.oddFooter.center.text = "Visuals 第 &P/&N 页"
    return ws


def sheet_mixed(wb):
    ws = wb.create_sheet("Mixed")
    ws.merge_cells("A1:H1")
    ws["A1"] = "长表分页 + 重复标题行"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    headers = ["ID", "名称", "类别", "数量", "单价", "金额", "状态", "备注"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.font = Font(bold=True)
        cell.border = BORDER

    cats = ["硬件", "软件", "服务"]
    states = ["未结", "已结", "挂起"]
    for i in range(60):
        r = 4 + i
        ws.cell(row=r, column=1, value=1000 + i)
        ws.cell(row=r, column=2, value=f"物料 {i + 1}")
        ws.cell(row=r, column=3, value=cats[i % 3])
        ws.cell(row=r, column=4, value=(i % 9) + 1)
        ws.cell(row=r, column=5, value=19.9 + (i % 7) * 10)
        ws.cell(row=r, column=6, value=f"=D{r}*E{r}")
        ws.cell(row=r, column=7, value=states[i % 3])
        ws.cell(row=r, column=8, value="自动生成")
        for col in range(1, 9):
            ws.cell(row=r, column=col).border = BORDER
        ws.cell(row=r, column=5).number_format = "#,##0.00"
        ws.cell(row=r, column=6).number_format = "#,##0.00"

    tab = Table(displayName="TblMixed", ref="A3:H63")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    for k, v in {"A": 8, "B": 14, "C": 10, "D": 8, "E": 10, "F": 12, "G": 8, "H": 14}.items():
        ws.column_dimensions[k].width = v

    ws.print_area = "$A$1:$H$63"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.gridLines = False
    ws.oddFooter.center.text = "Mixed 第 &P/&N 页"
    return ws


def main():
    make_logo()
    wb = Workbook()
    sheet_data(wb)
    sheet_visuals(wb)
    sheet_mixed(wb)
    wb.active = 0
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("saved:", OUT)
    for ws in wb.worksheets:
        print("  sheet:", ws.title, "dims:", ws.dimensions)


if __name__ == "__main__":
    main()
