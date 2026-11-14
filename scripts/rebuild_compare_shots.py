# -*- coding: utf-8 -*-
# 重拼批次1/批次2 三列导入还原对照图：Excel 基线(COM Chart.Export) | SpreadJS 导入 | Univer 导入
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = r"D:\workbuddy\spreadjs vs univer"
POC = os.path.join(ROOT, "poc-compare")
BASE = os.path.join(POC, "baseline_export")

FONT = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 22)
FONT_S = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 18)

COL_W = 500
HEAD_H = 56
ROW_H = 340  # 单元格可视高度（等比缩放后不足补白）
PAD = 8

HEADS = ["Excel 基线（COM 造数）", "SpreadJS 19.2.2 导入", "Univer Pro(playground) 导入"]

def load_fit(path, w=COL_W, h=ROW_H):
    if not os.path.exists(path):
        img = Image.new("RGB", (w, h), "#EEEEEE")
        d = ImageDraw.Draw(img)
        d.text((10, h // 2 - 12), "（缺素材）", font=FONT_S, fill="#888888")
        return img
    img = Image.open(path).convert("RGB")
    r = min(w / img.width, h / img.height)
    img = img.resize((max(1, int(img.width * r)), max(1, int(img.height * r))), Image.LANCZOS)
    cell = Image.new("RGB", (w, h), "white")
    cell.paste(img, (0, 0))
    return cell

def build(rows, out):
    # rows: [(行标注, excel路径, sjs路径, univer路径), ...]
    W = COL_W * 3 + PAD * 4
    H = HEAD_H + (HEAD_H + ROW_H + 26 + PAD) * len(rows) + PAD
    canvas = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(canvas)
    # 列标题
    for c, t in enumerate(HEADS):
        x = PAD + c * (COL_W + PAD)
        d.rectangle([x, 0, x + COL_W, HEAD_H - 6], fill="#1F3B5C")
        d.text((x + 12, 12), t, font=FONT, fill="white")
    y = HEAD_H
    for label, p_e, p_s, p_u in rows:
        # 行标注条
        d.rectangle([PAD, y, W - PAD, y + 26], fill="#FFF3E0")
        d.text((PAD + 8, y + 2), label, font=FONT_S, fill="#B26A00")
        y += 30
        for c, p in enumerate([p_e, p_s, p_u]):
            x = PAD + c * (COL_W + PAD)
            cell = load_fit(p)
            canvas.paste(cell, (x, y))
            d.rectangle([x, y, x + COL_W, y + ROW_H], outline="#CCCCCC")
        y += ROW_H + PAD
    canvas.save(out)
    print("saved:", out, canvas.size)

B1 = [
    ("01 标题样式（文本/字体/颜色）",
     r"{}\b1_01标题_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_01标题.png"), os.path.join(POC, "univer_api_batch1_01.png")),
    ("02 轴刻度（类型/主刻度单位/min-max）",
     r"{}\b1_02轴刻度_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_02轴刻度.png"), os.path.join(POC, "univer_api_batch1_02.png")),
    ("03 刻度标签（旋转45°/字体颜色/粗细）",
     r"{}\b1_03刻度标签_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_03刻度标签.png"), os.path.join(POC, "univer_api_batch1_03.png")),
    ("04 网格线（红色虚线/颜色/粗细）",
     r"{}\b1_04网格线_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_04网格线.png"), os.path.join(POC, "univer_api_batch1_04.png")),
    ("05 图例（位置/字体/颜色）",
     r"{}\b1_05图例_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_05图例.png"), os.path.join(POC, "univer_api_batch1_05.png")),
    ("06 数据标签（显隐/格式/字体）",
     r"{}\b1_06数据标签_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_06数据标签.png"), os.path.join(POC, "univer_api_batch1_06.png")),
    ("07 轴标题（文本/字体/颜色/旋转）",
     r"{}\b1_07轴标题_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch1_07轴标题.png"), os.path.join(POC, "univer_api_batch1_07.png")),
]
B2 = [
    ("08 系列样式（填充/边框/标记/透明度）",
     r"{}\b2_08系列样式_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch2_08系列样式.png"), os.path.join(POC, "univer_api_batch2_08.png")),
    ("09 高级元素（次轴组合图/趋势线+公式R²/误差线）",
     r"{}\b2_09高级元素_c1.png".format(BASE), os.path.join(POC, "sjs_api_batch2_09高级元素.png"), os.path.join(POC, "univer_api_batch2_09.png")),
]

# 备份旧图
for name in ("证据图-批次1配置还原对照.png", "证据图-批次2系列与高级元素对照.png"):
    p = os.path.join(ROOT, name)
    if os.path.exists(p):
        os.replace(p, p + ".bak")
        print("backup:", name + ".bak")

build(B1, os.path.join(ROOT, "证据图-批次1配置还原对照.png"))
build(B2, os.path.join(ROOT, "证据图-批次2系列与高级元素对照.png"))
