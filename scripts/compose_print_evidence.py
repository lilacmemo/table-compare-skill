# -*- coding: utf-8 -*-
"""拼打印 / PDF 证据图：Excel / SpreadJS / Univer 三栏对照。
用法: python tools/compose_print_evidence.py
输出: tools/证据图-打印三栏对照.png、tools/证据图-打印图表形状保真.png
"""
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out", "print")

FONT_PATHS = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
]
H = 640          # 每栏图片高度
PAD = 18
BAR = 46
BG = (245, 246, 248)
LINE = (210, 214, 220)


def font(size):
    for p in FONT_PATHS:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def strip(path, title, target_h=H):
    img = Image.open(path).convert("RGB")
    w = int(img.width * target_h / img.height)
    return title, img.resize((w, target_h), Image.LANCZOS)


def compose(items, out_name, note=None):
    cols = [strip(p, t) for p, t in items]
    total_w = sum(c[1].width for c in cols) + PAD * (len(cols) + 1)
    total_h = H + BAR + PAD * 2 + (34 if note else 0)
    canvas = Image.new("RGB", (total_w, total_h), BG)
    d = ImageDraw.Draw(canvas)
    f = font(22)
    fs = font(16)

    x = PAD
    for title, img in cols:
        d.rectangle([x, PAD, x + img.width, PAD + BAR - 8], fill=(31, 41, 55))
        tw = d.textlength(title, font=f)
        d.text((x + (img.width - tw) / 2, PAD + 8), title, font=f, fill=(255, 255, 255))
        canvas.paste(img, (x, PAD + BAR))
        d.rectangle([x, PAD + BAR, x + img.width, PAD + BAR + H], outline=LINE, width=1)
        x += img.width + PAD

    if note:
        d.text((PAD, total_h - 26), note, font=fs, fill=(90, 96, 105))

    path = os.path.join(HERE, out_name)
    canvas.save(path)
    print("saved", path, canvas.size)


def main():
    compose(
        [
            (os.path.join(OUT, "excel-p1.png"), "Excel 打印 / PDF（基线）"),
            (os.path.join(OUT, "sjs-0-Data-p1.png"), "SpreadJS savePDF"),
            (os.path.join(OUT, "univer-print-p1.png"), "Univer print() 预览"),
        ],
        "证据图-打印三栏对照.png",
        "同一份数据与打印设置（A4 横向 / 页眉页脚 / 条件格式 / 迷你图）。Univer 侧元素由 API 重建，非 xlsx 导入。",
    )
    compose(
        [
            (os.path.join(OUT, "excel-p2.png"), "Excel：3 图表 + 形状 + 图片"),
            (os.path.join(OUT, "sjs-1-Visuals-p1.png"), "SpreadJS：3 图表 + 形状 + 图片"),
        ],
        "证据图-打印图表形状保真.png",
        "SpreadJS 侧图表/形状/图片全部出现在 PDF 中；Univer OSS 无图表与形状能力，无法构造同类页面。",
    )


if __name__ == "__main__":
    main()
