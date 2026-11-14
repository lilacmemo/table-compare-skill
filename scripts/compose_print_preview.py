# -*- coding: utf-8 -*-
"""打印预览 / 输出 三方对照图（基准 = Excel 本体）

布局：三行竖排
  ① Excel 打印输出第 1 页（ExportAsFixedFormat 出的 PDF 渲染）—— 判定基准
  ② SpreadJS Designer 文件 → 打印 面板（含预览）
  ③ Univer Pro 打印对话框（含预览）

为什么把 Excel 放第一行：这个项目所有比对都以 Excel 为基准线。打印这一块尤其明显——
"页面设置有没有被读进来"只有拿 Excel 自己的 PageSetup 值才能判，拿
"SJS 读了什么 / Univer 读了什么" 互相比是没有基准的。

用法：python tools/compose_print_preview.py
"""
import io
import json
import os

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOTS = os.path.join(ROOT, 'out', 'shots')
DATA = os.path.join(ROOT, 'out', 'data')
ASSETS = os.path.join(ROOT, 'docs', 'assets')

FONTS = ['C:/Windows/Fonts/msyhbd.ttc', 'C:/Windows/Fonts/msyh.ttc', 'C:/Windows/Fonts/simhei.ttf']


def font(size, bold=False):
    cands = (['C:/Windows/Fonts/msyhbd.ttc'] if bold else []) + FONTS
    for f in cands:
        if os.path.exists(f):
            try:
                return ImageFont.truetype(f, size)
            except Exception:
                pass
    return ImageFont.load_default()


def trim_white(im, pad=18, tol=8, max_frac=1.0):
    """裁掉 UI 截图四周的大片空白（两块面板下半部全是白），否则拼出来的图一半是空气。
       只对界面截图用；Excel 的整页输出不能裁——页边距本身就是证据。
       坑：面板本身有 1px 边框（#d3d8e2），直接算 bbox 会命中边框、等于没裁，
       所以先往里缩 4px 再求内容范围。"""
    import numpy as np
    a = np.asarray(im.convert('RGB'), dtype=np.int16)
    h, w = a.shape[:2]
    inset = 4
    # 只用右半边判断"内容到哪儿结束"：两块面板左侧都有一条贯穿到底的分栏竖线，
    # 算进来就等于没裁。宽度保留完整（左栏菜单是证据的一部分），只收高度。
    inner = a[inset:h - inset, int(w * 0.35):w - inset]
    nonwhite = (np.abs(inner - 255).max(axis=2) > tol)
    rows = np.where(nonwhite.any(axis=1))[0]
    if not len(rows):
        return im
    # 两块面板底部都有各自的底栏（Designer 的状态栏 / Univer 的缩放按钮），
    # 它们撑到底、且对打印结论没有价值，所以再压一道上限
    b = min(h, int(h * max_frac), rows[-1] + inset + pad)
    return im.crop((0, 0, w, b))


def load_oracle():
    p = os.path.join(DATA, 'print-excel-oracle.json')
    if not os.path.exists(p):
        return None
    try:
        return json.load(io.open(p, encoding='utf-8-sig'))
    except Exception:
        return None


def compose():
    rows = [
        ('excel-print-p1.png', 'Excel（基准）· 打印输出第 1 页', '#1d6f42',
         '横向 A4（842×595pt）· 共 1 页 · 页眉三栏（&8 机密 / &"Arial,Bold"&11 季度销售报表 / &D）· 页脚三栏（由 Excel 生成 / 第 &P 页 共 &N 页 / &F）'),
        ('print-ui-sjs.png', 'SpreadJS Designer 19.2.2 · 文件 → 打印', '#b91c1c',
         '横向 · A4 8.27"×11.69" · 将所有列调整为一页 · 自定义页边距 上0.7 下0.7 左0.5 右0.5 · 预览「第 1 页 / 共 1 页」，页眉页脚与 Excel 逐字一致'),
        ('print-ui-uni.png', 'Univer Pro 1.3.1 · 打印对话框', '#1d4ed8',
         '纸张 A4 一致；其余回落默认：纵向 · 缩放 100% · 边距「正常」· 页眉页脚全空 · 网格线开 → 同一张表被拆成 2 页'),
    ]
    imgs = []
    for fn, title, color, note in rows:
        p = os.path.join(SHOTS, fn)
        if not os.path.exists(p):
            print('[skip] 缺图', fn)
            return None
        im = Image.open(p).convert('RGB')
        # Excel 那一行是整页 PDF 输出，页边距是证据，保留全页；两家 UI 面板裁掉空白
        if 'print-ui-' in fn:
            im = trim_white(im, max_frac=0.74)
        imgs.append((im, title, color, note))

    W = 1180
    GAP, PAD, HEAD = 20, 24, 74
    TAG_H, NOTE_H = 30, 46

    fitted = []
    for im, title, color, note in imgs:
        w, h = im.size
        nh = int(h * W / w)
        fitted.append((im.resize((W, nh), Image.LANCZOS), title, color, note))

    H = HEAD + PAD
    for im, _, _, note in fitted:
        H += TAG_H + im.size[1] + (NOTE_H if note else 0) + GAP
    H += PAD

    canvas = Image.new('RGB', (W + PAD * 2, H), '#eceef3')
    d = ImageDraw.Draw(canvas)

    d.rectangle([0, 0, canvas.size[0], HEAD], fill='#1b2433')
    d.text((PAD, 16), 'SpreadJS vs Univer Pro · 打印预览与打印输出（基准：Excel 本体）', font=font(23, True), fill='#ffffff')
    d.text((PAD, 46), '同一份《打印-PDF-多元素样本.xlsx》：Excel 自己打印的页面 → 两家打开后看到的打印界面',
           font=font(13), fill='#9aa8c0')

    y = HEAD + PAD
    for im, title, color, note in fitted:
        tw = d.textlength(title, font=font(14, True))
        d.rectangle([PAD, y, PAD + tw + 20, y + TAG_H - 6], fill=color)
        d.text((PAD + 10, y + 5), title, font=font(14, True), fill='#ffffff')
        yy = y + TAG_H
        canvas.paste(im, (PAD, yy))
        d.rectangle([PAD - 1, yy - 1, PAD + W, yy + im.size[1]], outline='#c9d0dc')
        if note:
            d.text((PAD, yy + im.size[1] + 10), note, font=font(13), fill='#3c4658')
        y += TAG_H + im.size[1] + (NOTE_H if note else 0) + GAP

    if not os.path.isdir(ASSETS):
        os.makedirs(ASSETS)
    out = os.path.join(ASSETS, 'print-preview-compare.png')
    canvas.save(out, optimize=True)
    print('写出', os.path.relpath(out, ROOT), canvas.size)
    return out


if __name__ == '__main__':
    compose()
