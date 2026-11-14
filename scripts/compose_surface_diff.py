# 合成"曲面图导入丢失"证据图
#   为什么需要它：样本里「雷达气泡曲面图」表原有 9 张图，导入后 SpreadJS 9 张、
#   Univer 只剩 5 张。数字写在表里没有说服力，必须把同一张表在两侧的画面并排，
#   并把"丢的是哪 4 张"标出来。
#   底图：tools/shot_sheet.js 在 25% 缩放下截的全页图（两侧同缩放、同表、无人工调整）。
import os
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'out', 'shots', '_charts-surface-diff.png')
OUT = os.path.join(ROOT, 'docs', 'assets', 'charts-surface-diff.png')

W = 1240          # 成品宽度
PAD = 26
BAR = 58          # 每段标题条高度
FONT = 'C:/Windows/Fonts/msyhbd.ttc'
FONT_R = 'C:/Windows/Fonts/msyh.ttc'

TITLE = '同一份「雷达气泡曲面图」表：SpreadJS 9 张对象 / Univer Pro 5 张对象'
SUB = '25% 缩放、同一工作表、两侧均未做手动调整。SpreadJS 那 4 张曲面图画出来是簇状柱形（官方限制，见正文）；Univer 连对象都没建'


def f(size, bold=True):
    return ImageFont.truetype(FONT if bold else FONT_R, size)


src = Image.open(SRC)
sw, sh = src.size
half = sw // 2
# 两侧面板：左 0..half，右 half..sw；纵向去掉顶部工具条与底部状态栏
top = int(sh * 0.235)
bot = int(sh * 0.755)
left = src.crop((0, top, half - 6, bot))
right = src.crop((half + 10, top, sw, bot))


def fit(img):
    w = W - PAD * 2
    h = int(img.size[1] * w / img.size[0])
    return img.resize((w, h), Image.LANCZOS)


left, right = fit(left), fit(right)

H = 96 + BAR + left.size[1] + 18 + BAR + right.size[1] + 34
canvas = Image.new('RGB', (W, H), '#ffffff')
d = ImageDraw.Draw(canvas)

d.text((PAD, 20), TITLE, font=f(28), fill='#111827')
d.text((PAD, 58), SUB, font=f(15, False), fill='#4b5563')


def panel(img, y, tag, count, color):
    d.rectangle([PAD, y, W - PAD, y + BAR], fill=color)
    d.text((PAD + 14, y + 16), tag, font=f(19), fill='#ffffff')
    d.text((W - PAD - 14 - d.textlength(count, font=f(19)), y + 16), count, font=f(19), fill='#ffffff')
    canvas.paste(img, (PAD, y + BAR))
    d.rectangle([PAD, y + BAR, PAD + img.size[0], y + BAR + img.size[1]], outline='#d1d5db', width=2)
    return y + BAR + img.size[1]


y = 96
y = panel(left, y, 'SpreadJS 19.2.2 导入：9 张对象都在（4 张曲面图被画成簇状柱形）', '9 / 9 对象', '#1d6f42')
y += 18
y = panel(right, y, 'Univer Pro 1.3.1 导入：只有 5 张（4 张曲面图对象没建）', '5 / 9 对象', '#b91c1c')

canvas.save(OUT)
print('写出', OUT, canvas.size)
