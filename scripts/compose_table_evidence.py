# 把 show-sjs.png / show-uni.png 拼成一张对照证据图
#   python tools/compose_table_evidence.py
from PIL import Image, ImageDraw, ImageFont
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, 'out')
DST = os.path.join(HERE, '证据图-表模块对照.png')

FONT = 'C:/Windows/Fonts/msyh.ttc'


def f(size, bold=False):
    path = 'C:/Windows/Fonts/msyhbd.ttc' if bold else FONT
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.truetype(FONT, size)


a = Image.open(os.path.join(OUT, 'show-sjs.png'))
b = Image.open(os.path.join(OUT, 'show-uni.png'))

# 只保留表格所在区域（2 倍图，裁掉右侧空白与底部大段空行）
CROP = (0, 0, 1010, 566)
a = a.crop(CROP)
b = b.crop(CROP)

GAP = 36
PAD = 34
HEAD = 168
FOOT = 168
W = PAD * 2 + a.width + GAP + b.width
H = HEAD + a.height + FOOT

canvas = Image.new('RGB', (W, H), '#ffffff')
d = ImageDraw.Draw(canvas)

d.text((PAD, 26), '表模块对照：两侧跑同一套代码（建表 + 表样式 + 计算列 + 筛选按钮 + 汇总行）',
       font=f(34, True), fill='#1a1a1a')
d.text((PAD, 74), '样本：区域/产品/数量/单价/金额 5 列 4 行；金额列 =[@数量]*[@单价]；末行 总计 = SUBTOTAL(109,[金额])',
       font=f(22), fill='#5b6470')

# 面板标题条
y = HEAD - 44
d.rectangle([PAD, y, PAD + a.width, y + 34], fill='#eef3fb')
d.text((PAD + 12, y + 5), 'SpreadJS 19.2.5', font=f(22, True), fill='#1f4e79')
x2 = PAD + a.width + GAP
d.rectangle([x2, y, x2 + b.width, y + 34], fill='#eef7f1')
d.text((x2 + 12, y + 5), 'Univer OSS 1.3.1', font=f(22, True), fill='#1e6b45')

canvas.paste(a, (PAD, HEAD))
canvas.paste(b, (x2, HEAD))

for px in (PAD, x2):
    d.rectangle([px, HEAD, px + a.width - 1, HEAD + a.height - 1], outline='#c9d2dc', width=2)

# 底部结论条
fy = HEAD + a.height + 22
rows = [
    ('镶边行 / 首列高亮', '左侧有 bandRows()/highlightFirstColumn() 开关；右侧 ITableOptions 无字段 → 只能靠样式 ID 的整体配色'),
    ('汇总行', '左侧 showFooter+SUBTOTAL 出「总计 375」；右侧 hasTotalRow 标 @todo，表只有 5 行、无汇总行'),
    ('筛选按钮', '两侧都有（Univer 需同时注册 sheets-filter 与 sheets-filter-ui，只装前者列头不画漏斗）'),
]
d.line([PAD, fy - 8, W - PAD, fy - 8], fill='#dde3ea', width=2)
yy = fy
for k, v in rows:
    d.text((PAD, yy), k, font=f(21, True), fill='#1a1a1a')
    d.text((PAD + 230, yy), v, font=f(20), fill='#3d4a57')
    yy += 44

canvas.save(DST)
print('已生成', DST, canvas.size)
