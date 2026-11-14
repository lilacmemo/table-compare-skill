# -*- coding: utf-8 -*-
"""把两张导入截图拼成一张带标注的证据图

用法：
  python tools/compose_io_figure.py                 # 导入侧·左右并排 → out/shots/<模块>-对比图.png
  python tools/compose_io_figure.py --doc           # 导入侧·竖排文档版 → docs/assets/<模块>-io-compare.png
  python tools/compose_io_figure.py --rt            # 导出侧（往返）·并排 → out/shots/<模块>-对比图-往返.png
  python tools/compose_io_figure.py --rt --doc      # 导出侧·竖排文档版 → docs/assets/<模块>-io-roundtrip.png
  python tools/compose_io_figure.py charts print    # 只处理指定模块

为什么要两个环节、两个尺寸：
  环节：导入后的默认视图只说明"读进来对不对"；还要把各家导出的 xlsx 重新打开截一张，
        才证明"写出去的文件本身还能读"。两者合起来才是完整的导入 / 导出兼容性。
  尺寸：并排版宽 1696px，塞进 md 正文（按 800px 显示）时每侧只剩 400px，表里的字全糊了。
        文档版改成上下竖排、单侧 1100px，缩放后仍能看清，且文件名全 ASCII，
        转 Word / 飞书 / GitHub 时不会因为中文路径丢图。
"""
import io
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOTS = os.path.join(ROOT, 'out', 'shots')
DATA = os.path.join(ROOT, 'out', 'data')
ASSETS = os.path.join(ROOT, 'docs', 'assets')

FONTS = [
    'C:/Windows/Fonts/msyh.ttc',
    'C:/Windows/Fonts/msyhbd.ttc',
    'C:/Windows/Fonts/simhei.ttf',
]


def font(size, bold=False):
    cands = (['C:/Windows/Fonts/msyhbd.ttc'] if bold else []) + FONTS
    for f in cands:
        if os.path.exists(f):
            try:
                return ImageFont.truetype(f, size)
            except Exception:
                pass
    return ImageFont.load_default()


TITLES = {
    'charts': '图表导入还原',
    'print': '打印与 PDF 元素还原',
    'worksheet': '工作表 / 工作簿能力域还原',
    'format': '格式与条件格式还原',
    'table': '表 / 排序 / 筛选 / 验证还原',
    'formula': '公式与计算结果还原',
}

SUB_IMPORT = '同一份 xlsx 导入后的默认视图（未做任何手动调整）'
SUB_RT = '各家导出的 xlsx 重新打开后的视图（导出 → 再导入，未做任何手动调整）'


def load_meta(mod):
    p = os.path.join(DATA, mod + '-io.json')
    if not os.path.exists(p):
        return None
    try:
        return json.load(io.open(p, encoding='utf-8'))
    except Exception:
        return None


def caption(mod, meta):
    if not meta:
        return ''
    s = (meta.get('stats', {}).get('sjs') or {}).get('totals') or {}
    u = (meta.get('stats', {}).get('uni') or {}).get('totals') or {}
    sjs_ms = (meta.get('importMs', {}).get('sjs') or {}).get('ms')
    uni_ms = (meta.get('importMs', {}).get('uni') or {}).get('ms')
    parts = []
    if s:
        parts.append('SJS 表 %s / 图 %s / 公式 %s / 导入 %sms'
                     % (s.get('sheets', '-'), s.get('charts', '-'), s.get('formulas', '-'), sjs_ms or '-'))
    if u:
        parts.append('Univer 表 %s / 图 %s / 公式 %s / 导入 %sms'
                     % (u.get('sheets', '-'), u.get('charts', '-'), u.get('formulas', '-'), uni_ms or '-'))
    return '　|　'.join(parts)


def compose(mod, doc=False, kind='import'):
    """kind: 'import' 用导入后的截图；'rt' 用导出文件重新打开后的截图"""
    sfx = '-rt' if kind == 'rt' else ''
    a = os.path.join(SHOTS, mod + '-sjs' + sfx + '.png')
    b = os.path.join(SHOTS, mod + '-uni' + sfx + '.png')
    if not (os.path.exists(a) and os.path.exists(b)):
        print('[skip]', mod, kind, '缺截图', os.path.basename(a))
        return None
    ia, ib = Image.open(a).convert('RGB'), Image.open(b).convert('RGB')

    W = 1100 if doc else 820

    def fit(im):
        w, h = im.size
        nh = int(h * W / w)
        return im.resize((W, nh), Image.LANCZOS)
    ia, ib = fit(ia), fit(ib)

    GAP, PAD, HEAD, FOOT = (18 if doc else 16), 20, 62, 58
    LABEL_H = 26
    meta = load_meta(mod)
    cap = caption(mod, meta)

    title = 'SpreadJS vs Univer Pro · ' + TITLES.get(mod, mod)
    sub = SUB_RT if kind == 'rt' else SUB_IMPORT
    if kind == 'rt':
        # 页脚换口径：往返验证关心的是"再打开后对象还在不在"
        rt = (meta or {}).get('roundtrip') or {}
        bits = []
        for side, label in (('sjs', 'SJS'), ('uni', 'Univer')):
            r = rt.get(side) or {}
            t = r.get('totals') or {}
            if r.get('error'):
                bits.append('%s 重开失败：%s' % (label, r['error'][:40]))
            elif t:
                bits.append('%s 重开 %sms / 表 %s / 图 %s / 公式 %s'
                            % (label, r.get('ms', '-'), t.get('sheets', '-'),
                               t.get('charts', '-'), t.get('formulas', '-')))
        if bits:
            cap = '　|　'.join(bits)

    if doc:
        # 竖排：标题 → [标签 + SJS] → [标签 + Univer] → 页脚
        H = HEAD + PAD + (LABEL_H + ia.size[1]) + GAP + (LABEL_H + ib.size[1]) + (FOOT + PAD if cap else PAD)
        canvas = Image.new('RGB', (W + PAD * 2, H), '#eceef3')
        d = ImageDraw.Draw(canvas)
        d.rectangle([0, 0, canvas.size[0], HEAD], fill='#1b2433')
        d.text((PAD, 16), title, font=font(22, True), fill='#ffffff')
        d.text((PAD, 42), sub, font=font(13), fill='#9aa8c0')

        def block(img, ytop, text, color):
            tw = d.textlength(text, font=font(13, True))
            d.rectangle([PAD, ytop, PAD + tw + 18, ytop + LABEL_H - 4], fill=color)
            d.text((PAD + 9, ytop + 4), text, font=font(13, True), fill='#ffffff')
            canvas.paste(img, (PAD, ytop + LABEL_H))
            d.rectangle([PAD - 1, ytop + LABEL_H - 1, PAD + W, ytop + LABEL_H + img.size[1]],
                        outline='#c9d0dc')

        y = HEAD + PAD
        block(ia, y, 'SpreadJS 19.2.5', '#b91c1c')
        y += LABEL_H + ia.size[1] + GAP
        block(ib, y, 'Univer Pro 1.3.1（已授权）', '#1d4ed8')
        y += LABEL_H + ib.size[1]
        if cap:
            d.text((PAD, y + 12), cap, font=font(12), fill='#48536a')

        if not os.path.isdir(ASSETS):
            os.makedirs(ASSETS)
        out = os.path.join(ASSETS, mod + ('-io-roundtrip.png' if kind == 'rt' else '-io-compare.png'))
    else:
        H = max(ia.size[1], ib.size[1])
        canvas = Image.new('RGB', (W * 2 + GAP + PAD * 2, HEAD + H + FOOT + PAD * 2), '#eceef3')
        d = ImageDraw.Draw(canvas)
        d.rectangle([0, 0, canvas.size[0], HEAD], fill='#1b2433')
        d.text((PAD, 16), title, font=font(22, True), fill='#ffffff')
        d.text((PAD, 42), sub, font=font(13), fill='#9aa8c0')

        x1, x2 = PAD, PAD + W + GAP
        y = HEAD + PAD + LABEL_H
        canvas.paste(ia, (x1, y))
        canvas.paste(ib, (x2, y))

        def tag(x, text, color):
            tw = d.textlength(text, font=font(13, True))
            d.rectangle([x, y - LABEL_H, x + tw + 18, y - 4], fill=color)
            d.text((x + 9, y - LABEL_H + 4), text, font=font(13, True), fill='#ffffff')
        tag(x1, 'SpreadJS 19.2.5', '#b91c1c')
        tag(x2, 'Univer Pro 1.3.1（已授权）', '#1d4ed8')

        d.rectangle([x1 - 1, y - 1, x1 + W, y + ia.size[1]], outline='#c9d0dc')
        d.rectangle([x2 - 1, y - 1, x2 + W, y + ib.size[1]], outline='#c9d0dc')

        if cap:
            d.text((PAD, y + H + 12), cap, font=font(12), fill='#48536a')

        out = os.path.join(SHOTS, mod + ('-对比图-往返.png' if kind == 'rt' else '-对比图.png'))

    canvas.save(out, optimize=True)
    print('写出', os.path.relpath(out, ROOT), canvas.size)
    return out


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    doc = '--doc' in sys.argv
    kind = 'rt' if '--rt' in sys.argv else 'import'
    mods = args or ['charts', 'print', 'worksheet', 'format', 'table', 'formula']
    for m in mods:
        compose(m, doc=doc, kind=kind)
