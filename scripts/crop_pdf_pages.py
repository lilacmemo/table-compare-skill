# -*- coding: utf-8 -*-
"""把 _<prefix>-raw-pN.png 里的白纸页面裁出来 → _<prefix>-pN.png

PDF 查看器（Chromium 内置）背景是灰色、页面是白纸，用"最大近白矩形"就能定位页面本体。
"""
import os
import sys

import numpy as np
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHOTS = os.path.join(ROOT, 'out', 'shots')


def crop_one(path):
    im = Image.open(path).convert('RGB')
    a = np.asarray(im, dtype=np.int16)
    near_white = (a.min(axis=2) > 246)
    rows = np.where(near_white.sum(axis=1) > a.shape[1] * 0.15)[0]
    cols = np.where(near_white.sum(axis=0) > a.shape[0] * 0.15)[0]
    if not len(rows) or not len(cols):
        return None
    box = (int(cols[0]), int(rows[0]), int(cols[-1]) + 1, int(rows[-1]) + 1)
    return im.crop(box)


if __name__ == '__main__':
    prefix = sys.argv[1]
    n = 0
    for i in range(1, 9):
        src = os.path.join(SHOTS, '_%s-raw-p%d.png' % (prefix, i))
        if not os.path.exists(src):
            break
        out = crop_one(src)
        if out is None:
            print('[skip]', os.path.basename(src))
            continue
        dst = os.path.join(SHOTS, '_%s-p%d.png' % (prefix, i))
        out.save(dst)
        print('crop', os.path.basename(dst), out.size)
        n += 1
    print('共', n, '页')
