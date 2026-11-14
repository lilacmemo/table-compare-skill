# -*- coding: utf-8 -*-
"""从图表样本里派生「只有一张图」的变体工作簿（每张图一份）。

为什么要派生：Univer 侧没法在不破坏渲染的前提下把邻居图藏起来 ——
  · chart.remove() 掉其余 90 张后，幸存的那张**不再绘制**（getCharts() 还在、size/position 都在，画布上就是没有）；
  · chart.setAbsolutePosition() 传负坐标会被夹到 0，90 张图全叠到左上角，剪裁框里一片花。
而 SpreadJS 侧已经改用 chart.toImageSrcAsync()（图表控件自己渲染），根本不需要邻居离场。
所以只有 Univer 需要「单图载体」，做法是在 drawing1.xml 里只留该图对应的那个 twoCellAnchor。

顺带把样本自带的灰色说明文本框（每个 case 一条 Shapes.AddTextbox）一起丢掉 —— 它压在图上沿，
Excel 的 Chart.Export 却不含它，留着三方图就对不齐。

用法：python tools/build_single_chart_variants.py
输出：out/variants/{caseId}.xlsx + out/data/chart-variants.json
"""
import json
import os
import re
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE = os.path.join(ROOT, 'samples', '图表用例载体-逐项配置.xlsx')
OUTDIR = os.path.join(ROOT, 'out', 'variants')
DRAWING = 'xl/drawings/drawing1.xml'
# 「用例」表的 dimension 原本只有 A1（表里只有表头，图全是浮动对象）。Univer 按 dimension 决定
# 工作表行数（默认 150 行），而图表锚点有落在第 193 / 543 / 894 行的 —— 单图变体没有「多图撑开行数」
# 的副作用，这些图就会被夹到最后一行、高度退化成 1px。所以派生变体时把 dimension 撑到 A1:AE1000。
DIM_MIN_ROWS = 1000
DIM_MIN_COLS = 31

ANCHOR = re.compile(r'<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>', re.S)
NAME = re.compile(r'<xdr:cNvPr[^>]*\bname="([^"]+)"')


def main():
    os.makedirs(OUTDIR, exist_ok=True)
    zin = zipfile.ZipFile(SAMPLE)
    # 一次读完所有成员再写：Python 3.13 的 zipfile 在同一句柄上「读一次→写别的 zip→再读」会抛
    # BadZipFile: Bad CRC-32（写 ZipInfo 复用了原 CRC），提前缓存成 bytes 就绕开了
    members = [(it, zin.read(it.filename)) for it in zin.infolist()]
    sheet_with_drawing = [it.filename for it, b in members
                          if it.filename.startswith('xl/worksheets/sheet') and b'drawing' in b][0]
    zin.close()
    xml = dict((it.filename, b) for it, b in members)[DRAWING].decode('utf-8')
    anchors = ANCHOR.findall(xml)
    head = xml[:xml.find(anchors[0])]
    tail = xml[xml.rfind(anchors[-1]) + len(anchors[-1]):]

    charts = []
    for a in anchors:
        if '<c:chart' in a:            # 图表；剩下的都是 <xdr:sp>（样本自带的灰底说明框）
            m = NAME.search(a)
            charts.append((m.group(1) if m else '?', a))
    print('anchors', len(anchors), 'charts', len(charts), 'textboxes', len(anchors) - len(charts))

    index = []
    for cid, anchor in charts:
        body = head + anchor + tail
        out = os.path.join(OUTDIR, '%s.xlsx' % cid)
        tmp = out + '.tmp'
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for it, data in members:
                if it.filename == DRAWING:
                    data = body.encode('utf-8')
                elif it.filename == sheet_with_drawing:
                    data = re.sub(rb'<dimension ref="[^"]*"',
                                  b'<dimension ref="A1:AE%d"' % DIM_MIN_ROWS, data, count=1)
                zout.writestr(zipfile.ZipInfo(it.filename, it.date_time), data, zipfile.ZIP_DEFLATED)
        os.replace(tmp, out)
        index.append({'id': cid, 'file': orel(out)})

    with open(os.path.join(ROOT, 'out', 'data', 'chart-variants.json'), 'w', encoding='utf-8') as f:
        json.dump(index, f, ensure_ascii=False, indent=1)
    print('wrote', len(index), 'variants ->', OUTDIR)
    print('样例', index[0], index[-1])


def orel(p):
    return os.path.relpath(p, ROOT).replace('\\', '/')


if __name__ == '__main__':
    main()
