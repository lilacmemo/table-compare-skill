# -*- coding: utf-8 -*-
"""导出文件差异比对：原始 xlsx  vs  SpreadJS 导出  vs  Univer 导出

用法：
    python tools/diff_io.py                 # 比对 out/export 下全部
    python tools/diff_io.py charts          # 只比某个模块

比对维度分两层：
  A. 包结构层（zip 里有什么）：图表 XML、图片、表格、透视、切片器、批注、绘图部件
  B. 内容层（openpyxl 读得到什么）：工作表、非空单元格、公式、合并、条件格式、数据验证、打印设置

为什么要分两层：Univer/SpreadJS 都可能「内容还原但部件不生成」或反之，
只数一种会得出相反结论。
"""
import io
import json
import os
import re
import sys
import zipfile
from collections import OrderedDict

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXPD = os.path.join(ROOT, 'out', 'export')
DATA = os.path.join(ROOT, 'out', 'data')
SAMPLES = os.path.join(ROOT, 'samples')

SAMPLE_OF = {
    'charts': 'Excel原生图表效果.xlsx',
    'print': '打印-PDF-多元素样本.xlsx',
    'worksheet': '工作表与工作簿-能力域样本.xlsx',
    'format': '格式与条件格式-全量用例-A至K.xlsx',
    'table': '表-排序-筛选-验证.xlsx',
    'formula': '公式对比-Excel365基准519函数-三方-v2.xlsx',
}


def zip_parts(path):
    """包结构层：数部件。"""
    out = OrderedDict()
    if not os.path.exists(path):
        return out
    z = zipfile.ZipFile(path)
    names = z.namelist()

    def cnt(pat):
        return len([n for n in names if re.search(pat, n)])

    out['file_kb'] = round(os.path.getsize(path) / 1024.0, 1)
    out['parts'] = len(names)
    # chart 与 chartEx 都要数：新式图表（瀑布/漏斗/箱形/树状…）存在 chartExN.xml 里，
    # 只数 chartN 会把 7 张新式图漏掉，得出「原始只有 56 张」的错误基线。
    out['charts'] = cnt(r'^xl/charts/chart\d+\.xml$')
    out['chartex'] = cnt(r'^xl/charts/chartEx\d+\.xml$')
    out['chart_sheets'] = cnt(r'^xl/chartsheets/')
    out['images'] = cnt(r'^xl/media/')
    out['tables'] = cnt(r'^xl/tables/')
    out['pivots'] = cnt(r'^xl/pivot(Cache)?/')
    out['slicers'] = cnt(r'^xl/slicers/')
    out['comments'] = cnt(r'^xl/comments\d*\.xml$')
    out['drawings'] = cnt(r'^xl/drawings/drawing\d+\.xml$')
    out['vml'] = cnt(r'^xl/drawings/vmlDrawing')
    out['defined_names'] = cnt(r'^xl/definedNames')
    # 条件格式 / 验证在 sheet XML 里，用关键字计数
    cf = 0
    dv = 0
    for n in names:
        if re.match(r'^xl/worksheets/sheet\d+\.xml$', n):
            t = z.read(n).decode('utf-8', 'ignore')
            cf += t.count('<conditionalFormatting')
            dv += t.count('<dataValidation ')
    out['cond_formats'] = cf
    out['data_validations'] = dv
    z.close()
    return out


def content_stats(path):
    """内容层：openpyxl 读。只用 data_only=False，避免公式被算成值。"""
    out = OrderedDict()
    if not os.path.exists(path):
        return out
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as e:
        out['error'] = 'openpyxl 打不开: %s' % e
        return out
    sheets = []
    tot_cells = 0
    tot_formulas = 0
    tot_merges = 0
    tot_cf = 0
    tot_dv = 0
    tot_tables = 0
    for ws in wb.worksheets:
        # SpreadJS 评估版导出会多一张 "Evaluation Version" 授权说明表，不是样本内容，剔除
        if (ws.title or '').strip().lower() in ('evaluation version', 'evaluation'):
            continue
        cells = 0
        formulas = 0
        for row in ws.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                cells += 1
                if isinstance(c.value, str) and c.value.startswith('='):
                    formulas += 1
        merges = len(ws.merged_cells.ranges)
        cf = len(ws.conditional_formatting._cf_rules) if hasattr(ws.conditional_formatting, '_cf_rules') else len(list(ws.conditional_formatting))
        dv = len(ws.data_validations.dataValidation) if ws.data_validations else 0
        tables = len(ws.tables) if hasattr(ws, 'tables') else 0
        sheets.append(OrderedDict([
            ('name', ws.title),
            ('dims', '%dx%d' % (ws.max_row or 0, ws.max_column or 0)),
            ('cells', cells),
            ('formulas', formulas),
            ('merges', merges),
            ('cf', cf),
            ('dv', dv),
            ('tables', tables),
            ('orientation', ws.page_setup.orientation),
            ('fitToWidth', ws.page_setup.fitToWidth),
            ('paperSize', ws.page_setup.paperSize),
        ]))
        tot_cells += cells
        tot_formulas += formulas
        tot_merges += merges
        tot_cf += cf
        tot_dv += dv
        tot_tables += tables
    out['sheets'] = len(sheets)
    out['cells'] = tot_cells
    out['formulas'] = tot_formulas
    out['merges'] = tot_merges
    out['cond_formats'] = tot_cf
    out['data_validations'] = tot_dv
    out['tables'] = tot_tables
    out['sheet_detail'] = sheets
    wb.close()
    return out


def compare(mod):
    sample = os.path.join(SAMPLES, SAMPLE_OF.get(mod, ''))
    files = OrderedDict([
        ('原始', sample),
        ('SpreadJS', os.path.join(EXPD, mod + '-sjs.xlsx')),
        ('Univer', os.path.join(EXPD, mod + '-uni.xlsx')),
    ])
    res = OrderedDict()
    for k, f in files.items():
        res[k] = {'package': zip_parts(f), 'content': content_stats(f)}
    return res


def fmt_rows(res):
    """把结果压成 markdown 行：指标 | 原始 | SpreadJS | Univer"""
    keys_a = ['file_kb', 'charts', 'chartex', 'chart_sheets', 'images', 'tables', 'pivots', 'slicers',
              'comments', 'drawings', 'cond_formats', 'data_validations']
    keys_b = ['sheets', 'cells', 'formulas', 'merges', 'cond_formats', 'data_validations', 'tables']
    names = list(res.keys())
    rows = []
    rows.append('| 指标 | ' + ' | '.join(names) + ' |')
    rows.append('| --- | ' + ' | '.join(['---'] * len(names)) + ' |')
    rows.append('| **包结构** |  |  |  |')
    for k in keys_a:
        vals = []
        for n in names:
            vals.append(str(res[n]['package'].get(k, '-')))
        rows.append('| %s | %s |' % (k, ' | '.join(vals)))
    rows.append('| **内容** |  |  |  |')
    for k in keys_b:
        vals = []
        for n in names:
            vals.append(str(res[n]['content'].get(k, '-')))
        rows.append('| %s | %s |' % (k, ' | '.join(vals)))
    return '\n'.join(rows)


def main():
    mods = sys.argv[1:] or list(SAMPLE_OF.keys())
    allres = OrderedDict()
    os.makedirs(DATA, exist_ok=True)
    for mod in mods:
        if not os.path.exists(os.path.join(EXPD, mod + '-sjs.xlsx')) and not os.path.exists(os.path.join(EXPD, mod + '-uni.xlsx')):
            print('[skip] %s 无导出文件' % mod)
            continue
        r = compare(mod)
        allres[mod] = r
        md = fmt_rows(r)
        with io.open(os.path.join(DATA, mod + '-diff.md'), 'w', encoding='utf-8') as f:
            f.write(md + '\n')
        print('\n=== %s ===' % mod)
        print(md)
    with io.open(os.path.join(DATA, '_diff-all.json'), 'w', encoding='utf-8') as f:
        f.write(json.dumps(allres, ensure_ascii=False, indent=1))


if __name__ == '__main__':
    main()
