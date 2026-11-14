# -*- coding: utf-8 -*-
"""从 SpreadJS 官方《功能清单》xlsx 反推测试项。

用法:
  python extract_feature_inventory.py <功能清单.xlsx>                 # 打印模块统计
  python extract_feature_inventory.py <xlsx> --module 05.单元格       # 导出该模块全部条目(md 表格)
  python extract_feature_inventory.py <xlsx> --grep 格式,填充,条件格式  # 按特性/功能名筛选
  python extract_feature_inventory.py <xlsx> --module 05.单元格 --out out.md

列结构（v19.2 清单，第 5 行是表头，数据从第 6 行起）:
  A 版本 | B 版本号(V15.0/V19.2 新增) | C 模块中文 | D 模块En | E 特性中文
  F 特性En | G 模块>功能路径 | H 功能中文 | I 功能En | J 全称
  K~M 训练营/中级/高级认证等级 | N~P 数据填报/类Excel报表/表格协同 场景权重 | Q 功能简述
"""
import argparse
import sys

import openpyxl

HDR_ROW = 5
COLS = {
    'ver_tag': 1, 'module': 2, 'module_en': 3, 'feature': 4, 'feature_en': 5,
    'path': 6, 'func': 7, 'func_en': 8, 'full': 9, 'desc': 16,
}


def load_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=HDR_ROW + 1, values_only=True):
        if not r or not (r[COLS['module']] or '').strip():
            continue
        item = {k: (r[i] if i < len(r) else None) for k, i in COLS.items()}
        rows.append(item)
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('xlsx')
    ap.add_argument('--module', default='', help='模块名，支持前缀匹配，如 05 或 05.单元格')
    ap.add_argument('--grep', default='', help='在特性/功能名里搜关键词，逗号分隔，任一命中')
    ap.add_argument('--out', default='')
    args = ap.parse_args()

    rows = load_rows(args.xlsx)

    if not args.module and not args.grep:
        seen = {}
        for it in rows:
            m = (it['module'] or '').strip()
            seen[m] = seen.get(m, 0) + 1
        print('模块统计（共 %d 条）:' % len(rows))
        for m, n in seen.items():
            print('  %-16s %3d' % (m, n))
        return

    sel = rows
    if args.module:
        sel = [it for it in sel if (it['module'] or '').startswith(args.module)]
    if args.grep:
        keys = [k.strip() for k in args.grep.split(',') if k.strip()]
        sel = [it for it in sel if any(
            k in ((it['feature'] or '') + (it['func'] or '') + (it['full'] or '')) for k in keys)]

    lines = ['| 版本 | 模块 | 特性 | 功能 | 完整路径 | 功能简述 |', '| --- | --- | --- | --- | --- | --- |']
    for it in sel:
        desc = (it['desc'] or '').replace('\n', ' ')[:100]
        lines.append('| %s | %s | %s | %s | %s | %s |' % (
            it['ver_tag'] or '', it['module'] or '', it['feature'] or '',
            it['func'] or '', it['path'] or '', desc))
    out = '\n'.join(lines)
    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            f.write(out + '\n')
        print('written %s (%d items)' % (args.out, len(sel)))
    else:
        print(out)


if __name__ == '__main__':
    sys.exit(main())
