# 从样本 xlsx 里抽出「公式用例 + Excel oracle」，落成 JSON 供工程内复算使用。
# 为什么不在浏览器里读：对比表的 E 列是活公式（=Excel实测!Z2），引擎导入后会重算，
# 拿引擎自己算出来的值当标准答案就是循环论证。Excel 缓存的 F 列文本才是真 oracle。
import json, sys
from openpyxl import load_workbook

SRC = 'samples/公式对比-Excel365基准519函数-三方-v2.xlsx'
OUT = 'js/data/formula-oracle.json'

wb = load_workbook(SRC, data_only=True)
ws = wb['对比表']
rows = []
for r in range(5, ws.max_row + 1):
    name = ws.cell(row=r, column=2).value
    formula = ws.cell(row=r, column=4).value
    exp = ws.cell(row=r, column=6).value   # Excel 结果文本（缓存值）
    if not name or not formula:
        continue
    formula = str(formula).strip()
    if not formula.startswith('='):
        continue
    rows.append({'id': str(name).strip(), 'formula': formula[1:], 'exp': '' if exp is None else str(exp)})

json.dump({'source': SRC.split('/')[-1], 'count': len(rows), 'rows': rows},
          open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('写出', OUT, len(rows), '条')
print('样例', json.dumps(rows[:3], ensure_ascii=False))
