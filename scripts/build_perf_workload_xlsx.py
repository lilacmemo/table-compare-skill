# 生成《性能基准-工作负载样本.xlsx》：把 perf_workload.js（种子 42）生成的负载落成 xlsx，供查看/核对
# 用法：python build_perf_workload_xlsx.py    （需同目录有 perf_workload.js，且已装 openpyxl）
import json, re, io, os, subprocess
from openpyxl import Workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), '性能基准-工作负载样本.xlsx')

# 1) 用 node 调 perf_workload.js 生成确定性负载，落临时 JSON
tmp = os.path.join(HERE, 'tmp_workload_dump.json')
subprocess.run(['node', '-e',
    "const w=require('./perf_workload');require('fs').writeFileSync('tmp_workload_dump.json',JSON.stringify(w.generate()))"],
    cwd=HERE, check=True)
d = json.load(io.open(tmp, encoding='utf-8'))

def col_to_idx(letters):
    n = 0
    for ch in letters: n = n * 26 + ord(ch) - 64
    return n - 1

wb = Workbook()
ws = wb.active
ws.title = '说明'
for r in [
    ['性能基准工作负载样本（SpreadJS vs Univer 公式计算性能测试所用文件）'],
    [''],
    ['生成方式', 'perf_workload.js（确定性随机种子 42），与双端性能实测写入的公式完全一致'],
    ['公式总数', '43,259 条'],
    ['Data', '10,000 行 × 4 公式列（F 金额=D*E、G 税后=ROUND(F*(1+税率))、H VLOOKUP、I 嵌套 IF 评级），首行不计 → 39,996 条'],
    ['Report', '120 SUMIFS + 120 SUMPRODUCT（扫描 Data 10,000 行数据区）→ 240 条'],
    ['Chain', '1,000 行 × 3 链式依赖公式（逐行传递累计）→ 3,000 条'],
    ['Dynamic', 'UNIQUE/SORT/FILTER 溢出 + SUMIFS → 23 条动态数组公式'],
    [''],
    ['注', '本文件无缓存值，Excel 打开时自动全量重算（已设 fullCalcOnLoad）；性能实测是双端引擎内逐条写入，本文件仅用于查看/核对负载内容。'],
]: ws.append(r)
ws.column_dimensions['A'].width = 12
ws.column_dimensions['B'].width = 110

# Data：值 + F..I 公式（与 perf_sjs.html / perf_univer.html 的写入口径一致：跳过第 1 行）
ws = wb.create_sheet('Data')
for r, row in enumerate(d['dataValues']):
    for c, v in enumerate(row):
        if v is not None: ws.cell(r + 1, c + 1, v)
n_data = 0
for r in range(1, len(d['dataFormulas'])):
    row = d['dataFormulas'][r]
    for c in range(5, 9):
        if row[c]:
            ws.cell(r + 1, c + 1, row[c]); n_data += 1

ws = wb.create_sheet('Params')
for row in d['paramsValues']: ws.append(row)
ws = wb.create_sheet('Products')
for row in d['productsValues']: ws.append(row)

ws = wb.create_sheet('Report')
for r, row in enumerate(d['reportValues']):
    for c, v in enumerate(row):
        if v is not None: ws.cell(r + 1, c + 1, v)
r0 = len(d['reportValues'])
for row in d['reportPart2']:
    r0 += 1
    for c, v in enumerate(row):
        if v is not None: ws.cell(r0, c + 1, v)
n_report = 0
for f in d['reportFormulas']:
    m = re.match(r'([A-Z]+)(\d+)', f['cell'])
    ws.cell(int(m.group(2)), col_to_idx(m.group(1)) + 1, f['formula']); n_report += 1

ws = wb.create_sheet('Chain')
for row in d['chainValues']: ws.append(row)
n_chain = 0
for f in d['chainFormulas']:
    m = re.match(r'([A-Z]+)(\d+)', f['cell'])
    ws.cell(int(m.group(2)), col_to_idx(m.group(1)) + 1, f['formula']); n_chain += 1

ws = wb.create_sheet('Dynamic')
n_dyn = 0
for f in d['dynFormulas']:
    m = re.match(r'([A-Z]+)(\d+)', f['cell'])
    ws.cell(int(m.group(2)), col_to_idx(m.group(1)) + 1, f['formula']); n_dyn += 1

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
os.remove(tmp)
print('saved', OUT, '| formulas:', n_data, n_report, n_chain, n_dyn, 'total', n_data + n_report + n_chain + n_dyn)
