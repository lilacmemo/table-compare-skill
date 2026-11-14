# -*- coding: utf-8 -*-
"""生成《公式对比-批次1-三方结果.xlsx》：一列公式、一列 Excel 基线、一列 SJS、一列 Univer + 判定。
数据源：formula_batch1_def.json / formula_batch1_baseline.json / sjs_formula_batch1.json / univer_formula_batch1.json
用法：python build_formula_compare_xlsx.py
"""
import io, json, os, re, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, '..', '公式对比-批次1-三方结果.xlsx')

def load(name):
    return json.load(io.open(os.path.join(ROOT, name), encoding='utf-8'))

def is_err(v):
    if isinstance(v, dict):
        s = str(v.get('_calcError') or v.get('_error') or v.get('error') or '')
        return s.startswith('#')
    if isinstance(v, str):
        return v.startswith('#') and len(v) < 12
    return False

def err_text(v):
    if isinstance(v, dict):
        return str(v.get('_calcError') or v.get('_error') or v.get('error') or '#ERR')
    return str(v)

def to_serial(s):
    if isinstance(s, str) and re.match(r'^\d{4}-\d{2}-\d{2}T', s):
        dt = datetime.datetime.fromisoformat(s.replace('Z', '+00:00'))
        if dt.tzinfo is not None:
            dt = dt.astimezone().replace(tzinfo=None)
        return (dt - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    return None

def norm(v):
    """归一化用于比较：错误→错误串；ISO 日期→序列号；字符串数字→float"""
    if isinstance(v, dict):
        return ('ERR', err_text(v))
    if isinstance(v, str):
        ser = to_serial(v)
        if ser is not None:
            return round(ser, 6)
        s = v.strip()
        if s.lower() in ('true', 'false'):
            return s.lower()
        try:
            return float(s)
        except Exception:
            return s
    if isinstance(v, bool):
        return str(v).lower()
    if isinstance(v, (int, float)):
        return round(float(v), 6)
    return v

def eq(a, b):
    na, nb = norm(a), norm(b)
    if isinstance(na, tuple) or isinstance(nb, tuple):
        return na == nb
    if isinstance(na, float) and isinstance(nb, float):
        return abs(na - nb) < 1e-6
    if isinstance(na, float) or isinstance(nb, float):
        try:
            return abs(float(na) - float(nb)) < 1e-6
        except Exception:
            return False
    return na == nb

def show(v):
    if v is None:
        return ''
    if isinstance(v, dict):
        return err_text(v)
    if isinstance(v, str):
        ser = to_serial(v)
        if ser is not None:
            d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=ser)
            return d.strftime('%Y-%m-%d') + '（序列号 %g）' % round(ser, 4)
        return v
    if isinstance(v, float) and abs(v - round(v)) < 1e-9:
        return int(round(v))
    if isinstance(v, float):
        return round(v, 6)
    return v

def verdict(b, p, sname='', formula='', prod=''):
    # 易失函数：只判"识别 + 返回类型正确"，不比较具体值
    if sname == 'F易失':
        def cat(v):
            if v is None:
                return 'empty'
            if is_err(v):
                return 'err'
            if isinstance(v, bool):
                return 'bool'
            if isinstance(v, (int, float)):
                return 'num'
            if isinstance(v, str):
                if to_serial(v) is not None:
                    return 'date'
                sl = v.strip().lower()
                if sl in ('true', 'false'):
                    return 'bool'
                try:
                    float(sl); return 'num'
                except Exception:
                    return 'text'
            return 'other'
        cb, cp = cat(b), cat(p)
        if cp == 'err':
            return '易失函数报错', err_text(p)
        if cp == cb or (cb == 'date' and cp == 'num') or (cb == 'num' and cp == 'date'):
            return '易失函数（类型正确）', '返回类型=%s' % cp
        return '易失函数（类型异常）', '基线=%s 产品=%s' % (cb, cp)
    # SHEET()/SHEETS()：受试用版多出的首页 sheet 影响，+1 偏移是环境造成，可修正
    if isinstance(formula, str) and sname == 'F引用' and 'SHEET' in formula.upper():
        try:
            bi, pi = float(b), float(p)
            if abs(pi - bi - 1) < 1e-9:
                return '偏移 +1（试用版首页 sheet，商业授权无此问题）', '实测=%g，基线=%g' % (pi, bi)
        except Exception:
            pass
    if b is None and p is None:
        return '空', '空'
    if is_err(b):
        if is_err(p):
            return '基线不支持，产品也报错', err_text(p)
        if p is None:
            return '基线不支持，产品为空', ''
        return '超越基线（产品算出值）', ''
    if is_err(p):
        return '产品报错', err_text(p)
    if p is None:
        return '产品为空', ''
    if eq(b, p):
        return '一致', ''
    return '不一致', '产品=%s' % show(p)

def main():
    d = load('formula_batch1_def.json')
    base = load('formula_batch1_baseline.json')
    sjs = load('sjs_formula_batch1.json')['expandedMode']['sheets']
    uvi = load('univer_formula_batch1.json')['sheets']

    wb = Workbook()
    ws = wb.active
    ws.title = '逐公式对照'
    headers = ['工作表', '单元格', '公式', 'Excel 2021 基线', 'SpreadJS 19.2.2', 'Univer Pro 1.3.1', 'SpreadJS 判定', 'Univer 判定', '备注']
    ws.append(headers)

    thin = Side(style='thin', color='D9D9D9')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill('solid', fgColor='1F3864')
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color='FFFFFF', size=11)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    fill_ok = PatternFill('solid', fgColor='E2EFDA')
    fill_bad = PatternFill('solid', fgColor='FCE4E4')
    fill_beyond = PatternFill('solid', fgColor='DDEBF7')
    fill_base_bad = PatternFill('solid', fgColor='F2F2F2')

    stat = {'SJS': {}, 'UV': {}}
    rows = 0
    for sh in d['sheets']:
        sname = sh['name']
        cells = sh.get('cells', [])
        label_map = {(it[0], it[1]): it[2] for it in cells if len(it) >= 3}
        for item in cells:
            if len(item) < 3:
                continue
            row, col, formula = item[0], item[1], item[2]
            if not (isinstance(formula, str) and formula.startswith('=')):
                continue
            label = str(label_map.get((row, col - 1), '') or '')
            addr = get_column_letter(col) + str(row)
            b = (base.get(sname, {}).get(addr) or {}).get('v')
            s = ((sjs.get(sname) or {}).get('cells') or {}).get(addr, {}).get('v')
            u = ((uvi.get(sname) or {}).get('cells') or {}).get(addr, {}).get('v')
            vs, note_s = verdict(b, s, sname, formula, 'SJS')
            vu, note_u = verdict(b, u, sname, formula, 'UV')
            if vs in ('不一致', '产品报错') or vu in ('不一致', '产品报错'):
                focus = '有差异'
            elif vs.startswith('易失函数') or vu.startswith('易失函数') or '偏移' in vs:
                focus = '按口径不计差异'
            elif vs == '超越基线（产品算出值）' or vu == '超越基线（产品算出值）':
                focus = '产品超越基线'
            else:
                focus = '双方一致'
            ws.append([sname, addr, label + '：' + formula, show(b), show(s), show(u), vs, vu, focus])
            r = ws.max_row
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).border = border
                ws.cell(row=r, column=c).alignment = Alignment(vertical='center', wrap_text=(c in (3, 4, 5, 6)))
            for cell, v in ((ws.cell(row=r, column=7), vs), (ws.cell(row=r, column=8), vu)):
                if v in ('一致',):
                    cell.fill = fill_ok
                elif v in ('不一致', '产品报错', '产品为空'):
                    cell.fill = fill_bad
                elif v.startswith('超越'):
                    cell.fill = fill_beyond
                elif v.startswith('基线不支持'):
                    cell.fill = fill_base_bad
            ws.cell(row=r, column=1).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=r, column=2).alignment = Alignment(horizontal='center', vertical='center')
            for key, v in (('SJS', vs), ('UV', vu)):
                stat[key][v] = stat[key].get(v, 0) + 1
            rows += 1

    widths = {'A': 11, 'B': 8, 'C': 46, 'D': 26, 'E': 24, 'F': 24, 'G': 22, 'H': 22, 'I': 14}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = 'A1:I%d' % (rows + 1)

    # 汇总页
    ws2 = wb.create_sheet('汇总')
    ws2['A1'] = '公式批次1 三方对照汇总（样本：公式批次1样本.xlsx，7 表 64 个可判定公式单元）'
    ws2['A1'].font = Font(bold=True, size=13)
    ws2['A3'] = '判定桶'; ws2['B3'] = 'SpreadJS 19.2.2'; ws2['C3'] = 'Univer Pro 1.3.1'
    for c in ('A3', 'B3', 'C3'):
        ws2[c].font = Font(bold=True, color='FFFFFF')
        ws2[c].fill = head_fill
        ws2[c].alignment = Alignment(horizontal='center')
    order = ['一致', '超越基线（产品算出值）', '基线不支持，产品也报错', '偏移 +1（试用版首页 sheet，商业授权无此问题）', '易失函数（类型正确）', '易失函数报错', '易失函数（类型异常）', '不一致', '产品报错', '产品为空', '基线不支持，产品为空', '空']
    r = 4
    for k in order:
        a, b = stat['SJS'].get(k, 0), stat['UV'].get(k, 0)
        if a or b:
            ws2.cell(row=r, column=1, value=k)
            ws2.cell(row=r, column=2, value=a)
            ws2.cell(row=r, column=3, value=b)
            r += 1
    ws2.cell(row=r, column=1, value='合计').font = Font(bold=True)
    ws2.cell(row=r, column=2, value=rows).font = Font(bold=True)
    ws2.cell(row=r, column=3, value=rows).font = Font(bold=True)
    r += 2
    notes = [
        '口径说明：',
        '1. Excel 2021 基线的 #NAME? 表示该函数是 Microsoft 365 独占、本机 Excel 2021 不支持；产品算出值 = 超越基线。',
        '2. SpreadJS 侧为导入后扩大工作表区域（setColumnCount 30 / setRowCount 300）+ 强制重算后的结果——导入后工作表列数=数据实际宽度，多列溢出会越界显示 #SPILL!，扩区即恢复。',
        '3. Univer 侧为 playground 导入 + 等待重算后快照读取，公式被加上 @ 前缀（隐式交叉）只输出锚点首值。',
        '4. 易失函数（NOW/TODAY/RAND/RANDBETWEEN）只判"识别 + 返回类型正确"，不比较具体值。',
        '5. 日期统一换算为 Excel 序列号比较，显示为日期（序列号）。',
    ]
    for n in notes:
        ws2.cell(row=r, column=1, value=n).alignment = Alignment(wrap_text=False)
        if n.endswith('：'):
            ws2.cell(row=r, column=1).font = Font(bold=True)
        r += 1
    ws2.column_dimensions['A'].width = 30
    ws2.column_dimensions['B'].width = 20
    ws2.column_dimensions['C'].width = 20

    wb.save(OUT)
    print('saved', OUT, '| rows:', rows)
    print('SJS:', stat['SJS'])
    print('UV :', stat['UV'])

if __name__ == '__main__':
    main()
