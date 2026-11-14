# -*- coding: utf-8 -*-
# 公式批次1比对：Excel 2021 基线 vs SJS(默认态/动态态) vs Univer
import json, io, os, sys, re

ROOT = os.path.dirname(os.path.abspath(__file__))
baseline = json.load(io.open(os.path.join(ROOT, 'formula_batch1_baseline.json'), encoding='utf-8'))

def num_eq(a, b):
    try:
        fa, fb = float(a), float(b)
        return abs(fa - fb) < max(1e-9, abs(fb) * 1e-9)
    except (TypeError, ValueError):
        return False

import datetime
def to_serial(v):
    # ISO 日期字符串 -> Excel 序列号（按本地时区取日，修 UTC 偏移导致的跨日错误）
    if isinstance(v, str) and re.match(r'^\d{4}-\d{2}-\d{2}T', v):
        try:
            dt = datetime.datetime.fromisoformat(v.replace('Z', '+00:00'))
            if dt.tzinfo is not None:
                dt = dt.astimezone()  # 转本机时区（GMT+8）
            dt = dt.replace(tzinfo=None)
            base = datetime.datetime(1899, 12, 30)
            return (dt - base).total_seconds() / 86400
        except Exception:
            return v
    return v

def val_eq(v1, v2):
    v1, v2 = to_serial(v1), to_serial(v2)
    if v1 is None and v2 is None: return True
    # 双方可转数值则按数值比较（覆盖 '46330' vs 46330.0）
    try:
        f1, f2 = float(v1), float(v2)
        return num_eq(f1, f2)
    except (TypeError, ValueError):
        pass
    s1, s2 = str(v1).strip(), str(v2).strip()
    if s1.lower() in ('true', 'false') and s2.lower() in ('true', 'false'):
        return s1.lower() == s2.lower()
    return s1 == s2

def norm_pv(pv):
    if isinstance(pv, dict):
        return pv.get('_error') or pv.get('_calcError') or str(pv)
    if isinstance(pv, str) and ('_error' in pv or '_calcError' in pv):
        try:
            o = json.loads(pv)
            return o.get('_error') or o.get('_calcError') or pv
        except Exception:
            return pv
    return pv

ERRS = {'#NAME?', '#REF!', '#DIV/0!', '#VALUE!', '#N/A', '#NUM!', '#NULL!', '#SPILL!', '#溢出!'}

VOLATILE_TYPE_ONLY = {'NOW': 'number', 'TODAY': 'number', 'RAND': 'number01', 'RANDBETWEEN': 'number01'}
SHEET_FNS = {'B1': 'SHEET()', 'B2': 'SHEETS()', 'B3': 'SHEET("Data")', 'B4': 'SHEET(Data!A1)'}

def classify(bv, pv, formula):
    """bv/pv: baseline/product value (str or None); returns (bucket, note)"""
    b_err = bv in ERRS if bv else False
    pv = norm_pv(pv)
    p_err = pv in ERRS if pv else False
    if b_err and p_err:
        return ('both_err', bv + '/' + pv) if bv != pv else ('both_err_same', bv)
    if b_err and not p_err:
        return ('prod_beyond_excel', 'Excel2021=%s 产品=%s' % (bv, pv))
    if not b_err and p_err:
        return ('prod_err', 'Excel=%s 产品=%s' % (bv, pv))
    if val_eq(bv, pv):
        return ('match', '')
    # 数值字符串兼容（"61" vs 61）
    if num_eq(bv, pv):
        return ('match', '')
    return ('mismatch', 'Excel=%s 产品=%s' % (bv, pv))

def run_compare(prod, prod_label):
    rows = []
    for sname, cells in baseline.items():
        psheet = prod['sheets'].get(sname)
        pcells = psheet['cells'] if psheet else {}
        for addr, b in sorted(cells.items(), key=lambda x: (int(x[0][1:] if x[0][1:].isdigit() else 9999), x[0])):
            bf, bv = b['f'], b['v']
            if bf is not None and not str(bf).startswith('='): bf = None
            p = pcells.get(addr, {})
            pv, pf = p.get('v'), p.get('f')
            if not bf:  # 溢出/普通数据单元格
                if bv is None or bv == '': continue
                if not pcells:
                    rows.append((sname, addr, 'no_sheet', '产品缺表'))
                    continue
                bkt, note = classify(bv, str(pv) if pv is not None else None, None)
                # 数据单元格：只有不匹配才记
                if bkt != 'match':
                    rows.append((sname, addr, bkt, note))
                continue
            # 公式单元格
            fname = (bf.split('(')[0].lstrip('=').upper())
            if sname == 'F易失' and fname in VOLATILE_TYPE_ONLY and bf and str(bf).startswith('='):
                # 易失函数：只判类型
                kind = VOLATILE_TYPE_ONLY[fname]
                ok = pv is not None and not (str(pv) in ERRS)
                if ok and kind == 'number': ok = isinstance(pv, (int, float)) or bool(re.match(r'^\d{4}-\d{2}-\d{2}T', str(pv))) or str(pv).replace('.', '').replace('-', '').isdigit()
                if ok and kind == 'number01':
                    try: ok = 0.0 <= float(pv) <= 100.0
                    except Exception: ok = False
                rows.append((sname, addr, 'volatile_ok' if ok else 'volatile_bad', '%s=%s' % (fname, pv)))
                continue
            if sname == 'F引用' and addr in SHEET_FNS:
                rows.append((sname, addr, 'sheet_fn', '%s: Excel=%s 产品=%s' % (SHEET_FNS[addr], bv, pv)))
                continue
            pv_s = None if pv is None else str(pv)
            bkt, note = classify(bv, pv, bf)  # 传原始 pv（dict/str 均可），classify 内 norm_pv 归一
            extra = '' if (pf or not bf) else ' 公式丢失'
            rows.append((sname, addr, bkt + extra, note))
    # 汇总
    from collections import Counter
    cnt = Counter(r[2].replace(' 公式丢失', '') for r in rows)
    print('=== %s ===' % prod_label)
    for k, v in cnt.most_common(): print('  %-22s %d' % (k, v))
    for r in rows:
        if r[2] not in ('match',) and not r[2].startswith('volatile_ok') and not r[2].startswith('sheet_fn') \
           and r[2] != 'both_err_same':
            print('  [%s %s] %s %s' % r)
    return rows

sjs = json.load(io.open(os.path.join(ROOT, 'sjs_formula_batch1.json'), encoding='utf-8'))

def norm(product):
    sheets = {}
    for name, s in product['sheets'].items():
        sheets[name] = {'cells': s['cells'] if isinstance(s, dict) and 'cells' in s else s}
    return {'sheets': sheets}

sjsN = norm(sjs['dynamicMode'])
print('SJS sheets:', list(sjsN['sheets']))
run_compare(sjsN, 'SJS v19.2.2 (dynamicMode)')
print()
print('SJS default-mode allowDynamicArray:', sjs['defaultMode']['options']['allowDynamicArray'])

# ---- Univer Pro 1.3.1 ----
uv = json.load(io.open(os.path.join(ROOT, 'univer_formula_batch1.json'), encoding='utf-8'))
uvN = {'sheets': {k: {'cells': v['cells']} for k, v in uv['sheets'].items()}}
run_compare(uvN, 'Univer Pro 1.3.1 (import+15s)')
# 动态数组 @ 前缀统计
at_cnt = sum(1 for s in uvN['sheets'].values() for c in s['cells'].values() if isinstance(c.get('f'), str) and c['f'].lstrip().startswith('=@'))
f_cnt = sum(1 for s in uvN['sheets'].values() for c in s['cells'].values() if isinstance(c.get('f'), str))
print('Univer @ 前缀公式: %d / %d' % (at_cnt, f_cnt))
