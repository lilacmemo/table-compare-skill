# -*- coding: utf-8 -*-
"""把三方结果归一化并给出支持/一致判定（Python 侧，供差异明细与统计使用）。

判定口径（与工作簿里的活公式一一对应）：
  1. 产品返回 #NAME? -> 不支持（函数未实现），不参与值比较
  2. 产品公式写入被拒 -> 无法写入（签名不兼容）
  3. 产品返回其他错误 -> 产品报错
  4. 双方都有值：数值按 1e-9 相对容差；布尔 vs 0/1 记“类型不同（布尔/数值）”；
     数值 vs 文本 记“类型不同（数值/文本）”；其余按文本精确比较
  5. 易失函数（NOW/TODAY/RAND/RANDBETWEEN/RANDARRAY）：支持且返回非错误值 →
     一律记「支持（易失函数，结果不对比）」，不计入一致也不计入不一致（2026-11-05 用户定版）
"""
import json, os, re, datetime, time

HERE = os.path.dirname(os.path.abspath(__file__))
CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def clean(v):
    return CTRL.sub('', v) if isinstance(v, str) else v


def sjs_serial(ms):
    offset_ms = -time.timezone * 1000
    dt = datetime.datetime(1970, 1, 1) + datetime.timedelta(milliseconds=ms + offset_ms)
    delta = dt - datetime.datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400.0 + delta.microseconds / 86400e6


def norm_sjs(r):
    t, v = r.get('valueType'), r.get('value')
    if t == 'error':
        e = str((v or {}).get('err') if isinstance(v, dict) else v)
        if not e.startswith('#'):
            e = '#' + e
        return ('ERR:' + e, 'err')
    if t == 'date':
        # 探测页已把 Date 按本地日历字段换算成 Excel 序列号，这里直接用
        return (round(float(v), 9), 'num')
    if t == 'boolean':
        return (bool(v), 'bool')
    if t == 'number':
        return (float(v), 'num')
    if t == 'null':
        return ('', 'empty')
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (float(v), 'num')
    if isinstance(v, bool):
        return (v, 'bool')
    return (clean(str(v)), 'text')


ERR_RE = re.compile(r'^#(NAME\?|N/A|VALUE!|NUM!|REF!|DIV/0!|NULL!|SPILL!|CALC!|GETTING_DATA)$')


def norm_univer(r):
    t, v = r.get('valueType'), r.get('value')
    # 快照 t=3 是 BOOLEAN 类型标记（v 字段序列化为 1/0，但显示层 getDisplayValue()
    # 返回 TRUE/FALSE、下游语义与 Excel 一致——2026-11-05 用户实测复核后修正口径）
    if r.get('t') == 3:
        return (bool(v), 'bool')
    # Univer 快照把错误值存成 '#NAME?' 这样的字符串（valueType 仍是 string），必须先识别出来
    if isinstance(v, str) and ERR_RE.match(v.strip()):
        return ('ERR:' + v.strip(), 'err')
    if t == 'error':
        raw = str(v.get('raw') if isinstance(v, dict) else v)
        # Univer 的快照把错误存成字符串 '#NAME?'；个别版本存成对象，做一次兜底
        m = re.search(r'#(NAME\?|N/A|VALUE!|NUM!|REF!|DIV/0!|NULL!|SPILL!|CALC!)', raw)
        return ('ERR:' + ('#' + m.group(1) if m else '#UNKNOWN'), 'err')
    if t == 'null' or v is None:
        return ('', 'empty')
    if isinstance(v, bool):
        return (bool(v), 'bool')
    if isinstance(v, (int, float)):
        return (float(v), 'num')
    return (clean(str(v)), 'text')


def norm_excel(r):
    t, v = r.get('valueType'), r.get('value')
    if t in ('error', 'reject'):
        return ('ERR:' + str(r.get('err') or v), 'err')
    if t == 'empty':
        return ('', 'empty')
    if t == 'boolean':
        return (bool(v), 'bool')
    if t == 'number':
        return (float(v), 'num')
    return (clean(str(v)), 'text')


CMPLX = re.compile(r'^([-+0-9.eE]+)([-+][0-9.eE]+)i$')


def as_number(x):
    """把 '1.6094379124341003+0.9272952180016122i' / '-7' / '0.001' 解析成可比较的数或复数"""
    if not isinstance(x, str):
        return None
    t = x.strip()
    m = CMPLX.match(t)
    if m:
        try:
            return complex(float(m.group(1)), float(m.group(2)))
        except ValueError:
            return None
    try:
        return float(t)
    except ValueError:
        return None


def num_equal(a, b):
    try:
        if isinstance(a, complex) or isinstance(b, complex):
            return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
        return abs(a - b) <= 1e-9 * max(1e-9, abs(a))
    except TypeError:
        return False


def fmt_num_text(x):
    """把 '1.6094379124341003+0.9272952180016122i' 这类文本按 15 位有效数字重排，
    与 Excel 的显示精度对齐；非数值文本原样返回。"""
    if not isinstance(x, str):
        return x
    t = x.strip()
    m = CMPLX.match(t)
    if m:
        try:
            re_, im_ = float(m.group(1)), float(m.group(2))
        except ValueError:
            return x
        sign = '+' if im_ >= 0 else ''
        return ('%.15g' % re_) + sign + ('%.15g' % im_) + 'i'
    try:
        f = float(t)
    except ValueError:
        return x
    if abs(f) > 1e15:
        return x
    return '%.15g' % f


def text_form(v, kind):
    """把结果转成“单元格显示文本”，与工作簿里 F 列 / K&"" 的口径一致"""
    if kind == 'bool':
        return 'TRUE' if v is True else 'FALSE'
    if kind == 'num':
        try:
            f = float(v)
            if f == int(f) and abs(f) < 1e15:
                return str(int(f))
            return ('%.15g' % f)
        except (TypeError, ValueError):
            return str(v)
    if kind == 'empty':
        return ''
    return str(v)


def judge(excel, prod):
    """excel / prod = (value, kind)"""
    pv, pk = prod
    ev, ek = excel
    if pk == 'err' and ek == 'err' and str(pv) == str(ev):
        return '一致（双方同为错误值）'
    if pk == 'err':
        if str(pv) == 'ERR:#NAME?':
            return '不支持（#NAME?）'
        if str(pv) == 'ERR:SET_REJECT':
            return '公式写入被拒'
        return '产品报错'
    if ek == 'err':
        return 'Excel 报错'
    if ek == 'num' and pk == 'num':
        if abs(ev - pv) <= max(1e-9, abs(ev) * 1e-9):
            return '一致'
        return '数值不同'
    if ek == 'bool' and pk == 'num':
        return '类型不同（布尔/数值）'
    if ek == 'num' and pk != 'num':
        return '类型不同（数值/文本）'
    a, b = text_form(ev, ek), text_form(pv, pk)
    if a == b:
        return '一致'
    return '文本不同'


DIFF_LABELS = {'数值不同', '文本不同', '类型不同', '类型不同（布尔/数值）',
               '类型不同（数值/文本）', '产品报错'}

# 易失函数（2026-11-05 用户定版）：每次重算结果都不同，值对比无意义。
# 判定口径：只要支持且返回了非错误值 → 一律记「支持（易失函数，结果不对比）」，不计入一致也不计入不一致。
VOLATILE_FNS = {'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY'}
VOLATILE_LABEL = '支持（易失函数，结果不对比）'


ENV_ERR_RE = re.compile(r'^#(NAME\?|N/A|VALUE!|NUM!|REF!|DIV/0!|NULL!|SPILL!|CALC!)$')


def norm_env_display(display, raw_value=None):
    """环境依赖探针结果归一：display 文本为主，识别错误值。返回 (value, kind)。"""
    d = '' if display is None else str(display).strip()
    if ENV_ERR_RE.match(d):
        return ('ERR:' + d, 'err')
    if d == '':
        return ('', 'empty')
    try:
        return (float(d), 'num')
    except ValueError:
        return (d, 'text')


def load():
    cases = json.load(open(os.path.join(HERE, 'cases_v2.json'), encoding='utf-8'))
    excluded = json.load(open(os.path.join(HERE, 'excluded_v2.json'), encoding='utf-8'))
    ex = {r['name']: r for r in json.load(open(os.path.join(HERE, 'excel_eval_v2.json'), encoding='utf-8'))}
    sj = {r['name']: r for r in json.load(open(os.path.join(HERE, 'probe_sjs_v2.json'), encoding='utf-8'))['rows']}
    un_path = os.path.join(HERE, 'probe_univer_v3.json')
    if not os.path.exists(un_path):
        un_path = os.path.join(HERE, 'probe_univer_v2.json')
    un = {r['name']: r for r in json.load(open(un_path, encoding='utf-8'))['rows']}
    # 环境依赖 19 函数：双端代码实测支持性（值不对比，仅判 #NAME?）
    env_path = os.path.join(HERE, 'probe_env.json')
    env_sjs = {}
    env_uni = {}
    if os.path.exists(env_path):
        envp = json.load(open(env_path, encoding='utf-8'))
        for r in envp['sjs']['rows']:
            t, v = r.get('valueType'), r.get('value')
            if t == 'error':
                e = str((v or {}).get('err') if isinstance(v, dict) else v)
                env_sjs[r['name']] = ('ERR:' + (e if e.startswith('#') else '#' + e), 'err')
            elif t == 'boolean':
                env_sjs[r['name']] = (bool(v), 'bool')
            elif t == 'number':
                env_sjs[r['name']] = (float(v), 'num')
            elif t == 'object':
                env_sjs[r['name']] = ('（返回图片对象）', 'text')
            else:
                env_sjs[r['name']] = norm_env_display(str(v) if t != 'null' else '')
        for r in envp['univer']['rows']:
            env_uni[r['name']] = norm_env_display(r.get('display'))
    for e in excluded:
        e['sjs_env'] = env_sjs.get(e['name'], ('未测', 'text'))
        e['uni_env'] = env_uni.get(e['name'], ('未测', 'text'))
    out = []
    for c in cases:
        n = c['name']
        e = norm_excel(ex[n]); s = norm_sjs(sj[n]); u = norm_univer(un[n])
        r = {
            'name': n, 'family': c['family'], 'formula': c['formula'], 'arrayNote': c.get('arrayNote'),
            'excel': e, 'sjs': s, 'univer': u,
            'sjs_verdict': judge(e, s), 'univer_verdict': judge(e, u),
            'sjs_support': ('不支持' if s[0] == 'ERR:#NAME?' else ('无法写入' if s[0] == 'ERR:SET_REJECT' else '支持')),
            'univer_support': ('不支持' if u[0] == 'ERR:#NAME?' else ('无法写入' if u[0] == 'ERR:SET_REJECT' else '支持')),
        }
        # 易失函数豁免：支持且返回非错误值 → 只判支持，不比结果
        if n in VOLATILE_FNS:
            if r['sjs_support'] == '支持' and s[1] != 'err':
                r['sjs_verdict'] = VOLATILE_LABEL
            if r['univer_support'] == '支持' and u[1] != 'err':
                r['univer_verdict'] = VOLATILE_LABEL
        out.append(r)
    return out, excluded


def summarize(rows):
    def stat(key):
        from collections import Counter
        c = Counter(r[key] for r in rows)
        return c
    return {'sjs': stat('sjs_verdict'), 'univer': stat('univer_verdict'),
            'sjs_support': stat('sjs_support'), 'univer_support': stat('univer_support')}


if __name__ == '__main__':
    rows, excluded = load()
    json.dump(rows, open(os.path.join(HERE, 'verdicts_v2.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    s = summarize(rows)
    print('用例', len(rows), '| 环境依赖排除', len(excluded))
    for k in ('sjs_support', 'univer_support', 'sjs', 'univer'):
        print(k, dict(s[k]))
    for who, key in (('SJS', 'sjs_verdict'), ('Univer', 'univer_verdict')):
        d = [r for r in rows if r[key] in DIFF_LABELS]
        print('\n%s 差异 %d 个:' % (who, len(d)))
        for r in d[:80]:
            print('  %-28s %-18s excel=%r prod=%r' % (r['name'], r[key], r[key == 'x' or 'excel'][0], r['sjs' if who == 'SJS' else 'univer'][0]))
