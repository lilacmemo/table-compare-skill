# -*- coding: utf-8 -*-
"""从 api-items.json 生成：PASS/FAIL 二值判定 + 分组统计 + 友好性五维评分。

判定口径（只有两条，全部按"公开 API"）：
  PASS = 该侧存在公开 API，且调用成功（setOk !== false）
  FAIL = 运行时没有公开入口（noApi），或调用即抛错（setOk === false）
  N/A  = Excel 本身没有这项能力（只对 Univer 侧的 tierHint='S-' 生效），不计入分母

"能不能读回"单独一列展示，不再参与 PASS/FAIL 判定——读不回不等于没有 API。
"""
import json, os
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'out', 'api-items.json')
d = json.load(open(SRC, encoding='utf-8'))
items = list(d['items'])

GROUP_NAME = {
    'A': '字体', 'B': '填充', 'C': '边框', 'D': '对齐与合并', 'E': '数字格式',
    'F': '行高列宽', 'G': '条件格式规则', 'H': '条件格式图形参数',
    'I': '条件格式管理', 'J': '样式复用与主题', 'K': '保护与视图',
}


def verdict(side):
    if not side or side.get('skipped'):
        return 'N/A', '不测'
    if side.get('tierHint') == 'S-':
        return 'N/A', 'Excel 本无此能力'
    if side.get('noApi'):
        return 'FAIL', side.get('note') or '运行时无公开入口'
    if side.get('setOk') is False:
        return 'FAIL', '调用即抛错：' + str(side.get('err'))[:50]
    return 'PASS', ''


rows = []
for i in items:
    vs, ws = verdict(i.get('sjs'))
    vu, wu = verdict(i.get('uni'))
    rows.append({
        'id': i['id'], 'g': i['g'], 'name': i['name'], 'excel': i['excel'],
        'sjs': vs, 'uni': vu, 'sjs_why': ws, 'uni_why': wu,
        'sjs_api': (i.get('sjs') or {}).get('api', ''),
        'uni_api': (i.get('uni') or {}).get('api', ''),
        'raw_sjs': i.get('sjs'), 'raw_uni': i.get('uni'),
    })

# 非基线项（Excel 本没有）：以 Univer 侧 tierHint='S-' 为准
NON_BASE = set(r['id'] for r in rows if r['uni'] == 'N/A' and r['uni_why'] == 'Excel 本无此能力')
base = [r for r in rows if r['id'] not in NON_BASE]


def tally(side, rs):
    g = defaultdict(lambda: defaultdict(int))
    for r in rs:
        g[r['g']][r[side]] += 1
    return g


def pct(n, m):
    return 0.0 if not m else n / float(m) * 100


# ---------- 分组矩阵 ----------
lines = ['| 组 | 能力项 | SJS PASS | SJS FAIL | Univer PASS | Univer FAIL | 差值（PASS 率） |',
         '|:--|--:|--:|--:|--:|--:|--:|']
for g in 'ABCDEFGHIJK':
    sub = [r for r in base if r['g'] == g]
    sp = sum(1 for r in sub if r['sjs'] == 'PASS')
    sf = sum(1 for r in sub if r['sjs'] == 'FAIL')
    up = sum(1 for r in sub if r['uni'] == 'PASS')
    uf = sum(1 for r in sub if r['uni'] == 'FAIL')
    diff = pct(sp, sp + sf) - pct(up, up + uf)
    lines.append('| %s %s | %d | %d | %d | %d | %d | %+.0f pp |' % (g, GROUP_NAME[g], len(sub), sp, sf, up, uf, diff))
sp = sum(1 for r in base if r['sjs'] == 'PASS')
sf = sum(1 for r in base if r['sjs'] == 'FAIL')
up = sum(1 for r in base if r['uni'] == 'PASS')
uf = sum(1 for r in base if r['uni'] == 'FAIL')
lines.append('| **合计（基线 %d 项）** | %d | **%d** | **%d** | **%d** | **%d** | **%+.0f pp** |'
             % (len(base), len(base), sp, sf, up, uf, pct(sp, sp + sf) - pct(up, up + uf)))
lines.append('')
lines.append('> 分母是"Excel 基线项"（%d 项），已剔除 %d 项 Excel 本身没有的能力（本报告只在附录里单独标注）。'
             % (len(base), len(NON_BASE)))
group_md = '\n'.join(lines)

# ---------- 友好性五维（只用 PASS 项算） ----------
SILENT = {'sjs': 4, 'uni': 5}


def five(side):
    api = [r['raw_' + side] for r in rows if r[side] == 'PASS']
    n = len(api) or 1
    step_map = {1: 5.0, 2: 3.8, 3: 2.4}
    pm = {'typed': 5.0, 'enum': 4.0, 'loose': 2.0}
    return {
        'n': len(api),
        '覆盖广度': len(api) / float(len(base)) * 5,
        '一步到位': sum(step_map.get(s.get('steps'), 1.2) for s in api) / n,
        '可读回': sum(1 for s in api if s.get('readable')) / n * 5,
        '参数类型化': sum(pm.get(s.get('param'), 3.0) for s in api) / n,
        '错误反馈': max(0.0, 5.0 - 0.5 * SILENT[side]),
        'steps': sum(s.get('steps') or 0 for s in api) / n,
        'unreadable': sum(1 for s in api if not s.get('readable')),
        'undoc': sum(1 for s in api if not s.get('doc')),
    }


fs, fu = five('sjs'), five('uni')
DIMS = ('覆盖广度', '一步到位', '可读回', '参数类型化', '错误反馈')
score_md = ['| 维度 | SpreadJS | Univer | 口径 |', '|:--|--:|--:|:--|']
score_md.append('| 覆盖广度 | %.1f | %.1f | PASS 项 / 基线 %d 项 → ×5 |' % (fs['覆盖广度'], fu['覆盖广度'], len(base)))
score_md.append('| 一步到位 | %.1f | %.1f | 平均调用步数 %.2f vs %.2f（1 步 5 / 2 步 3.8 / 3 步 2.4） |'
                % (fs['一步到位'], fu['一步到位'], fs['steps'], fu['steps']))
score_md.append('| 可读回 | %.1f | %.1f | PASS 项里能读回等值的比例（SJS %d 项读不回 / Univer %d 项） |'
                % (fs['可读回'], fu['可读回'], fs['unreadable'], fu['unreadable']))
score_md.append('| 参数类型化 | %.1f | %.1f | typed 5 / enum 4 / 裸值或魔法字符串 2 |' % (fs['参数类型化'], fu['参数类型化']))
score_md.append('| 错误反馈 | %.1f | %.1f | 5 − 0.5×静默失败条数（SJS %d / Univer %d） |'
                % (fs['错误反馈'], fu['错误反馈'], SILENT['sjs'], SILENT['uni']))
score_md.append('| **合计（/25）** | **%.1f** | **%.1f** | — |'
                % (sum(fs[k] for k in DIMS), sum(fu[k] for k in DIMS)))
score_md = '\n'.join(score_md)

# ---------- 逐项矩阵 ----------
mat = ['| ID | 组 | 能力项 | Excel 基线 | SpreadJS API | SJS | Univer API | Uni | 备注 |',
       '|:--|:--|:--|:--|:--|:-:|:--|:-:|:--|']


def short(s, n=44):
    s = str(s or '').replace('|', '\\|')
    return s if len(s) <= n else s[:n] + '…'


for r in rows:
    if r['sjs'] == 'N/A' and r['uni'] == 'N/A':
        continue
    note = '；'.join(x for x in [r['sjs_why'], r['uni_why']] if x)
    mat.append('| %s | %s | %s | %s | %s | %s | %s | %s | %s |'
               % (r['id'], r['g'], r['name'], short(r['excel'], 30), short(r['sjs_api'], 38), r['sjs'],
                  short(r['uni_api'], 38), r['uni'], short(note, 54)))
matrix_md = '\n'.join(mat)

out_dir = os.path.join(HERE, 'out')
os.makedirs(out_dir, exist_ok=True)
open(os.path.join(out_dir, 'api-group-scores.md'), 'w', encoding='utf-8').write(
    '### 分组 PASS/FAIL\n\n' + group_md + '\n\n### 友好性五维\n\n' + score_md + '\n')
open(os.path.join(out_dir, 'api-matrix.md'), 'w', encoding='utf-8').write(matrix_md + '\n')

print(group_md)
print()
print('基线项', len(base), '剔除', len(NON_BASE), NON_BASE)
print('SJS PASS %d / FAIL %d' % (sp, sf))
print('Uni PASS %d / FAIL %d' % (up, uf))
print()
print(score_md)
print()
print('已写出 out/api-group-scores.md、out/api-matrix.md')
