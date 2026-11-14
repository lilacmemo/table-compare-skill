# extract_sjs.py —— 从 batch1_sjs_dump.json 提取每图关键配置值
import json, re

d = json.load(open('batch1_sjs_dump.json', encoding='utf-8'))

def ln_info(sp):
    if not sp: return None
    ln = sp.get('ln', {})
    if ln.get('noFill'): return 'none'
    fill = ln.get('solidFill', {})
    if 'srgbClr' in fill: color = fill['srgbClr']['val']
    elif 'schemeClr' in fill: color = 'scheme:' + str(fill['schemeClr'].get('val'))
    else: color = None
    return {'color': color, 'w': ln.get('w'), 'dash': ln.get('prstDash')}

def rpr_info(p):
    pr = (p or {}).get('pPr', {}).get('defRPr', {}) or {}
    el = ((p or {}).get('elements') or [{}])
    el = el[0] if el else {}
    r = el.get('rPr', {}) or pr
    fill = r.get('solidFill', {})
    color = fill.get('srgbClr', {}).get('val') if isinstance(fill, dict) else None
    return {'text': el.get('t'), 'font': (r.get('latin') or {}).get('typeface'),
            'sz': r.get('sz'), 'b': r.get('b'), 'i': r.get('i'), 'color': color}

def text_of(title):
    if not title: return None
    try: return title['tx']['rich']['p'][0]['elements'][0]['t']
    except Exception: return '?'

out = []
for s in d['sheets']:
    if s['sheet'] == 'Evaluation Version': continue
    for c in s['charts']:
        j = c.get('toJSON') or {}
        ch = (((j.get('shapeData') or {}).get('graphicFrame') or {}).get('graphic') or {}).get('graphicData', {}).get('chart', {}).get('chartSpace', {}).get('chart', {})
        rec = {'sheet': s['sheet'], 'name': c['name'], 'type': c['type']}

        # 标题
        t = ch.get('title')
        rec['title_deleted'] = ch.get('autoTitleDeleted')
        rec['title'] = rpr_info(t['tx']['rich']['p'][0]) if (t and t.get('tx')) else (('无tx:' + json.dumps(t, ensure_ascii=False)[:120]) if t else None)
        rec['title_overlay'] = (t or {}).get('overlay')

        # 轴
        axes = ch.get('plotArea', {}).get('axes', [])
        axrec = []
        for ax in axes:
            a = {'axPos': ax.get('axPos'), 'deleted': ax.get('delete'),
                 'majorTickMark': ax.get('majorTickMark'), 'minorTickMark': ax.get('minorTickMark'),
                 'tickLblPos': ax.get('tickLblPos'),
                 'numFmt': (ax.get('numFmt') or {}).get('formatCode'),
                 'axis_line': ln_info(ax.get('spPr')),
                 'axis_title': None, 'label_rot': None, 'label_font': None}
            if ax.get('title'):
                try: a['axis_title'] = rpr_info(ax['title']['tx']['rich']['p'][0]) or None
                except Exception: a['axis_title'] = '?'
                try:
                    a['axis_title_rot'] = ax['title']['tx']['rich']['bodyPr'].get('rot')
                except Exception: pass
            txpr = ax.get('txPr')
            if txpr:
                a['label_rot'] = txpr.get('bodyPr', {}).get('rot')
                try: a['label_font'] = rpr_info(txpr['p'][0])
                except Exception: pass
            sc = ax.get('scaling', {})
            a['orientation'] = sc.get('orientation')   # 0=minMax? 1=?
            a['logBase'] = sc.get('logBase')
            a['min'] = ax.get('min'); a['max'] = ax.get('max'); a['majorUnit'] = ax.get('majorUnit'); a['minorUnit'] = ax.get('minorUnit')
            if ax.get('majorGridlines'): a['majorGrid'] = ln_info(ax['majorGridlines'].get('spPr'))
            if ax.get('minorGridlines'): a['minorGrid'] = ln_info(ax['minorGridlines'].get('spPr'))
            axrec.append(a)
        rec['axes'] = axrec

        # 图例
        lg = ch.get('legend')
        if lg:
            l = {'legendPos': lg.get('legendPos'), 'overlay': lg.get('overlay'),
                 'fill': ln_info(lg.get('spPr')), 'font': None}
            txpr = lg.get('txPr')
            if txpr:
                try: l['font'] = rpr_info(txpr['p'][0])
                except Exception: pass
            rec['legend'] = l
        else:
            rec['legend'] = None

        # 数据标签（series 级 + 点级）
        ser_recs = []
        for cg in ch.get('plotArea', {}).get('chartGroups', []):
            for ser in cg.get('ser', []):
                dl = ser.get('dLbls')
                srec = {'dLbls': None}
                if dl:
                    srec['dLbls'] = {'showVal': dl.get('showVal'), 'showPercent': dl.get('showPercent'),
                                     'showCatName': dl.get('showCatName'), 'showSerName': dl.get('showSerName'),
                                     'showBubbleSize': dl.get('showBubbleSize'), 'dLblPos': dl.get('dLblPos'),
                                     'numFmt': (dl.get('numFmt') or {}).get('formatCode')}
                    try: srec['dLbls']['font'] = rpr_info(dl['txPr']['p'][0])
                    except Exception: pass
                    pts = dl.get('dLbl')
                    if pts: srec['point_dLbls'] = [{ 'idx': p.get('idx'), 'tx': (text_of(p) if isinstance(p, dict) else None)} for p in pts]
                ser_recs.append(srec)
        rec['series'] = ser_recs
        out.append(rec)

json.dump(out, open('batch1_sjs_summary.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
for r in out:
    t = r['title']
    tdesc = t.get('text') if isinstance(t, dict) else t
    if tdesc is None and r['title_deleted']: tdesc = 'DELETED'
    print(r['sheet'], '|', r['name'], '| title=', tdesc,
          '| axes=', len(r['axes']), '| legend=', 'yes' if r['legend'] else 'none')
