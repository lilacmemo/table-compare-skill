# -*- coding: utf-8 -*-
"""
对 build_samples_all.ps1 生成的工作簿做 OOXML 级补丁。

Excel COM 的 FormatConditions.Add 有 3 个可选参数，PowerShell/后期绑定无法传
DISP_E_PARAMNOTFOUND，因此"特定文本""日期发生"两类规则无法用 COM 创建；
DataBar 的 min/max 类型与方向在该版本 COM 上不可写。
这里直接改 XML 补齐，其余内容原样保留（不能用 openpyxl 保存，会丢 x14 扩展）。

用法: python patch_cf.py [xlsx路径]
"""
import re
import shutil
import sys
import zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else '格式与条件格式-全量用例-A至K.xlsx'
TMP = SRC + '.tmp'

G_SHEET = 'G条件格式规则'
H_SHEET = 'H条件格式参数'


def load_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def save_zip(path, items, order):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in order:
            z.writestr(n, items[n])


def sheet_file_map(items):
    wb = items['xl/workbook.xml'].decode('utf-8')
    rels = items['xl/_rels/workbook.xml.rels'].decode('utf-8')
    pairs = re.findall(r'<sheet\b[^>]*/>', wb)
    name2rid = {}
    for p in pairs:
        nm = re.search(r'name="([^"]*)"', p)
        rid = re.search(r'r:id="([^"]*)"', p)
        if nm and rid:
            name2rid[nm.group(1)] = rid.group(1)
    rid2target = {}
    for p in re.findall(r'<Relationship\b[^>]*/>', rels):
        i = re.search(r'Id="([^"]*)"', p)
        t = re.search(r'Target="([^"]*)"', p)
        if i and t:
            rid2target[i.group(1)] = t.group(1)
    out = {}
    for nm, rid in name2rid.items():
        t = rid2target.get(rid, '')
        if not t:
            continue
        t = t.lstrip('/')
        out[nm] = t if t.startswith('xl/') else 'xl/' + t
    return out


def add_dxfs(styles_xml, fills):
    """在 <dxfs> 末尾追加 dxf，返回 (新xml, 起始索引)"""
    m = re.search(r'<dxfs count="(\d+)"[^>]*>(.*?)</dxfs>', styles_xml, re.S)
    if not m:
        raise RuntimeError('styles.xml 里找不到 <dxfs>')
    count = int(m.group(1))
    add = ''.join(
        '<dxf><fill><patternFill><bgColor rgb="%s"/></patternFill></fill></dxf>' % c
        for c in fills)
    new_dxfs = '<dxfs count="%d">%s%s</dxfs>' % (count + len(fills), m.group(2), add)
    return styles_xml[:m.start()] + new_dxfs + styles_xml[m.end():], count


def patch_g(xml, dxf_green, dxf_yellow):
    if 'sqref="C4:H4"' in xml and 'containsText' in xml:
        return xml, 'G2 已存在，跳过'
    rules = (
        '<conditionalFormatting sqref="C4:H4">'
        '<cfRule type="containsText" dxfId="%d" priority="90" operator="containsText" text="合格">'
        '<formula>NOT(ISERROR(SEARCH("合格",C4)))</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="C5:H5">'
        '<cfRule type="timePeriod" dxfId="%d" priority="91" timePeriod="today">'
        '<formula>FLOOR(C5,1)=TODAY()</formula></cfRule></conditionalFormatting>'
        % (dxf_green, dxf_yellow))
    i = xml.find('<conditionalFormatting')
    if i < 0:
        raise RuntimeError('G 表里找不到 conditionalFormatting 锚点')
    return xml[:i] + rules + xml[i:], 'G2/G3 已注入'


def patch_h(xml):
    log = []

    def set_cfvo(block, idx, typ, val=None):
        # 依次替换第 idx 个 <x14:cfvo .../>
        parts = list(re.finditer(r'<x14:cfvo\b[^>]*/>', block))
        if len(parts) <= idx:
            return block
        old = parts[idx].group(0)
        new = '<x14:cfvo type="%s" val="%s"/>' % (typ, val) if val is not None else '<x14:cfvo type="%s"/>' % typ
        return block[:parts[idx].start()] + new + block[parts[idx].end():]

    def set_base_cfvo(block, idx, typ, val=None):
        parts = list(re.finditer(r'<cfvo\b[^>]*/>', block))
        if len(parts) <= idx:
            return block
        old = parts[idx].group(0)
        new = '<cfvo type="%s" val="%s"/>' % (typ, val) if val is not None else '<cfvo type="%s"/>' % typ
        return block[:parts[idx].start()] + new + block[parts[idx].end():]

    jobs = [
        ('C3:H3', 0, 'num', '0', None),        # H1 最小值 = 数字 0
        ('C4:H4', 1, 'num', '100', None),      # H2 最大值 = 数字 100
        ('C9:H9', None, None, None, 'leftToRight'),   # H6-1 方向
        ('C10:H10', None, None, None, 'rightToLeft'),  # H6-2 方向
        ('C6:H6', None, None, None, None),     # H4 负值轴（单独处理）
    ]
    for sqref, idx, typ, val, direction in jobs:
        # x14 扩展块
        pat = re.compile(
            r'(<x14:conditionalFormatting\b[^>]*>(?:(?!</x14:conditionalFormatting>).)*?'
            r'<xm:sqref>%s</xm:sqref></x14:conditionalFormatting>)' % re.escape(sqref), re.S)
        m = pat.search(xml)
        if not m:
            log.append('%s 未找到 x14 块' % sqref)
            continue
        block = m.group(1)
        nb = block
        if idx is not None:
            nb = set_cfvo(nb, idx, typ, val)
        if direction:
            md = re.search(r'<x14:dataBar\b[^>]*>', nb)
            if md:
                tag = md.group(0)
                if 'direction="' in tag:
                    newtag = re.sub(r'direction="[^"]*"', 'direction="%s"' % direction, tag)
                else:
                    newtag = tag[:-1].rstrip(' /') + ' direction="%s">' % direction
                nb = nb[:md.start()] + newtag + nb[md.end():]
        if sqref == 'C6:H6':
            nb = nb.replace('axisPosition="none"', 'axisPosition="middle"')
            nb = nb.replace('negativeBarColorSameAsPositive="1"', 'negativeBarColorSameAsPositive="0"')
        if nb != block:
            xml = xml[:m.start(1)] + nb + xml[m.end(1):]
            log.append('%s x14 已改' % sqref)

        # 兼容层（cfRule 内的 <dataBar>）
        if idx is not None:
            pat2 = re.compile(
                r'(<conditionalFormatting sqref="%s">(?:(?!</conditionalFormatting>).)*?</conditionalFormatting>)'
                % re.escape(sqref), re.S)
            m2 = pat2.search(xml)
            if m2:
                b2 = m2.group(1)
                nb2 = set_base_cfvo(b2, idx, typ, val)
                if nb2 != b2:
                    xml = xml[:m2.start(1)] + nb2 + xml[m2.end(1):]
    return xml, '; '.join(log) if log else 'H 无可改项'


def main():
    items = load_zip(SRC)
    order = list(items.keys())
    fmap = sheet_file_map(items)
    styles = items['xl/styles.xml'].decode('utf-8')

    # 先算 dxf 索引（在追加之前拿到 count）
    m = re.search(r'<dxfs count="(\d+)"', styles)
    base = int(m.group(1)) if m else 0
    GREEN, YELLOW = base, base + 1

    styles2, _ = add_dxfs(styles, ['FFC6EFCE', 'FFFFEB9C'])
    items['xl/styles.xml'] = styles2.encode('utf-8')

    gf = fmap[G_SHEET]
    xml = items[gf].decode('utf-8')
    xml2, msg1 = patch_g(xml, GREEN, YELLOW)
    items[gf] = xml2.encode('utf-8')

    hf = fmap[H_SHEET]
    xmlh = items[hf].decode('utf-8')
    xmlh2, msg2 = patch_h(xmlh)
    items[hf] = xmlh2.encode('utf-8')

    save_zip(TMP, items, order)
    shutil.move(TMP, SRC)

    print('dxf 追加: 绿=%d 黄=%d' % (GREEN, YELLOW))
    print('G 表:', msg1)
    print('H 表:', msg2)
    print('done ->', SRC)


if __name__ == '__main__':
    main()
