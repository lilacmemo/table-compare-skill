# -*- coding: utf-8 -*-
"""
工作簿/工作表样本的三重核验 · 第三重：包级（解包数 XML 节点）

第一重是 COM 回读（annotate_worksheet_sample.ps1，23 项断言）
第二重是 Excel 打开并重存（同一脚本的 SaveAs，证明 Excel 认这个文件）
第三重就是本脚本：直接数包里的节点，确认每个载体真的落进了文件。

口径：只断言"样本里应当存在的东西确实存在"，不看渲染效果。
用法：python verify_sample_worksheet.py [xlsx 路径]
"""
import re
import sys
import zipfile

PATH = sys.argv[1] if len(sys.argv) > 1 else r"C:/temp/ws_sample_com.xlsx"

NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

z = zipfile.ZipFile(PATH)
parts = set(z.namelist())

wbx = z.read("xl/workbook.xml").decode("utf-8")
rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")

rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
rid_target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})

sheets = []  # (name, state, xml_path)
for m in re.finditer(r"<sheet\b([^>]*)/>", wbx):
    attrs = m.group(1)
    name = re.search(r'name="([^"]*)"', attrs).group(1)
    state = re.search(r'state="([^"]*)"', attrs)
    rid = re.search(r'r:id="([^"]+)"', attrs) or re.search(r'id="([^"]+)"', attrs)
    tgt = rid_target.get(rid.group(1), "")
    path = "xl/" + tgt.lstrip("/").replace("xl/", "") if tgt else ""
    if path and not path.startswith("xl/"):
        path = "xl/" + path
    sheets.append((name, state.group(1) if state else "visible", path))

results = []


def check(cid, what, ok, detail=""):
    results.append((cid, what, "PASS" if ok else "FAIL", detail))


def sheet_xml(name):
    for n, _s, p in sheets:
        if n == name:
            return z.read(p).decode("utf-8") if p in parts else ""
    return ""


check("S1", "sheet 总数 = 15", len(sheets) == 15, "实际 %d" % len(sheets))
hidden = [n for n, s, _ in sheets if s == "hidden"]
vhidden = [n for n, s, _ in sheets if s == "veryHidden"]
check("B2", "普通隐藏表 1 个", len(hidden) == 1, str(hidden))
check("B3", "深度隐藏表 1 个", len(vhidden) == 1, str(vhidden))

x4 = sheet_xml("03-保护")
styles = z.read("xl/styles.xml").decode("utf-8")
check("C1", "保护已开启 (<sheetProtection sheet=\"1\")",
      'sheet="1"' in x4 and "<sheetProtection" in x4)
# Excel 写的是旧式 password 哈希（4 位十六进制）；只有用强加密时才是 hashValue
prot = re.search(r"<sheetProtection[^>]*/>", x4)
prot_s = prot.group(0) if prot else ""
check("C2", "保护带密码 (password=/hashValue=)",
      ('password="' in prot_s) or ('hashValue="' in prot_s), prot_s)
for f, cid in [("insertRows", "C4"), ("insertColumns", "C4"),
               ("sort", "C6"), ("autoFilter", "C6"), ("formatCells", "C7")]:
    check(cid, "允许 %s（属性写 0）" % f, re.search(r'%s="0"' % f, prot_s) is not None)
# OOXML 里“禁止”= 属性缺省（默认 true），所以判据是属性不出现
for f, cid in [("deleteRows", "C5"), ("deleteColumns", "C5"),
               ("formatRows", "C8"), ("formatColumns", "C8"), ("insertHyperlinks", "C9")]:
    check(cid, "禁止 %s（属性缺省）" % f, ('%s=' % f) not in prot_s)
# 单元格锁定是样式层属性，落在 styles.xml 的 <protection locked="0"/>
check("C3", "解锁单元格样式 (styles.xml protection locked=0)",
      '<protection locked="0"/>' in styles)

x5 = sheet_xml("04-冻结")
pane = re.search(r"<pane[^>]*/>", x5)
check("D1", "冻结窗格 xSplit=3 ySplit=5",
      bool(pane) and 'xSplit="3"' in pane.group(0) and 'ySplit="5"' in pane.group(0),
      pane.group(0) if pane else "无 <pane>")

x6 = sheet_xml("05-分组")
lvl = dict(re.findall(r'<row r="(\d+)"[^>]*outlineLevel="(\d+)"', x6))
check("E1", "行分组 1 级 / 2 级嵌套",
      lvl.get("3") == "1" and lvl.get("7") == "2", str(sorted(lvl.items(), key=lambda kv: int(kv[0]))))
check("E2", "汇总行在上 (summaryBelow=0)", 'summaryBelow="0"' in x6)
check("E3", "列分组存在", 'outlineLevel' in re.sub(r"<row[^>]*>", "", x6))

x7 = sheet_xml("06-行列")
check("F2", "隐藏列 2 个 + 隐藏行 2 个",
      x7.count('hidden="1"') >= 4 or x7.count('hidden="true"') >= 4,
      "hidden 属性出现 %d 次" % (x7.count('hidden="1"') + x7.count('hidden="true"')))
check("F1", "行 4 自定义行高", re.search(r'<row r="4"[^>]*ht="36"', x7) is not None)
check("F3", "B 列自定义列宽", re.search(r'<col[^>]*min="2"[^>]*width="2[5-7]', x7) is not None)
check("F4", "合并单元格存在 (mergeCells)", "<mergeCell" in x7)

x8 = sheet_xml("07-视图")
check("G1", "网格线关闭 (showGridLines=0)", 'showGridLines="0"' in x8)
check("G3", "缩放 80 (zoomScale=80)", 'zoomScale="80"' in x8)
check("G4", "标签色存在 (tabColor)", "tabColor" in x8)

x12 = sheet_xml("11-RTL")
check("G5", "RTL (rightToLeft=1)", 'rightToLeft="1"' in x12 or 'rightToLeft="true"' in x12)

x13 = sheet_xml("12-背景图")
check("G6", "工作表背景图 (<picture r:id>)", "<picture" in x13 and "r:id" in x13)

check("H1", "批注部件 2 个", len([p for p in parts if re.match(r"xl/comments\d*\.xml", p)]) == 2)
check("H2", "VML 批注绘图 2 个", len([p for p in parts if p.endswith(".vml")]) == 2)
cmt1 = "".join(z.read(p).decode("utf-8") for p in parts if re.match(r"xl/comments1\.xml", p))
check("H3", "批注文本含关键字「促销」", "促销" in cmt1)

names = re.findall(r'<definedName name="([^"]+)"[^>]*>([^<]*)</definedName>', wbx)
nm = dict(names)
check("I2", "定义名称 TaxRate / DataBlock 都在",
      "TaxRate" in nm and "DataBlock" in nm, str(names))

core = z.read("docProps/core.xml").decode("utf-8")
for tag, cid in [("dc:title", "I1"), ("dc:creator", "I1"), ("dc:subject", "I1"),
                 ("cp:keywords", "I1"), ("dc:description", "I1")]:
    check(cid, "core.xml 含 %s" % tag, "<%s" % tag in core)
check("I1", "自定义文档属性 (custom.xml)", "docProps/custom.xml" in parts)
if "docProps/custom.xml" in parts:
    check("I1", "custom.xml 含 CompareItem", "CompareItem" in z.read("docProps/custom.xml").decode())

x11 = sheet_xml("10-区域")
check("L4", "数据验证存在 (dataValidation)", "<dataValidation" in x11)
check("L5", "仅样式单元格 H3 (无值有样式)",
      re.search(r'<c r="H3"[^>]*s="\d+"', x11) is not None)

media = [p for p in parts if p.startswith("xl/media/")]
check("G6", "嵌入图片 1 个 (xl/media)", len(media) == 1, str(media))

w = max(len(r[1]) for r in results)
print("file:", PATH)
print("sheets: %d (visible %d / hidden %d / veryHidden %d)" % (
    len(sheets), len(sheets) - len(hidden) - len(vhidden), len(hidden), len(vhidden)))
print("-" * (w + 22))
for cid, what, res, detail in results:
    print("%-4s %-*s %-6s %s" % (cid, w, what, res, detail))
bad = [r for r in results if r[2] == "FAIL"]
print("-" * (w + 22))
print("verify: %d checks, %d pass, %d fail" % (len(results), len(results) - len(bad), len(bad)))
sys.exit(1 if bad else 0)
