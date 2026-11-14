# -*- coding: utf-8 -*-
"""重建 成果展示工程/index.html：多份 md 原样渲染 + 图片 base64 内嵌 + 重建左侧导航。
用法: python poc-compare/build_showcase.py
说明: 复用现有 index.html 的 <head>/<style>/<script> 骨架（只替换 <nav> 与 <main> 内容），
      因此样式调整只需改 index.html 一次即长期生效；新增模块只改 DOCS。
"""
import base64
import os
import re
import sys

import markdown

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "成果展示工程", "index.html")
# (doc id, md 文件名, 导航显示名)
DOCS = [
    ("chart", "SpreadJS vs Univer 图表能力对比报告（合并版）.md", "图表能力对比"),
    ("formula", "SpreadJS vs Univer 公式能力验证报告（终版）.md", "公式能力验证"),
    ("format", "SpreadJS vs Univer 格式与条件格式 API 对比报告.md", "格式与条件格式"),
    ("table", "SpreadJS vs Univer 表-排序-筛选-验证 API 对比报告.md", "表 / 排序 / 筛选 / 验证"),
    ("print", "SpreadJS vs Univer 打印与 PDF 导出对比报告.md", "打印与 PDF 导出"),
]

missing = []


def inline_images(html):
    def repl(m):
        src = m.group(1)
        if src.startswith("data:"):
            return m.group(0)
        path = os.path.join(ROOT, src.replace("/", os.sep))
        if not os.path.exists(path):
            missing.append(src)
            return m.group(0)
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return m.group(0).replace(src, "data:image/png;base64," + b64)

    return re.sub(r'<img[^>]+src="([^"]+)"', repl, html)


def render(md_path):
    with open(os.path.join(ROOT, md_path), encoding="utf-8") as f:
        text = f.read()
    html = markdown.markdown(text, extensions=["tables", "fenced_code", "toc", "md_in_html"])
    return inline_images(html)


def main():
    if not os.path.exists(OUT):
        sys.exit("缺少 index.html 骨架，请先手工建一次")
    with open(OUT, encoding="utf-8") as f:
        old = f.read()

    head = old[:old.index("<nav>")]
    tail = old[old.index("</main>") + len("</main>"):]

    nav = ['<nav>\n  <h1>SpreadJS vs Univer<br>调研报告</h1>\n']
    for idx, (key, _md, label) in enumerate(DOCS):
        cls = ' class="active"' if idx == 0 else ""
        nav.append('  <a data-target="%s"%s>%s</a>\n' % (key, cls, label))
    nav.append("</nav>\n<main>\n")

    docs = []
    for idx, (key, md, _label) in enumerate(DOCS):
        cls = "doc visible" if idx == 0 else "doc"
        docs.append('<div class="%s" id="doc-%s">\n%s\n</div>\n' % (cls, key, render(md)))

    new = head + "".join(nav) + "".join(docs) + "</main>" + tail

    with open(OUT, "w", encoding="utf-8") as f:
        f.write(new)
    print("docs: %d  ->  %s" % (len(DOCS), ", ".join(k for k, _, _ in DOCS)))
    print("written: %.2f MB" % (len(new.encode("utf-8")) / 1024 / 1024))
    print("img embedded: %d" % new.count("data:image/png;base64,"))
    if missing:
        print("MISSING:", missing)


if __name__ == "__main__":
    main()
