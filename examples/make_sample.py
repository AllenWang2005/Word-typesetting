#!/usr/bin/env python3
"""Build a deliberately NON-COMPLIANT sample DOCX to demonstrate the auditor.

Requires python-docx (`pip install python-docx`).

Usage:
    python examples/make_sample.py [output.docx]
    python scripts/audit_docx_format.py output.docx

The generated document intentionally violates several rules so the audit script
has something to report:
  * ASCII punctuation inside Chinese prose
  * a centered level-1 heading
  * a centered / indented abstract paragraph
  * table text that is not small-four (12 pt)
  * full-width and field-less citation brackets
  * plain-text formula / visible LaTeX instead of OMML
  * non-black body font color
"""

from __future__ import annotations

import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor


def build(path: str) -> None:
    doc = Document()

    # Centered level-1 heading (should be left aligned).
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run("一、绪论")

    # Abstract that is centered and indented (should be left aligned, indent 0).
    abs = doc.add_paragraph()
    abs.alignment = WD_ALIGN_PARAGRAPH.CENTER
    abs.paragraph_format.first_line_indent = Pt(24)
    abs.add_run("摘要 本文研究梯级水库的优化调度方法与防洪效益评估。")

    # Body prose with ASCII punctuation in Chinese text, plus a non-black run.
    p = doc.add_paragraph()
    p.add_run("本设计针对某流域的水利计算问题展开分析,采用多方案比较的方法进行评价.")
    red = p.add_run("（本句为红色正文，属于不应保留的彩色字体）")
    red.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    # Plain-text formula and visible LaTeX (should be native OMML).
    doc.add_paragraph("式中 N_p 为水泵台数，反映系统的装机规模。")
    doc.add_paragraph("两者比值可用 \\frac{a}{b} 表示，需要渲染为公式对象。")

    # Citations with full-width brackets and a field-less ASCII bracket.
    doc.add_paragraph("已有研究表明该方法有效［1］，并在工程中得到验证[2]。")

    # A bare superscript citation number that dropped its brackets (should be [3]).
    bare = doc.add_paragraph()
    bare.add_run("该结论亦见于相关文献")
    sup = bare.add_run("3")
    sup.font.superscript = True
    bare.add_run("。")

    # A table whose text is not small-four (12 pt).
    table = doc.add_table(rows=1, cols=2)
    for idx, label in enumerate(("方案", "流量")):
        cell = table.cell(0, idx)
        run = cell.paragraphs[0].add_run(label)
        run.font.size = Pt(14)  # -> w:sz=28, not the required 24

    doc.save(path)
    print(f"wrote {path}")


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "sample.docx"
    build(out)
