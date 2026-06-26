"""Unit tests for scripts/audit_docx_format.py.

Standard library only (unittest) so CI needs no third-party dependencies:

    python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import os
import pathlib
import sys
import tempfile
import unittest
import zipfile
from typing import Optional
from xml.etree import ElementTree as ET


def _load_audit_module():
    path = pathlib.Path(__file__).resolve().parent.parent / "scripts" / "audit_docx_format.py"
    spec = importlib.util.spec_from_file_location("audit_docx_format", path)
    module = importlib.util.module_from_spec(spec)
    # Register before exec so dataclasses can resolve the module's namespace.
    sys.modules["audit_docx_format"] = module
    spec.loader.exec_module(module)
    return module


aud = _load_audit_module()

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def make_p(text: str, jc: Optional[str] = None, first_line: Optional[str] = None, style: Optional[str] = None) -> ET.Element:
    ppr = ""
    if style:
        ppr += f'<w:pStyle w:val="{style}"/>'
    if jc:
        ppr += f'<w:jc w:val="{jc}"/>'
    if first_line:
        ppr += f'<w:ind w:firstLine="{first_line}"/>'
    if ppr:
        ppr = f"<w:pPr>{ppr}</w:pPr>"
    xml = (
        f'<w:p xmlns:w="{W}">{ppr}'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )
    return ET.fromstring(xml)


def make_runs_p(parts) -> ET.Element:
    """Build a paragraph from a list of (text, is_superscript) tuples."""
    runs = ""
    for text, sup in parts:
        rpr = '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>' if sup else ""
        runs += f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
    return ET.fromstring(f'<w:p xmlns:w="{W}">{runs}</w:p>')


def make_p_with_font(text: str, eastasia: str, style: Optional[str] = None) -> ET.Element:
    pstyle = f'<w:pStyle w:val="{style}"/>' if style else ""
    ppr = f"<w:pPr>{pstyle}</w:pPr>" if pstyle else ""
    xml = (
        f'<w:p xmlns:w="{W}">{ppr}'
        f'<w:r><w:rPr><w:rFonts w:eastAsia="{eastasia}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )
    return ET.fromstring(xml)


def make_doc(body_inner: str) -> ET.Element:
    return ET.fromstring(f'<w:document xmlns:w="{W}"><w:body>{body_inner}</w:body></w:document>')


TABLE_XML = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"


def codes(issues) -> set[str]:
    return {issue.code for issue in issues}


class DetectLanguageTests(unittest.TestCase):
    def test_chinese_heavy_is_zh(self):
        language, _, _ = aud.detect_language("中" * 60)
        self.assertEqual(language, "zh")

    def test_english_is_en(self):
        language, _, _ = aud.detect_language("This is a fully English sentence with words.")
        self.assertEqual(language, "en")


class PunctuationTests(unittest.TestCase):
    def test_ascii_comma_between_cjk_flagged(self):
        issues = []
        aud.audit_punctuation([make_p("这是中文,后面还是中文。")], "zh", issues)
        self.assertIn("ZH_PUNCT", codes(issues))

    def test_protected_tokens_not_flagged(self):
        issues = []
        aud.audit_punctuation([make_p("详见网址 https://example.com/a 与 doi:10.1000/xyz 。")], "zh", issues)
        self.assertNotIn("ZH_PUNCT", codes(issues))

    def test_chinese_punct_in_english_doc_flagged(self):
        issues = []
        aud.audit_punctuation([make_p("This sentence ends with a Chinese period。")], "en", issues)
        self.assertIn("EN_PUNCT", codes(issues))


class FormulaTests(unittest.TestCase):
    def test_subscript_quantity_flagged(self):
        issues = []
        aud.audit_formulas([make_p("式中 N_p 为水泵台数。")], issues)
        self.assertIn("FORMULA_TEXT", codes(issues))

    def test_generic_symbol_not_domain_specific(self):
        # A symbol unrelated to any hydraulics vocabulary must still be detected,
        # proving the auditor is domain-neutral.
        issues = []
        aud.audit_formulas([make_p("式中 G_k 为永久荷载标准值。")], issues)
        self.assertIn("FORMULA_TEXT", codes(issues))

    def test_bare_quantity_symbols_in_definition_context_flagged(self):
        issues = []
        aud.audit_formulas([make_p("式中 Q 为流量，N 为出力。")], issues)
        self.assertIn("FORMULA_TEXT", codes(issues))

    def test_bare_quantity_symbol_with_definition_verb_flagged(self):
        issues = []
        aud.audit_formulas([make_p("其中 H 表示水头，V 为流速。")], issues)
        self.assertIn("FORMULA_TEXT", codes(issues))

    def test_visible_latex_flagged(self):
        issues = []
        aud.audit_formulas([make_p("比值可用 \\frac{a}{b} 计算。")], issues)
        self.assertIn("VISIBLE_LATEX", codes(issues))

    def test_ordinary_prose_not_flagged(self):
        # Regression: plain prose that merely contains common words like 为/取/表示
        # plus an incidental capital letter must not be treated as a formula.
        issues = []
        prose = [
            make_p("本方案的目标为提高发电效益，T 型布置更合理。"),
            make_p("经济指标表示该方案整体更优。"),
            make_p("下面选取三个样本进行统计分析。"),
        ]
        aud.audit_formulas(prose, issues)
        self.assertEqual(codes(issues), set())


class CitationTests(unittest.TestCase):
    def test_fullwidth_brackets_flagged(self):
        issues = []
        aud.audit_citations("", "参见文献［1］的结论。", issues)
        self.assertIn("CITATION_BRACKETS", codes(issues))

    def test_ascii_citation_without_field_flagged(self):
        issues = []
        aud.audit_citations("<w:document></w:document>", "参见文献[1]的结论。", issues)
        self.assertIn("CITATION_FIELDS", codes(issues))

    def test_ascii_citation_with_ref_field_ok(self):
        issues = []
        xml = '<w:document> ... <w:fldSimple w:instr=" REF ref_001 \\h "/> ... </w:document>'
        aud.audit_citations(xml, "参见文献[1]的结论。", issues)
        self.assertNotIn("CITATION_FIELDS", codes(issues))


class HeadingTests(unittest.TestCase):
    def test_heading1_style_centered_flagged(self):
        issues = []
        aud.audit_headings([make_p("Introduction", jc="center", style="Heading1")], issues)
        self.assertIn("H1_CENTER", codes(issues))

    def test_chinese_title_style_centered_flagged(self):
        issues = []
        aud.audit_headings([make_p("绪论", jc="center", style="标题1")], issues)
        self.assertIn("H1_CENTER", codes(issues))

    def test_arabic_heading_text_centered_flagged(self):
        issues = []
        aud.audit_headings([make_p("1 绪论", jc="center")], issues)
        self.assertIn("H1_CENTER", codes(issues))


class BareCitationTests(unittest.TestCase):
    def test_bare_superscript_number_after_cjk_flagged(self):
        issues = []
        aud.audit_bare_citations([make_runs_p([("研究表明该方法可行", False), ("1", True), ("。", False)])], issues)
        self.assertIn("CITATION_NO_BRACKETS", codes(issues))

    def test_bracketed_superscript_citation_ok(self):
        issues = []
        aud.audit_bare_citations([make_runs_p([("研究表明该方法可行", False), ("[1]", True), ("。", False)])], issues)
        self.assertNotIn("CITATION_NO_BRACKETS", codes(issues))

    def test_superscript_exponent_not_flagged(self):
        # "m" + superscript "2" is an exponent (m²), not a citation.
        issues = []
        aud.audit_bare_citations([make_runs_p([("断面面积为 100 m", False), ("2", True), ("。", False)])], issues)
        self.assertNotIn("CITATION_NO_BRACKETS", codes(issues))


class HeadingPunctuationTests(unittest.TestCase):
    def test_heading_with_trailing_punctuation_flagged(self):
        issues = []
        aud.audit_headings([make_p("二、研究方法。", style="Heading1")], issues)
        self.assertIn("HEADING_PUNCT", codes(issues))

    def test_heading_without_trailing_punctuation_ok(self):
        issues = []
        aud.audit_headings([make_p("二、研究方法", style="Heading1")], issues)
        self.assertNotIn("HEADING_PUNCT", codes(issues))


class FontTests(unittest.TestCase):
    def test_heading_in_songti_flagged(self):
        issues = []
        aud.audit_fonts([make_p_with_font("研究方法", "宋体", style="Heading1")], issues)
        self.assertIn("HEADING_FONT", codes(issues))

    def test_body_in_heiti_flagged(self):
        issues = []
        aud.audit_fonts([make_p_with_font("这是一段正文内容", "黑体")], issues)
        self.assertIn("BODY_FONT", codes(issues))

    def test_level3_heading_in_songti_not_flagged(self):
        # Level-3 headings are intentionally bold Songti, so Songti must not be flagged.
        issues = []
        aud.audit_fonts([make_p_with_font("研究细节", "宋体", style="Heading3")], issues)
        self.assertNotIn("HEADING_FONT", codes(issues))

    def test_correct_fonts_ok(self):
        issues = []
        aud.audit_fonts(
            [make_p_with_font("研究方法", "黑体", style="Heading1"), make_p_with_font("这是正文", "宋体")],
            issues,
        )
        self.assertEqual(codes(issues), set())


class CaptionPositionTests(unittest.TestCase):
    def test_table_caption_below_table_flagged(self):
        root = make_doc(TABLE_XML + '<w:p><w:r><w:t>表 1-1 方案比较</w:t></w:r></w:p>')
        issues = []
        aud.audit_captions(root, issues)
        self.assertIn("CAPTION_POSITION", codes(issues))

    def test_table_caption_above_table_ok(self):
        root = make_doc('<w:p><w:r><w:t>表 1-1 方案比较</w:t></w:r></w:p>' + TABLE_XML)
        issues = []
        aud.audit_captions(root, issues)
        self.assertNotIn("CAPTION_POSITION", codes(issues))


class TableSizeTests(unittest.TestCase):
    def _table(self, sz: str) -> ET.Element:
        return make_doc(
            f'<w:tbl><w:tr><w:tc><w:p><w:r><w:rPr><w:sz w:val="{sz}"/></w:rPr>'
            f'<w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )

    def test_wuhao_accepted(self):
        issues = []
        aud.audit_tables(self._table("21"), issues)  # 五号
        self.assertNotIn("TABLE_SIZE", codes(issues))

    def test_xiaosi_accepted(self):
        issues = []
        aud.audit_tables(self._table("24"), issues)  # 小四
        self.assertNotIn("TABLE_SIZE", codes(issues))

    def test_oversized_table_text_flagged(self):
        issues = []
        aud.audit_tables(self._table("28"), issues)  # 四号, larger than body
        self.assertIn("TABLE_SIZE", codes(issues))


class ColorTests(unittest.TestCase):
    def test_hyperlink_color_not_flagged_but_body_color_is(self):
        root = make_doc(
            '<w:p><w:hyperlink><w:r><w:rPr><w:color w:val="0563C1"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink></w:p>'
            '<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>red</w:t></w:r></w:p>'
        )
        issues = []
        aud.audit_color(root, issues)
        self.assertEqual(sum(1 for i in issues if i.code == "COLOR"), 1)


class InputErrorTests(unittest.TestCase):
    def test_non_docx_raises_audit_input_error(self):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as handle:
            handle.write(b"this is not a zip package")
            path = handle.name
        try:
            with self.assertRaises(aud.AuditInputError):
                aud.load_document_xml(pathlib.Path(path))
        finally:
            os.unlink(path)


class JsonOutputTests(unittest.TestCase):
    def test_json_output_has_expected_keys(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "min.docx")
            document = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document xmlns:w="{W}"><w:body>'
                f'<w:p><w:r><w:t>这是一段用于测试输出结构的正常中文正文内容，全部使用中文标点符号，'
                f'并且不包含任何明显的排版问题，因此机器审计的结果应当顺利通过检查。</w:t></w:r></w:p>'
                f'</w:body></w:document>'
            )
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", document)
            saved_argv = sys.argv
            sys.argv = ["audit", path, "--json"]
            buffer = io.StringIO()
            try:
                with contextlib.redirect_stdout(buffer):
                    return_code = aud.main()
            finally:
                sys.argv = saved_argv
        payload = json.loads(buffer.getvalue())
        for key in ("status", "scope", "summary", "issues"):
            self.assertIn(key, payload)
        self.assertEqual(return_code, 0)


class SummaryTests(unittest.TestCase):
    def test_omitted_issue_count_reported(self):
        issues = []
        counts = {}
        for index in range(10):
            aud.add_issue(issues, "FAIL", "ZH_PUNCT", f"issue {index}", counts)
        summary, omitted = aud.summarize_issues(issues, counts)
        self.assertEqual(len(issues), 8)
        self.assertEqual(summary["FAIL"], 10)
        self.assertEqual(summary["omitted"], 2)
        self.assertEqual(omitted[0]["code"], "ZH_PUNCT")


if __name__ == "__main__":
    unittest.main()
