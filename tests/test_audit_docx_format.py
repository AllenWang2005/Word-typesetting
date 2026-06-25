"""Unit tests for scripts/audit_docx_format.py.

Standard library only (unittest) so CI needs no third-party dependencies:

    python -m unittest discover -s tests -v
"""

from __future__ import annotations

import importlib.util
import pathlib
import sys
import unittest
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


def make_p(text: str, jc: str | None = None, first_line: str | None = None) -> ET.Element:
    ppr = ""
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
            make_p("经济指标表示该方案 P 值更优。"),
            make_p("下面取 N 个样本进行统计分析。"),
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


if __name__ == "__main__":
    unittest.main()
