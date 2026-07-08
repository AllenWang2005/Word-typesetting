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
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"



CONTENT_TYPES = (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
)
DOCUMENT_RELS = (
    '<?xml version="1.0"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)


def write_minimal_docx(path: str, document: str) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        archive.writestr("word/document.xml", document)


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


def make_p_with_font(text: str, eastasia: str, style: Optional[str] = None, jc: Optional[str] = None) -> ET.Element:
    inner = ""
    if style:
        inner += f'<w:pStyle w:val="{style}"/>'
    if jc:
        inner += f'<w:jc w:val="{jc}"/>'
    ppr = f"<w:pPr>{inner}</w:pPr>" if inner else ""
    xml = (
        f'<w:p xmlns:w="{W}">{ppr}'
        f'<w:r><w:rPr><w:rFonts w:eastAsia="{eastasia}"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )
    return ET.fromstring(xml)


def run_main_json(body_inner: str) -> dict:
    """Write a minimal DOCX with the given <w:body> inner XML and run main --json."""
    with tempfile.TemporaryDirectory() as folder:
        path = os.path.join(folder, "doc.docx")
        document = f'<w:document xmlns:w="{W}"><w:body>{body_inner}</w:body></w:document>'
        write_minimal_docx(path, document)
        saved_argv = sys.argv
        sys.argv = ["audit", path, "--json", "--language", "zh"]
        buffer = io.StringIO()
        try:
            with contextlib.redirect_stdout(buffer):
                aud.main()
        finally:
            sys.argv = saved_argv
    return json.loads(buffer.getvalue())


def make_doc(body_inner: str) -> ET.Element:
    return ET.fromstring(f'<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>{body_inner}</w:body></w:document>')


def make_styles(styles_inner: str) -> ET.Element:
    return ET.fromstring(f'<w:styles xmlns:w="{W}">{styles_inner}</w:styles>')


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

    def test_label_prose_not_flagged(self):
        # Regression: "为 A 方案" / "取 N 个" are labels, not symbol definitions.
        issues = []
        aud.audit_formulas([make_p("本设计的最优方案为 A 方案，故取 N 个代表性断面。")], issues)
        self.assertNotIn("FORMULA_TEXT", codes(issues))


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

    def test_centered_heiti_cover_title_not_flagged(self):
        # The cover title is Heiti by spec; being centered, it must not be flagged as body.
        issues = []
        aud.audit_fonts([make_p_with_font("某某水库水利计算课程设计", "黑体", jc="center")], issues)
        self.assertNotIn("BODY_FONT", codes(issues))


class HeadingStyleFontTests(unittest.TestCase):
    def test_localized_numeric_heading_style_id_level2_and_3(self):
        self.assertEqual(aud.heading_level(make_p("1.1 Method", style="21"), "1.1 Method"), 2)
        self.assertEqual(aud.heading_level(make_p("1.1.1 Detail", style="31"), "1.1.1 Detail"), 3)

    def test_used_heading_style_without_font_is_fail(self):
        styles = make_styles(
            '<w:style w:type="paragraph" w:styleId="21">'
            '<w:pPr><w:outlineLvl w:val="1"/></w:pPr>'
            '</w:style>'
        )
        issues = []
        aud.audit_heading_style_fonts(styles, [make_p("1.1 Method", style="21")], issues)
        self.assertIn("HEADING_STYLE_FONT", codes(issues))
        self.assertEqual({i.severity for i in issues if i.code == "HEADING_STYLE_FONT"}, {"FAIL"})

    def test_heading1_style_with_heiti_not_bold_ok(self):
        styles = make_styles(
            '<w:style w:type="paragraph" w:styleId="1">'
            '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
            '<w:rPr><w:rFonts w:eastAsia="Heiti" w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:b w:val="0"/></w:rPr>'
            '</w:style>'
        )
        issues = []
        aud.audit_heading_style_fonts(styles, [make_p("1 Introduction", style="1")], issues)
        self.assertNotIn("HEADING_STYLE_FONT", codes(issues))
        self.assertNotIn("HEADING_BOLD", codes(issues))

    def test_heading1_style_bold_is_fail(self):
        styles = make_styles(
            '<w:style w:type="paragraph" w:styleId="1">'
            '<w:pPr><w:outlineLvl w:val="0"/></w:pPr>'
            '<w:rPr><w:rFonts w:eastAsia="Heiti"/><w:b/></w:rPr>'
            '</w:style>'
        )
        issues = []
        aud.audit_heading_style_fonts(styles, [make_p("1 Introduction", style="1")], issues)
        self.assertIn("HEADING_BOLD", codes(issues))


class HeadingNoStyleTests(unittest.TestCase):
    def test_heading_like_text_without_style_flagged(self):
        issues = []
        aud.audit_heading_styles([make_p("1.1 研究背景")], issues)
        self.assertIn("HEADING_NO_STYLE", codes(issues))

    def test_heading_like_text_with_style_ok(self):
        issues = []
        aud.audit_heading_styles([make_p("1.1 研究背景", style="Heading2")], issues)
        self.assertNotIn("HEADING_NO_STYLE", codes(issues))

    def test_numbered_prose_not_flagged(self):
        issues = []
        aud.audit_heading_styles([make_p("2020 年完成了大量基础工作并取得成果")], issues)
        self.assertNotIn("HEADING_NO_STYLE", codes(issues))


class TableBordersTests(unittest.TestCase):
    def _table(self, vertical_border: bool) -> ET.Element:
        borders = '<w:tcBorders><w:left w:val="single" w:sz="4"/></w:tcBorders>' if vertical_border else ""
        return make_doc(
            f'<w:tbl><w:tr><w:tc><w:tcPr>{borders}</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )

    def test_vertical_border_flagged(self):
        issues = []
        aud.audit_table_borders(self._table(True), issues)
        self.assertIn("TABLE_BORDERS", codes(issues))

    def test_three_line_table_ok(self):
        issues = []
        aud.audit_table_borders(self._table(False), issues)
        self.assertNotIn("TABLE_BORDERS", codes(issues))

    def test_row_exception_grid_borders_flagged(self):
        root = make_doc(
            '<w:tbl><w:tr><w:tblPrEx><w:tblBorders>'
            '<w:left w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/>'
            '</w:tblBorders></w:tblPrEx>'
            '<w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )
        issues = []
        aud.audit_table_borders(root, issues)
        self.assertIn("TABLE_BORDERS", codes(issues))


class TableShadingTests(unittest.TestCase):
    def _table(self, shd: str = "", tbl_style: str = "") -> ET.Element:
        style_ref = f'<w:tblPr><w:tblStyle w:val="{tbl_style}"/></w:tblPr>' if tbl_style else ""
        tcpr = f"<w:tcPr>{shd}</w:tcPr>" if shd else ""
        return make_doc(
            f'<w:tbl>{style_ref}<w:tr><w:tc>{tcpr}<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )

    def test_shaded_header_cell_flagged(self):
        issues = []
        aud.audit_table_shading(self._table('<w:shd w:val="clear" w:fill="D9E2F3"/>'), issues)
        self.assertIn("TABLE_SHADING", codes(issues))

    def test_clear_auto_shading_ok(self):
        issues = []
        aud.audit_table_shading(self._table('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'), issues)
        self.assertNotIn("TABLE_SHADING", codes(issues))

    def test_white_fill_ok(self):
        issues = []
        aud.audit_table_shading(self._table('<w:shd w:val="clear" w:fill="FFFFFF"/>'), issues)
        self.assertNotIn("TABLE_SHADING", codes(issues))

    def test_pattern_shading_flagged(self):
        issues = []
        aud.audit_table_shading(self._table('<w:shd w:val="pct10" w:fill="auto"/>'), issues)
        self.assertIn("TABLE_SHADING", codes(issues))

    def test_style_driven_first_row_shading_flagged(self):
        # The header shading lives in the table style's firstRow conditional format
        # (styles.xml), not in document.xml — the classic "blue header" failure.
        styles = ET.fromstring(
            f'<w:styles xmlns:w="{W}">'
            f'<w:style w:type="table" w:styleId="GridTable4-Accent1">'
            f'<w:tblStylePr w:type="firstRow"><w:tcPr>'
            f'<w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>'
            f'</w:style></w:styles>'
        )
        issues = []
        aud.audit_table_shading(self._table(tbl_style="GridTable4-Accent1"), issues, styles_root=styles)
        self.assertIn("TABLE_SHADING", codes(issues))

    def test_plain_style_reference_ok(self):
        styles = ET.fromstring(
            f'<w:styles xmlns:w="{W}">'
            f'<w:style w:type="table" w:styleId="TableNormal"/></w:styles>'
        )
        issues = []
        aud.audit_table_shading(self._table(tbl_style="TableNormal"), issues, styles_root=styles)
        self.assertNotIn("TABLE_SHADING", codes(issues))


class TableRulesTests(unittest.TestCase):
    def _three_line(self, top_sz: str = "12", header_sz: str = "6", top: bool = True) -> ET.Element:
        top_xml = f'<w:top w:val="single" w:sz="{top_sz}"/>' if top else ""
        return make_doc(
            f'<w:tbl><w:tblPr><w:tblBorders>{top_xml}'
            f'<w:bottom w:val="single" w:sz="{top_sz}"/>'
            f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            f'</w:tblBorders></w:tblPr>'
            f'<w:tr><w:tc><w:tcPr><w:tcBorders>'
            f'<w:bottom w:val="single" w:sz="{header_sz}"/></w:tcBorders></w:tcPr>'
            f'<w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>'
            f'<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )

    def test_proper_three_line_table_ok(self):
        issues = []
        aud.audit_table_rules(self._three_line(), issues)
        self.assertNotIn("TABLE_RULES", codes(issues))

    def test_borderless_table_flagged(self):
        issues = []
        aud.audit_table_rules(make_doc(TABLE_XML), issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_missing_top_rule_flagged(self):
        issues = []
        aud.audit_table_rules(self._three_line(top=False), issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_thick_middle_line_flagged(self):
        # The "no thin header rule, thick middle line" failure: header rule >= top rule.
        issues = []
        aud.audit_table_rules(self._three_line(top_sz="6", header_sz="12"), issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_visible_inside_h_flagged(self):
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblBorders>'
            '<w:top w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/>'
            '<w:insideH w:val="single" w:sz="4"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        )
        issues = []
        aud.audit_table_rules(root, issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_row_exception_inside_h_flagged(self):
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblBorders>'
            '<w:top w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/>'
            '<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tr><w:tblPrEx><w:tblBorders>'
            '<w:insideH w:val="single" w:sz="4"/>'
            '</w:tblBorders></w:tblPrEx><w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>'
            '</w:tbl>'
        )
        issues = []
        aud.audit_table_rules(root, issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_body_cell_bottom_borders_flagged(self):
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblBorders>'
            '<w:top w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/>'
            '<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tr><w:tc><w:tcPr><w:tcBorders>'
            '<w:bottom w:val="single" w:sz="6"/></w:tcBorders></w:tcPr>'
            '<w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:tcPr><w:tcBorders>'
            '<w:bottom w:val="single" w:sz="6"/></w:tcBorders></w:tcPr>'
            '<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc></w:tr>'
            '</w:tbl>'
        )
        issues = []
        aud.audit_table_rules(root, issues)
        self.assertIn("TABLE_RULES", codes(issues))

    def test_last_row_cell_bottom_rule_ok(self):
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblBorders>'
            '<w:top w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/>'
            '<w:insideH w:val="none"/><w:insideV w:val="none"/>'
            '</w:tblBorders></w:tblPr>'
            '<w:tr><w:tc><w:tcPr><w:tcBorders>'
            '<w:bottom w:val="single" w:sz="6"/></w:tcBorders></w:tcPr>'
            '<w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>'
            '<w:tr><w:tc><w:tcPr><w:tcBorders>'
            '<w:bottom w:val="single" w:sz="12"/></w:tcBorders></w:tcPr>'
            '<w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>'
            '</w:tbl>'
        )
        issues = []
        aud.audit_table_rules(root, issues)
        self.assertNotIn("TABLE_RULES", codes(issues))


class FormulaLayoutTableTests(unittest.TestCase):
    def _layout_table(self) -> ET.Element:
        return make_doc(
            '<w:tbl><w:tr>'
            '<w:tc><w:p/></w:tc>'
            '<w:tc><w:p><m:oMath><m:r><m:t>F</m:t></m:r></m:oMath></w:p></w:tc>'
            '<w:tc><w:p><w:r><w:t>(3-1)</w:t></w:r></w:p></w:tc>'
            '</w:tr></w:tbl>'
        )

    def test_formula_layout_table_detected(self):
        table = self._layout_table().find(".//w:tbl", aud.NS)
        self.assertTrue(aud.is_formula_layout_table(table))

    def test_formula_layout_table_not_forced_to_three_line(self):
        issues = []
        aud.audit_table_rules(self._layout_table(), issues)
        self.assertNotIn("TABLE_RULES", codes(issues))

    def test_formula_layout_table_skips_table_formula_text(self):
        issues = []
        aud.audit_table_formula_text(self._layout_table(), issues)
        self.assertNotIn("FORMULA_TEXT_TABLE", codes(issues))


class MathDuplicationTests(unittest.TestCase):
    MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def _p(self, plain: str, math: str) -> ET.Element:
        return ET.fromstring(
            f'<w:p xmlns:w="{W}" xmlns:m="{self.MATH_NS}">'
            f'<w:r><w:t xml:space="preserve">{plain}</w:t></w:r>'
            f'<m:oMath><m:r><m:t>{math}</m:t></m:r></m:oMath></w:p>'
        )

    def test_appended_math_duplicating_prose_flagged(self):
        # Append-instead-of-replace: the prose still says F=44.5 and the same
        # expression was appended as OMML at the paragraph end.
        issues = []
        aud.audit_math_duplication([self._p("控制面积F=44.5平方公里。", "F=44.5")], issues)
        self.assertIn("MATH_DUPLICATE", codes(issues))

    def test_in_place_math_ok(self):
        issues = []
        aud.audit_math_duplication([self._p("控制面积为下式所示。", "F=44.5")], issues)
        self.assertNotIn("MATH_DUPLICATE", codes(issues))

    def test_short_shared_symbol_not_flagged(self):
        # A lone symbol like "Q" may legitimately appear in both prose and math.
        issues = []
        aud.audit_math_duplication([self._p("其中Q代表流量。", "Q")], issues)
        self.assertNotIn("MATH_DUPLICATE", codes(issues))


class TableCellExclusionTests(unittest.TestCase):
    def test_numeric_table_cell_not_flagged_as_heading(self):
        body = (
            '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
            '<w:r><w:t>12.5</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:p><w:r><w:t>这是一段正文，用于让文档不为空并被识别为中文。</w:t></w:r></w:p>'
        )
        payload = run_main_json(body)
        result_codes = {issue["code"] for issue in payload["issues"]}
        self.assertNotIn("H1_CENTER", result_codes)
        self.assertNotIn("H1_GAP", result_codes)

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

    def _table_with_formula(self, sz: str) -> ET.Element:
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        return ET.fromstring(
            f'<w:document xmlns:w="{W}" xmlns:m="{math_ns}"><w:body><w:tbl><w:tr><w:tc><w:p>'
            f'<m:oMath><m:r><w:rPr><w:sz w:val="{sz}"/></w:rPr><m:t>x</m:t></m:r></m:oMath>'
            f'</w:p></w:tc></w:tr></w:tbl></w:body></w:document>'
        )

    def test_in_table_formula_oversized_flagged(self):
        issues = []
        aud.audit_tables(self._table_with_formula("28"), issues)  # 四号 formula in a table
        self.assertIn("TABLE_SIZE", codes(issues))

    def test_in_table_formula_wuhao_ok(self):
        issues = []
        aud.audit_tables(self._table_with_formula("21"), issues)  # 五号 formula
        self.assertNotIn("TABLE_SIZE", codes(issues))


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
            write_minimal_docx(path, document)
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


class FormulaItalicTests(unittest.TestCase):
    def _math(self, run_xml: str) -> ET.Element:
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        return ET.fromstring(
            f'<w:document xmlns:w="{W}" xmlns:m="{math_ns}"><w:body><w:p>'
            f'<m:oMath>{run_xml}</m:oMath></w:p></w:body></w:document>'
        )

    def test_italic_digit_flagged(self):
        issues = []
        aud.audit_formula_digit_italics(self._math('<m:r><w:rPr><w:i/></w:rPr><m:t>1</m:t></m:r>'), issues)
        self.assertIn("FORMULA_DIGIT_ITALIC", codes(issues))

    def test_upright_digit_ok(self):
        issues = []
        aud.audit_formula_digit_italics(self._math('<m:r><m:t>1</m:t></m:r>'), issues)
        self.assertNotIn("FORMULA_DIGIT_ITALIC", codes(issues))

    def test_italic_variable_letter_ok(self):
        issues = []
        aud.audit_formula_digit_italics(self._math('<m:r><w:rPr><w:i/></w:rPr><m:t>x</m:t></m:r>'), issues)
        self.assertNotIn("FORMULA_DIGIT_ITALIC", codes(issues))


class EquationNumberTests(unittest.TestCase):
    def _eq(self, jc: Optional[str]) -> ET.Element:
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        jc_xml = f'<w:jc w:val="{jc}"/>' if jc else ""
        return ET.fromstring(
            f'<w:p xmlns:w="{W}" xmlns:m="{math_ns}"><w:pPr>{jc_xml}</w:pPr>'
            f'<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:r><w:t>(式 3-3)</w:t></w:r></w:p>'
        )

    def test_centered_numbered_formula_flagged(self):
        issues = []
        aud.audit_equation_numbers([self._eq("center")], issues)
        self.assertIn("EQUATION_NUMBER_CENTER", codes(issues))

    def test_left_numbered_formula_ok(self):
        issues = []
        aud.audit_equation_numbers([self._eq("left")], issues)
        self.assertNotIn("EQUATION_NUMBER_CENTER", codes(issues))


class FormulaMixedItalicTests(unittest.TestCase):
    MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def _math(self, run_xml: str) -> ET.Element:
        return ET.fromstring(
            f'<w:document xmlns:w="{W}" xmlns:m="{self.MATH_NS}"><w:body><w:p>'
            f'<m:oMath>{run_xml}</m:oMath></w:p></w:body></w:document>'
        )

    def test_mixed_italic_run_with_digits_flagged(self):
        # A single italic run 'F=44.5' slants its digits even though it contains a letter.
        issues = []
        aud.audit_formula_digit_italics(
            self._math('<m:r><w:rPr><w:i/></w:rPr><m:t>F=44.5</m:t></m:r>'), issues
        )
        self.assertIn("FORMULA_DIGIT_ITALIC", codes(issues))

    def test_omml_sty_mval_italic_digits_flagged(self):
        # OMML style italics use m:sty m:val="i" (the m namespace), not w:val.
        issues = []
        aud.audit_formula_digit_italics(
            self._math(
                '<m:r><m:rPr><m:sty m:val="i"/></m:rPr><m:t>44.5</m:t></m:r>'
            ),
            issues,
        )
        self.assertIn("FORMULA_DIGIT_ITALIC", codes(issues))

    def test_multiletter_default_italic_flagged(self):
        # 'km' with no upright style renders math-italic by default — no w:i needed.
        issues = []
        aud.audit_formula_multiletter_italics(self._math('<m:r><m:t>km</m:t></m:r>'), issues)
        self.assertIn("FORMULA_MULTILETTER_ITALIC", codes(issues))

    def test_multiletter_upright_sty_ok(self):
        issues = []
        aud.audit_formula_multiletter_italics(
            self._math('<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>km</m:t></m:r>'), issues
        )
        self.assertNotIn("FORMULA_MULTILETTER_ITALIC", codes(issues))

    def test_multiletter_nor_ok(self):
        issues = []
        aud.audit_formula_multiletter_italics(
            self._math('<m:r><m:rPr><m:nor/></m:rPr><m:t>max</m:t></m:r>'), issues
        )
        self.assertNotIn("FORMULA_MULTILETTER_ITALIC", codes(issues))

    def test_single_letter_variable_ok(self):
        issues = []
        aud.audit_formula_multiletter_italics(self._math('<m:r><m:t>Q</m:t></m:r>'), issues)
        self.assertNotIn("FORMULA_MULTILETTER_ITALIC", codes(issues))


class ManualItalicMathTests(unittest.TestCase):
    def _p(self, text: str, italic: bool) -> ET.Element:
        rpr = "<w:rPr><w:i/></w:rPr>" if italic else ""
        return ET.fromstring(
            f'<w:p xmlns:w="{W}"><w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        )

    def test_italic_equation_run_flagged(self):
        issues = []
        aud.audit_manual_italic_math([self._p("F = 44.5 km²", True)], issues)
        self.assertIn("MANUAL_ITALIC_MATH", codes(issues))

    def test_upright_equation_run_not_flagged_here(self):
        # Plain-text math without italics is FORMULA_TEXT's job, not this check's.
        issues = []
        aud.audit_manual_italic_math([self._p("F = 44.5 km²", False)], issues)
        self.assertNotIn("MANUAL_ITALIC_MATH", codes(issues))

    def test_italic_prose_without_math_ok(self):
        issues = []
        aud.audit_manual_italic_math([self._p("Journal of Hydrology", True)], issues)
        self.assertNotIn("MANUAL_ITALIC_MATH", codes(issues))


class CaptionAlignmentTests(unittest.TestCase):
    def test_left_aligned_caption_flagged(self):
        issues = []
        aud.audit_caption_alignment([make_p("表 1-1 方案比较", jc="left")], issues)
        self.assertIn("CAPTION_ALIGN", codes(issues))

    def test_centered_caption_ok(self):
        issues = []
        aud.audit_caption_alignment([make_p("图 2-3 流量过程线", jc="center")], issues)
        self.assertNotIn("CAPTION_ALIGN", codes(issues))

    def test_styled_caption_without_direct_jc_ok(self):
        # A named style may center the caption; only direct formatting is judged.
        issues = []
        aud.audit_caption_alignment([make_p("表 1-1 方案比较", style="Caption")], issues)
        self.assertNotIn("CAPTION_ALIGN", codes(issues))


class NumberUnitSpacingTests(unittest.TestCase):
    def test_glued_unit_flagged(self):
        issues = []
        aud.audit_number_unit_spacing([make_p("最大流量为20km处的断面控制。")], issues)
        self.assertIn("NUMBER_UNIT_SPACING", codes(issues))

    def test_glued_cubic_metre_flagged(self):
        issues = []
        aud.audit_number_unit_spacing([make_p("设计流量为216m³/s。")], issues)
        self.assertIn("NUMBER_UNIT_SPACING", codes(issues))

    def test_spaced_unit_ok(self):
        issues = []
        aud.audit_number_unit_spacing([make_p("设计流量为 216 m³/s，距离为 20 km。")], issues)
        self.assertNotIn("NUMBER_UNIT_SPACING", codes(issues))

    def test_space_before_percent_flagged(self):
        issues = []
        aud.audit_number_unit_spacing([make_p("设计频率为 0.1 %的洪水。")], issues)
        self.assertIn("NUMBER_UNIT_SPACING", codes(issues))

    def test_attached_percent_ok(self):
        issues = []
        aud.audit_number_unit_spacing([make_p("设计频率为0.1%的洪水，水温 25℃。")], issues)
        self.assertNotIn("NUMBER_UNIT_SPACING", codes(issues))


class FloatOrderTests(unittest.TestCase):
    def test_caption_before_first_reference_flagged(self):
        issues = []
        paragraphs = [
            make_p("表 1-1 方案比较"),
            make_p("正文在表格之后才提到，如表 1-1 所示。"),
        ]
        aud.audit_float_order(paragraphs, issues)
        self.assertIn("FLOAT_ORDER", codes(issues))

    def test_reference_before_caption_ok(self):
        issues = []
        paragraphs = [
            make_p("各方案指标如表 1-1 所示。"),
            make_p("表 1-1 方案比较"),
        ]
        aud.audit_float_order(paragraphs, issues)
        self.assertNotIn("FLOAT_ORDER", codes(issues))

    def test_unreferenced_caption_not_flagged(self):
        # No mention anywhere: left to human review, not flagged as inverted order.
        issues = []
        aud.audit_float_order([make_p("表 1-1 方案比较")], issues)
        self.assertNotIn("FLOAT_ORDER", codes(issues))


class TableHeaderRepeatTests(unittest.TestCase):
    def _table(self, header: bool, rows: int = 2) -> ET.Element:
        trpr = "<w:trPr><w:tblHeader/></w:trPr>" if header else ""
        first = f"<w:tr>{trpr}<w:tc><w:p><w:r><w:t>h</w:t></w:r></w:p></w:tc></w:tr>"
        body = "<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr>" * (rows - 1)
        return make_doc(f"<w:tbl>{first}{body}</w:tbl>")

    def test_multirow_without_repeat_flagged(self):
        issues = []
        aud.audit_table_header_repeat(self._table(header=False), issues)
        self.assertIn("TABLE_HEADER_REPEAT", codes(issues))

    def test_multirow_with_repeat_ok(self):
        issues = []
        aud.audit_table_header_repeat(self._table(header=True), issues)
        self.assertNotIn("TABLE_HEADER_REPEAT", codes(issues))

    def test_single_row_table_ok(self):
        issues = []
        aud.audit_table_header_repeat(self._table(header=False, rows=1), issues)
        self.assertNotIn("TABLE_HEADER_REPEAT", codes(issues))


class TableFormulaTextTests(unittest.TestCase):
    def _table_with_cell(self, cell_text: str) -> ET.Element:
        return make_doc(
            f'<w:tbl><w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">{cell_text}</w:t></w:r></w:p>'
            f"</w:tc></w:tr></w:tbl>"
        )

    def test_equation_in_cell_flagged(self):
        issues = []
        aud.audit_table_formula_text(self._table_with_cell("Q_p=W/T"), issues)
        self.assertIn("FORMULA_TEXT_TABLE", codes(issues))

    def test_numeric_cell_not_flagged(self):
        issues = []
        aud.audit_table_formula_text(self._table_with_cell("216.5"), issues)
        self.assertNotIn("FORMULA_TEXT_TABLE", codes(issues))

    def test_plain_label_cell_not_flagged(self):
        issues = []
        aud.audit_table_formula_text(self._table_with_cell("方案比较"), issues)
        self.assertNotIn("FORMULA_TEXT_TABLE", codes(issues))


class EquationNumberTabsTests(unittest.TestCase):
    MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def _eq(self, tabs: bool) -> ET.Element:
        tabs_xml = (
            '<w:tabs><w:tab w:val="center" w:pos="4536"/><w:tab w:val="right" w:pos="9072"/></w:tabs>'
            if tabs
            else ""
        )
        return ET.fromstring(
            f'<w:p xmlns:w="{W}" xmlns:m="{self.MATH_NS}"><w:pPr>{tabs_xml}</w:pPr>'
            f"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:r><w:t>(3-3)</w:t></w:r></w:p>"
        )

    def test_numbered_equation_without_right_tab_flagged(self):
        issues = []
        aud.audit_equation_number_tabs([self._eq(tabs=False)], issues)
        self.assertIn("EQUATION_NUMBER_TABS", codes(issues))

    def test_numbered_equation_with_right_tab_ok(self):
        issues = []
        aud.audit_equation_number_tabs([self._eq(tabs=True)], issues)
        self.assertNotIn("EQUATION_NUMBER_TABS", codes(issues))


class FieldsUpdateTests(unittest.TestCase):
    def test_fields_without_updatefields_flagged(self):
        issues = []
        aud.audit_fields_update('<w:instrText> REF ref_001 \\h </w:instrText>', None, issues)
        self.assertIn("FIELDS_UPDATE", codes(issues))

    def test_fields_with_updatefields_ok(self):
        issues = []
        settings = '<w:settings><w:updateFields w:val="true"/></w:settings>'
        aud.audit_fields_update('<w:instrText> REF ref_001 \\h </w:instrText>', settings, issues)
        self.assertNotIn("FIELDS_UPDATE", codes(issues))

    def test_document_without_fields_ok(self):
        issues = []
        aud.audit_fields_update("<w:document><w:body/></w:document>", None, issues)
        self.assertNotIn("FIELDS_UPDATE", codes(issues))


class FirstlineIndentUnitTests(unittest.TestCase):
    def _p(self, ind: str) -> ET.Element:
        return ET.fromstring(
            f'<w:p xmlns:w="{W}"><w:pPr><w:ind {ind}/></w:pPr>'
            f"<w:r><w:t>正文段落。</w:t></w:r></w:p>"
        )

    def test_fixed_twip_indent_flagged(self):
        issues = []
        aud.audit_firstline_indent_unit([self._p('w:firstLine="480"')], issues)
        self.assertIn("FIRSTLINE_FIXED", codes(issues))

    def test_char_based_indent_ok(self):
        issues = []
        aud.audit_firstline_indent_unit(
            [self._p('w:firstLine="480" w:firstLineChars="200"')], issues
        )
        self.assertNotIn("FIRSTLINE_FIXED", codes(issues))


class PackageIntegrityTests(unittest.TestCase):
    def test_missing_required_parts_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "bare.docx")
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", "<w:document/>")
            issues = []
            aud.audit_package_integrity(pathlib.Path(path), issues)
        self.assertIn("PACKAGE_INTEGRITY", codes(issues))

    def test_complete_package_ok(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "ok.docx")
            write_minimal_docx(path, f'<w:document xmlns:w="{W}"><w:body/></w:document>')
            issues = []
            aud.audit_package_integrity(pathlib.Path(path), issues)
        self.assertNotIn("PACKAGE_INTEGRITY", codes(issues))


class TableStyleGridTests(unittest.TestCase):
    def test_style_driven_grid_borders_flagged_as_fail(self):
        # Grid borders live in styles.xml (e.g. Table Grid); document.xml shows
        # only the style reference — exactly how "所有表格都不是三线表" slipped through.
        styles = ET.fromstring(
            f'<w:styles xmlns:w="{W}">'
            f'<w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders>'
            f'<w:insideV w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>'
            f"</w:tblBorders></w:tblPr></w:style></w:styles>"
        )
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>'
            "<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        issues = []
        aud.audit_table_borders(root, issues, styles_root=styles)
        self.assertIn("TABLE_BORDERS", codes(issues))
        self.assertEqual([i.severity for i in issues if i.code == "TABLE_BORDERS"], ["FAIL"])

    def test_plain_style_reference_not_flagged(self):
        styles = ET.fromstring(
            f'<w:styles xmlns:w="{W}"><w:style w:type="table" w:styleId="TableNormal"/></w:styles>'
        )
        root = make_doc(
            '<w:tbl><w:tblPr><w:tblStyle w:val="TableNormal"/></w:tblPr>'
            "<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        issues = []
        aud.audit_table_borders(root, issues, styles_root=styles)
        self.assertNotIn("TABLE_BORDERS", codes(issues))


class MustFixSeverityTests(unittest.TestCase):
    def test_table_rules_and_digit_italic_are_fail_level(self):
        # Must-Fix items must block the gate: WARN does not stop delivery.
        issues = []
        aud.audit_table_rules(make_doc(TABLE_XML), issues)
        self.assertEqual({i.severity for i in issues if i.code == "TABLE_RULES"}, {"FAIL"})
        math_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        root = ET.fromstring(
            f'<w:document xmlns:w="{W}" xmlns:m="{math_ns}"><w:body><w:p><m:oMath>'
            f'<m:r><w:rPr><w:i/></w:rPr><m:t>44.5</m:t></m:r></m:oMath></w:p></w:body></w:document>'
        )
        issues = []
        aud.audit_formula_digit_italics(root, issues)
        self.assertEqual({i.severity for i in issues if i.code == "FORMULA_DIGIT_ITALIC"}, {"FAIL"})


class EquationNumberingTests(unittest.TestCase):
    MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

    def _eq_p(self, trailing_text: str) -> ET.Element:
        run = f'<w:r><w:t xml:space="preserve">{trailing_text}</w:t></w:r>' if trailing_text else ""
        return ET.fromstring(
            f'<w:p xmlns:w="{W}" xmlns:m="{self.MATH_NS}">'
            f"<m:oMath><m:r><m:t>W=0.278KFP</m:t></m:r></m:oMath>{run}</w:p>"
        )

    def test_unnumbered_display_equation_flagged(self):
        issues = []
        aud.audit_equation_numbering([self._eq_p("")], issues)
        self.assertIn("EQUATION_UNNUMBERED", codes(issues))
        self.assertEqual([i.severity for i in issues if i.code == "EQUATION_UNNUMBERED"], ["FAIL"])

    def test_numbered_display_equation_ok(self):
        issues = []
        aud.audit_equation_numbering([self._eq_p("(3-1)")], issues)
        self.assertNotIn("EQUATION_UNNUMBERED", codes(issues))

    def test_inline_equation_in_prose_not_flagged(self):
        issues = []
        aud.audit_equation_numbering([self._eq_p("为设计洪峰流量计算式。")], issues)
        self.assertNotIn("EQUATION_UNNUMBERED", codes(issues))

    def test_numbered_but_never_referenced_warned(self):
        issues = []
        paragraphs = [self._eq_p("(3-1)"), make_p("后续正文没有提到该公式。")]
        aud.audit_equation_references(paragraphs, issues)
        self.assertIn("EQUATION_NOT_REFERENCED", codes(issues))

    def test_referenced_equation_ok(self):
        issues = []
        paragraphs = [self._eq_p("(3-1)"), make_p("由式 (3-1) 可得设计洪峰流量。")]
        aud.audit_equation_references(paragraphs, issues)
        self.assertNotIn("EQUATION_NOT_REFERENCED", codes(issues))


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
