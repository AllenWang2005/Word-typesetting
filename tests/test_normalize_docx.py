"""Unit tests for scripts/normalize_docx.py (standard library only)."""

from __future__ import annotations

import importlib.util
import pathlib
import sys
import unittest


def _load_normalize_module():
    path = pathlib.Path(__file__).resolve().parent.parent / "scripts" / "normalize_docx.py"
    spec = importlib.util.spec_from_file_location("normalize_docx", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules["normalize_docx"] = module
    spec.loader.exec_module(module)
    return module


norm = _load_normalize_module()


class NormalizeTextTests(unittest.TestCase):
    def test_fullwidth_citation_brackets_become_ascii(self):
        out, counts = norm.normalize_document_xml("见文献［1］与【2,3】的结论")
        self.assertIn("[1]", out)
        self.assertIn("[2,3]", out)
        self.assertEqual(counts["citation_brackets"], 2)

    def test_ascii_punct_between_cjk_becomes_fullwidth(self):
        out, counts = norm.normalize_document_xml("这是中文,后面继续")
        self.assertIn("中文，后面", out)
        self.assertEqual(counts["cjk_punctuation"], 1)

    def test_decimals_are_not_touched(self):
        out, counts = norm.normalize_document_xml("结果为3.14倍")
        self.assertIn("3.14", out)
        self.assertEqual(counts["cjk_punctuation"], 0)

    def test_ascii_citation_not_touched(self):
        out, _ = norm.normalize_document_xml("见文献[1]的结论")
        self.assertIn("[1]", out)

    def test_url_comma_not_touched(self):
        # The comma sits between an ASCII letter and a CJK char, so it is left alone.
        out, counts = norm.normalize_document_xml("详见 http://example.com,继续阅读")
        self.assertIn("http://example.com,继续", out)
        self.assertEqual(counts["cjk_punctuation"], 0)


class UnitSpacingFixTests(unittest.TestCase):
    def test_glued_unit_gets_space(self):
        xml = "<w:t>距离为20km，流量216m³/s。</w:t>"
        out, fixed = norm.fix_unit_spacing(xml)
        self.assertIn("20 km", out)
        self.assertIn("216 m³/s", out)
        self.assertEqual(fixed, 2)

    def test_space_before_percent_removed(self):
        out, fixed = norm.fix_unit_spacing("<w:t>频率为 0.1 %，水温 25 ℃。</w:t>")
        self.assertIn("0.1%", out)
        self.assertIn("25℃", out)
        self.assertEqual(fixed, 2)

    def test_text_outside_wt_untouched(self):
        # Attribute values and tags must never be rewritten.
        xml = '<w:pict o:title="20km"/><w:t>正文20km</w:t>'
        out, _ = norm.fix_unit_spacing(xml)
        self.assertIn('o:title="20km"', out)
        self.assertIn("正文20 km", out)

    def test_already_spaced_untouched(self):
        out, fixed = norm.fix_unit_spacing("<w:t>距离为 20 km。</w:t>")
        self.assertEqual(fixed, 0)


class TableHygieneFixTests(unittest.TestCase):
    def test_shading_cleared_and_tbllook_zeroed(self):
        xml = (
            '<w:tbl><w:tblPr><w:tblLook w:val="04A0" w:firstRow="1"/></w:tblPr>'
            '<w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr>'
            "<w:p/></w:tc></w:tr></w:tbl>"
        )
        out, counts = norm.fix_table_hygiene(xml)
        self.assertNotIn("D9E2F3", out)
        self.assertIn('w:fill="auto"', out)
        self.assertIn('w:firstRow="0"', out)
        self.assertEqual(counts["shading_cleared"], 1)
        self.assertEqual(counts["tbllook_zeroed"], 1)

    def test_header_repeat_added_to_multirow_table(self):
        xml = "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
        out, counts = norm.fix_table_hygiene(xml)
        self.assertEqual(counts["header_repeat_added"], 1)
        first_row_end = out.find("</w:tr>")
        self.assertIn("<w:tblHeader/>", out[:first_row_end])

    def test_row_exception_borders_cleared(self):
        xml = (
            '<w:tbl><w:tr><w:tblPrEx><w:tblBorders>'
            '<w:left w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/>'
            '</w:tblBorders><w:tblCellMar><w:left w:w="108"/></w:tblCellMar></w:tblPrEx>'
            '<w:tc><w:p/></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'
        )
        out, counts = norm.fix_table_hygiene(xml)
        self.assertNotIn("<w:tblBorders>", out)
        self.assertIn("<w:tblCellMar>", out)
        self.assertEqual(counts["row_exception_borders_cleared"], 1)

    def test_single_row_table_untouched(self):
        xml = "<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
        out, counts = norm.fix_table_hygiene(xml)
        self.assertNotIn("tblHeader", out)
        self.assertEqual(counts["header_repeat_added"], 0)

    def test_shading_outside_tables_untouched(self):
        xml = '<w:p><w:pPr><w:shd w:val="clear" w:fill="FFFF00"/></w:pPr></w:p>'
        out, counts = norm.fix_table_hygiene(xml)
        self.assertIn("FFFF00", out)
        self.assertEqual(counts["shading_cleared"], 0)

    def test_in_table_formula_size_set_to_wuhao(self):
        xml = (
            '<w:tbl><w:tr><w:tc><w:p><m:oMath>'
            '<m:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><m:t>Q</m:t></m:r>'
            '</m:oMath></w:p></w:tc></w:tr></w:tbl>'
        )
        out, counts = norm.fix_table_hygiene(xml)
        self.assertIn('<w:sz w:val="21"/>', out)
        self.assertIn('<w:szCs w:val="21"/>', out)
        self.assertNotIn('w:val="24"', out)
        self.assertEqual(counts["table_formula_size_fixed"], 1)

    def test_in_table_formula_missing_size_gets_wuhao(self):
        xml = (
            '<w:tbl><w:tr><w:tc><w:p><m:oMath>'
            '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>Q</m:t></m:r>'
            '</m:oMath></w:p></w:tc></w:tr></w:tbl>'
        )
        out, counts = norm.fix_table_hygiene(xml)
        self.assertIn('<w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>', out)
        self.assertEqual(counts["table_formula_size_fixed"], 1)

    def test_formula_layout_table_keeps_body_formula_size(self):
        xml = (
            '<w:tbl><w:tr>'
            '<w:tc><w:p/></w:tc>'
            '<w:tc><w:p><m:oMath>'
            '<m:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><m:t>F</m:t></m:r>'
            '</m:oMath></w:p></w:tc>'
            '<w:tc><w:p><w:r><w:t>(3-1)</w:t></w:r></w:p></w:tc>'
            '</w:tr></w:tbl>'
        )
        out, counts = norm.fix_table_hygiene(xml)
        self.assertIn('w:val="24"', out)
        self.assertEqual(counts["table_formula_size_fixed"], 0)


class UpdateFieldsFixTests(unittest.TestCase):
    def test_updatefields_inserted(self):
        settings = '<w:settings xmlns:w="ns"><w:zoom/></w:settings>'
        out, changed = norm.fix_settings_update_fields(settings)
        self.assertTrue(changed)
        self.assertIn('<w:updateFields w:val="true"/>', out)
        self.assertLess(out.find("updateFields"), out.find("w:zoom"))

    def test_existing_updatefields_untouched(self):
        settings = '<w:settings><w:updateFields w:val="true"/></w:settings>'
        out, changed = norm.fix_settings_update_fields(settings)
        self.assertFalse(changed)
        self.assertEqual(out, settings)


if __name__ == "__main__":
    unittest.main()
