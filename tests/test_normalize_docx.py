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


if __name__ == "__main__":
    unittest.main()
