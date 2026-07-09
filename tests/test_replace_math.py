"""Unit tests for scripts/replace_math.py (standard library only).

The splice logic is tested with pre-converted OMML fragments so no Pandoc is
needed; one round-trip test exercises the Pandoc path and is skipped when
Pandoc is not installed.
"""

from __future__ import annotations

import importlib.util
import pathlib
import shutil
import sys
import unittest
from xml.etree import ElementTree as ET


def _load_module():
    path = pathlib.Path(__file__).resolve().parent.parent / "scripts" / "replace_math.py"
    spec = importlib.util.spec_from_file_location("replace_math", path)
    module = importlib.util.module_from_spec(spec)
    sys.modules["replace_math"] = module
    spec.loader.exec_module(module)
    return module


rm = _load_module()

W = rm.W_NS
M = rm.M_NS
OMML = f'<m:oMath xmlns:m="{M}" xmlns:w="{W}"><m:r><m:t>F=44.5</m:t></m:r></m:oMath>'


def make_document(body: str) -> str:
    return (
        f'<w:document xmlns:w="{W}" xmlns:m="{M}"><w:body>{body}</w:body></w:document>'
    )


def apply(document_xml: str, registry, sz: int = 24):
    return rm.apply_registry(document_xml, registry, sz)


def math_texts(xml: str) -> list:
    root = ET.fromstring(xml)
    ns = {"m": M}
    return [
        "".join(t.text or "" for t in math.findall(".//m:t", ns))
        for math in root.findall(".//m:oMath", ns)
    ]


class InlineReplaceTests(unittest.TestCase):
    def test_single_run_replacement_preserves_surroundings(self):
        doc = make_document(
            "<w:p><w:r><w:t>控制面积F=44.5平方公里，属小流域。</w:t></w:r></w:p>"
        )
        xml, report = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        self.assertEqual(report["replaced"]["F=44.5"], 1)
        self.assertEqual(report["not_found"], [])
        self.assertEqual(report["still_plain_text"], [])
        self.assertEqual(math_texts(xml), ["F=44.5"])
        root = ET.fromstring(xml)
        plain = "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
        self.assertEqual(plain, "控制面积平方公里，属小流域。")

    def test_cross_run_replacement(self):
        # The token is split across three runs, as real Word documents do.
        doc = make_document(
            "<w:p><w:r><w:t>面积F=</w:t></w:r><w:r><w:t>44</w:t></w:r>"
            "<w:r><w:t>.5平方公里</w:t></w:r></w:p>"
        )
        xml, report = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        self.assertEqual(report["replaced"]["F=44.5"], 1)
        root = ET.fromstring(xml)
        plain = "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
        self.assertEqual(plain, "面积平方公里")
        self.assertEqual(math_texts(xml), ["F=44.5"])

    def test_multiple_occurrences_all_replaced(self):
        doc = make_document(
            "<w:p><w:r><w:t>取F=44.5，复核F=44.5。</w:t></w:r></w:p>"
        )
        xml, report = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        self.assertEqual(report["replaced"]["F=44.5"], 2)
        self.assertEqual(len(math_texts(xml)), 2)

    def test_missing_token_reported(self):
        doc = make_document("<w:p><w:r><w:t>没有公式。</w:t></w:r></w:p>")
        _, report = apply(doc, [{"find": "Q=1", "omml": OMML}])
        self.assertEqual(report["not_found"], ["Q=1"])

    def test_font_size_stamped(self):
        doc = make_document("<w:p><w:r><w:t>F=44.5</w:t></w:r></w:p>")
        xml, _ = apply(doc, [{"find": "F=44.5", "omml": OMML, "sz": 21}])
        root = ET.fromstring(xml)
        sizes = [n.get(f"{{{W}}}val") for n in root.iter(f"{{{W}}}sz")]
        self.assertEqual(sizes, ["21"])

    def test_body_formula_default_size_is_xiaosi(self):
        doc = make_document("<w:p><w:r><w:t>F=44.5</w:t></w:r></w:p>")
        xml, _ = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        root = ET.fromstring(xml)
        sizes = [n.get(f"{{{W}}}val") for n in root.iter(f"{{{W}}}sz")]
        self.assertEqual(sizes, ["24"])

    def test_table_formula_default_size_is_wuhao(self):
        doc = make_document(
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>F=44.5</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        xml, _ = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        root = ET.fromstring(xml)
        sizes = [n.get(f"{{{W}}}val") for n in root.iter(f"{{{W}}}sz")]
        self.assertEqual(sizes, ["21"])

    def test_run_properties_preserved_on_split(self):
        doc = make_document(
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>前F=44.5后</w:t></w:r></w:p>'
        )
        xml, _ = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        root = ET.fromstring(xml)
        runs = root.findall(f".//{{{W}}}r")
        text_runs = [r for r in runs if r.find(f"{{{W}}}t") is not None]
        self.assertEqual(len(text_runs), 2)
        for run in text_runs:
            self.assertIsNotNone(run.find(f"{{{W}}}rPr/{{{W}}}b"))


class DisplayEquationTests(unittest.TestCase):
    def test_display_paragraph_gets_tabs_and_number(self):
        doc = make_document("<w:p><w:r><w:t>F=44.5</w:t></w:r></w:p>")
        xml, report = apply(
            doc, [{"find": "F=44.5", "omml": OMML, "display": True, "number": "(3-1)"}]
        )
        self.assertEqual(report["replaced"]["F=44.5"], 1)
        root = ET.fromstring(xml)
        paragraph = root.find(f".//{{{W}}}p")
        tabs = paragraph.findall(f"{{{W}}}pPr/{{{W}}}tabs/{{{W}}}tab")
        self.assertEqual([t.get(f"{{{W}}}val") for t in tabs], ["center", "right"])
        jc = paragraph.find(f"{{{W}}}pPr/{{{W}}}jc")
        self.assertEqual(jc.get(f"{{{W}}}val"), "left")
        plain = "".join(t.text or "" for t in paragraph.iter(f"{{{W}}}t"))
        self.assertEqual(plain, "(3-1)")
        self.assertEqual(math_texts(xml), ["F=44.5"])


class NamespacePreservationTests(unittest.TestCase):
    def test_root_tag_kept_verbatim(self):
        # mc:Ignorable must survive even though ET would drop the unused xmlns.
        doc = (
            f'<w:document xmlns:w="{W}" xmlns:m="{M}" '
            f'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            f'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
            f'mc:Ignorable="w14"><w:body>'
            f"<w:p><w:r><w:t>F=44.5</w:t></w:r></w:p></w:body></w:document>"
        )
        xml, _ = apply(doc, [{"find": "F=44.5", "omml": OMML}])
        self.assertIn('mc:Ignorable="w14"', xml)
        self.assertIn('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', xml)


@unittest.skipUnless(rm.find_pandoc(), "Pandoc not installed")
class PandocRoundTripTests(unittest.TestCase):
    def test_latex_converts_to_omath(self):
        (math,) = rm.latex_to_omml([("F = 44.5\\,\\mathrm{km^2}", False)])
        text = "".join(t.text or "" for t in math.iter(f"{{{M}}}t"))
        self.assertIn("44.5", text)
        self.assertIn("km", text)
        # \mathrm must arrive as an explicit upright style (m:sty val="p").
        styles = [s.get(f"{{{M}}}val") for s in math.iter(f"{{{M}}}sty")]
        self.assertIn("p", styles)

    def test_registry_with_latex_end_to_end(self):
        doc = make_document("<w:p><w:r><w:t>面积F=44.5 km²属小流域。</w:t></w:r></w:p>")
        xml, report = apply(
            doc, [{"find": "F=44.5 km²", "latex": "F = 44.5\\,\\mathrm{km^2}"}]
        )
        self.assertEqual(report["replaced"]["F=44.5 km²"], 1)
        self.assertEqual(report["still_plain_text"], [])
        root = ET.fromstring(xml)
        plain = "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
        self.assertEqual(plain, "面积属小流域。")


if __name__ == "__main__":
    unittest.main()
