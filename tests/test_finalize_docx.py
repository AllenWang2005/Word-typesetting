"""Smoke tests for scripts/finalize_docx.py (standard library only)."""

from __future__ import annotations

import os
import pathlib
import subprocess
import sys
import tempfile
import unittest
import zipfile

SCRIPT = pathlib.Path(__file__).resolve().parent.parent / "scripts" / "finalize_docx.py"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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


def write_docx(path: str, body: str) -> None:
    document = f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        archive.writestr("word/document.xml", document)


def run_finalize(path: str, *extra: str) -> "subprocess.CompletedProcess[str]":
    return subprocess.run(
        [sys.executable, str(SCRIPT), path, *extra],
        capture_output=True,
        text=True,
    )


class FinalizeGateTests(unittest.TestCase):
    def test_clean_document_passes_gate(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "clean.docx")
            write_docx(
                path,
                "<w:p><w:r><w:t>这是一段完全使用中文标点的正常正文内容，用于验证交付门禁在"
                "没有任何问题时给出放行结论，同时保持文档语言判定为中文。</w:t></w:r></w:p>",
            )
            result = run_finalize(path)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("DELIVERY GATE: PASS", result.stdout)

    def test_failing_document_blocks_gate(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "bad.docx")
            # Trailing ASCII period in Chinese prose survives the mechanical fixes
            # (only punctuation *between* CJK chars is auto-fixed), so ZH_PUNCT FAILs.
            write_docx(
                path,
                "<w:p><w:r><w:t>这一段中文正文以英文句号结尾因此必须被审计器拦下来这就是门禁"
                "存在的意义也是本条测试的目的.</w:t></w:r></w:p>",
            )
            result = run_finalize(path)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("DELIVERY GATE: FAIL", result.stdout)

    def test_no_fix_leaves_file_untouched(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "doc.docx")
            write_docx(path, "<w:p><w:r><w:t>这是一段足够长的中文正文,用来保证语言检测把文档判定为中文而不是英文,否则中文标点反而会被当成英文文档里的违规标点而误报。</w:t></w:r></w:p>")
            before = pathlib.Path(path).read_bytes()
            run_finalize(path, "--no-fix")
            self.assertEqual(pathlib.Path(path).read_bytes(), before)

    def test_fix_mode_repairs_punctuation(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "doc.docx")
            write_docx(path, "<w:p><w:r><w:t>这是一段足够长的中文正文,用来保证语言检测把文档判定为中文而不是英文,否则中文标点反而会被当成英文文档里的违规标点而误报。</w:t></w:r></w:p>")
            result = run_finalize(path)
            with zipfile.ZipFile(path) as archive:
                document = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("正文，用来", document)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
