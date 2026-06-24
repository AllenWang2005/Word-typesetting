#!/usr/bin/env python3
"""Lightweight DOCX audit for Allen's Word report formatting skill."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
LATIN_RE = re.compile(r"[A-Za-z]")
CHINESE_H1_RE = re.compile(r"^[一二三四五六七八九十]+[、.]")
CHINESE_PUNCT_RE = re.compile(r"[，。；：？！、“”‘’（）【】［］《》]")
CITATION_RE = re.compile(r"[\[［【]\s*\d+(?:\s*(?:,|，|-|–|—)\s*\d+)*\s*[\]］】]")
FULLWIDTH_CITATION_RE = re.compile(r"[［【]\s*\d+(?:\s*(?:,|，|-|–|—)\s*\d+)*\s*[］】]")
CN_BIB_HEADING_RE = re.compile(r"^(?:第?[一二三四五六七八九十0-9]+(?:章|节)?[、.．]?)?(参考文献|参考资料)$")
EN_BIB_HEADING_RE = re.compile(r"^(?:\d+[.)]?)?references$", re.I)

PROTECTED_PATTERNS = [
    re.compile(r"https?://\S+", re.I),
    re.compile(r"\bdoi\s*:\s*\S+", re.I),
    re.compile(r"\b10\.\d{4,9}/\S+", re.I),
    re.compile(r"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b"),
    re.compile(r"\S+\.(?:docx?|xlsx?|pptx?|pdf|csv|txt|dat|py|r|m)\b", re.I),
    re.compile(r"\[[0-9,\-\s]+\]"),
    re.compile(r"\b\d+\.\d+\b"),
    re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b"),
    re.compile(r"\bref_\d+\b", re.I),
    re.compile(r"`[^`]+`"),
]


@dataclass
class Issue:
    severity: str
    code: str
    message: str


def w_attr(element: Optional[ET.Element], name: str) -> Optional[str]:
    if element is None:
        return None
    return element.get(f"{W}{name}")


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def strip_protected(text: str) -> str:
    cleaned = text
    for pattern in PROTECTED_PATTERNS:
        cleaned = pattern.sub(" ", cleaned)
    return cleaned


def add_issue(issues: list[Issue], severity: str, code: str, message: str, limit: int = 8) -> None:
    if sum(1 for issue in issues if issue.code == code) < limit:
        issues.append(Issue(severity, code, message))


def detect_language(text: str) -> tuple[str, int, int]:
    cjk_count = len(CJK_RE.findall(text))
    latin_count = len(LATIN_RE.findall(text))
    language = "zh" if cjk_count >= max(40, latin_count * 0.25) else "en"
    return language, cjk_count, latin_count


def paragraph_alignment(paragraph: ET.Element) -> str | None:
    return w_attr(paragraph.find("w:pPr/w:jc", NS), "val")


def has_nonzero_indent(paragraph: ET.Element) -> bool:
    ind = paragraph.find("w:pPr/w:ind", NS)
    if ind is None:
        return False
    for key in ("left", "firstLine", "hanging", "start"):
        value = w_attr(ind, key)
        if value not in (None, "0"):
            return True
    return False


def has_spacing_before(paragraph: ET.Element) -> bool:
    spacing = paragraph.find("w:pPr/w:spacing", NS)
    if spacing is None:
        return False
    for key in ("before", "beforeLines"):
        value = w_attr(spacing, key)
        if value not in (None, "0"):
            return True
    return False


def audit_punctuation(paragraphs: list[ET.Element], language: str, issues: list[Issue]) -> None:
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph)
        if not text.strip():
            continue
        cleaned = strip_protected(text)
        sample = text.strip().replace("\n", " ")[:90]
        if language == "zh":
            ascii_between_cjk = re.search(
                r"(?<=[\u3400-\u4dbf\u4e00-\u9fff])[,.;:?!](?=\s|$|[\u3400-\u4dbf\u4e00-\u9fff])|"
                r"(?<=[\u3400-\u4dbf\u4e00-\u9fff])[,.;:?!](?=[A-Za-z0-9])|"
                r"(?<=[A-Za-z0-9])[,.;:?!](?=[\u3400-\u4dbf\u4e00-\u9fff])",
                cleaned,
            )
            ascii_quotes_with_cjk = re.search(
                r"\"[^\"\n]*[\u3400-\u4dbf\u4e00-\u9fff][^\"\n]*\"|'[^'\n]*[\u3400-\u4dbf\u4e00-\u9fff][^'\n]*'",
                cleaned,
            )
            if ascii_between_cjk or ascii_quotes_with_cjk:
                add_issue(
                    issues,
                    "FAIL",
                    "ZH_PUNCT",
                    f"Paragraph {index} has likely ASCII punctuation in Chinese prose: {sample}",
                )
        else:
            if CHINESE_PUNCT_RE.search(cleaned):
                add_issue(
                    issues,
                    "FAIL",
                    "EN_PUNCT",
                    f"Paragraph {index} has Chinese punctuation in an English document: {sample}",
                )


def audit_abstract_keywords(paragraphs: list[ET.Element], issues: list[Issue]) -> None:
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        normalized = re.sub(r"\s+", "", text).lower()
        if normalized.startswith(("摘要", "关键词", "abstract", "keywords")):
            if paragraph_alignment(paragraph) == "center" or has_nonzero_indent(paragraph):
                add_issue(
                    issues,
                    "FAIL",
                    "ABSTRACT_INDENT",
                    f"Paragraph {index} looks like abstract/keywords but is centered or indented: {text[:80]}",
                )


def audit_headings(paragraphs: list[ET.Element], issues: list[Issue]) -> None:
    previous_text = ""
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if CHINESE_H1_RE.match(text):
            if paragraph_alignment(paragraph) == "center":
                add_issue(issues, "FAIL", "H1_CENTER", f"Level-1 heading is centered at paragraph {index}: {text[:80]}")
            if index > 1 and previous_text and not has_spacing_before(paragraph):
                add_issue(
                    issues,
                    "WARN",
                    "H1_GAP",
                    f"Level-1 heading may need a blank visual gap before it at paragraph {index}: {text[:80]}",
                )
        if text:
            previous_text = text


def audit_tables(root: ET.Element, issues: list[Issue]) -> None:
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        for run in table.findall(".//w:r", NS):
            if not paragraph_text(run).strip():
                continue
            size = w_attr(run.find("w:rPr/w:sz", NS), "val")
            size_cs = w_attr(run.find("w:rPr/w:szCs", NS), "val")
            for value in (size, size_cs):
                if value not in (None, "24"):
                    add_issue(
                        issues,
                        "FAIL",
                        "TABLE_SIZE",
                        f"Table {table_index} has direct font size {value}; expected small-four/12 pt (w:sz=24).",
                    )
                    break


def audit_color(root: ET.Element, issues: list[Issue]) -> None:
    for color in root.findall(".//w:color", NS):
        value = (w_attr(color, "val") or "").upper()
        if value and value not in ("000000", "AUTO", "C00000"):
            add_issue(
                issues,
                "WARN",
                "COLOR",
                f"Found direct font color {value}; verify it is intentional. Body/captions/tables/references should be black.",
            )


def audit_citations(document_xml: str, body_text: str, issues: list[Issue]) -> None:
    if FULLWIDTH_CITATION_RE.search(body_text):
        add_issue(issues, "FAIL", "CITATION_BRACKETS", "Body citations use full-width or non-ASCII brackets.")
    if CITATION_RE.search(body_text) and "REF ref_" not in document_xml:
        add_issue(
            issues,
            "FAIL",
            "CITATION_FIELDS",
            "Body citation patterns exist, but no Word REF ref_### cross-reference fields were found.",
        )


def load_document_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as archive:
        return archive.read("word/document.xml").decode("utf-8")


def is_bibliography_heading(text: str) -> bool:
    stripped = re.sub(r"\s+", "", text.strip())
    return bool(CN_BIB_HEADING_RE.match(stripped) or EN_BIB_HEADING_RE.match(stripped))


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit a DOCX against Allen's Word report formatting guardrails.")
    parser.add_argument("docx", type=Path, help="Path to the DOCX file to audit.")
    parser.add_argument("--language", choices=("zh", "en"), help="Override dominant document language detection.")
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        return 2

    document_xml = load_document_xml(args.docx)
    root = ET.fromstring(document_xml)
    paragraphs = root.findall(".//w:body//w:p", NS)
    texts = [paragraph_text(paragraph) for paragraph in paragraphs]
    bib_index = next((index for index, text in enumerate(texts) if is_bibliography_heading(text)), len(texts))
    body_text = "\n".join(texts[:bib_index])
    detected_language, cjk_count, latin_count = detect_language(body_text)
    language = args.language or detected_language

    issues: list[Issue] = []
    audit_punctuation(paragraphs[:bib_index], language, issues)
    audit_abstract_keywords(paragraphs, issues)
    audit_headings(paragraphs, issues)
    audit_tables(root, issues)
    audit_color(root, issues)
    audit_citations(document_xml, body_text, issues)

    print(f"Document language: {language} (CJK chars={cjk_count}, Latin letters={latin_count})")
    if not issues:
        print("PASS: no machine-detected guardrail issues.")
        return 0

    for severity in ("FAIL", "WARN"):
        for issue in issues:
            if issue.severity == severity:
                print(f"{issue.severity} {issue.code}: {issue.message}")

    return 1 if any(issue.severity == "FAIL" for issue in issues) else 0


if __name__ == "__main__":
    raise SystemExit(main())
