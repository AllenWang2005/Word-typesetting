#!/usr/bin/env python3
"""Lightweight DOCX audit for Allen's Word report formatting skill."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = f"{{{W_NS}}}"
NS = {"w": W_NS, "m": M_NS}
AUDIT_SCOPE = (
    "Main document story only: word/document.xml. Headers, footers, footnotes, "
    "endnotes, comments, and separate embedded parts are not audited by this script. "
    "Table-cell text is checked for font size and borders only, not punctuation/heading/formula."
)

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff]")
LATIN_RE = re.compile(r"[A-Za-z]")
H1_TEXT_RE = re.compile(r"^(?:[一二三四五六七八九十]+[、.．]|第?[一二三四五六七八九十0-9]+章\b|[0-9]+(?:[、.．]\s*|\s+).+)")
H1_STYLE_IDS = {"heading1", "标题1", "1"}
CHINESE_PUNCT_RE = re.compile(r"[，。；：？！、“”‘’（）【】［］《》]")
CITATION_RE = re.compile(r"[\[［【]\s*\d+(?:\s*(?:,|，|-|–|—)\s*\d+)*\s*[\]］】]")
FULLWIDTH_CITATION_RE = re.compile(r"[［【]\s*\d+(?:\s*(?:,|，|-|–|—)\s*\d+)*\s*[］】]")
# A superscript citation number with its brackets stripped, e.g. a bare "1" or "1,2"/"1-3".
SUPERSCRIPT_CITATION_RE = re.compile(r"^\d+(?:\s*[,，、\-–—]\s*\d+)*$")
# Heading detection for font/punctuation checks. Uses heading styles plus the unambiguous
# Chinese markers only (not the loose Arabic-number branch) to avoid flagging numbered prose.
HEADING_STYLE_RE = re.compile(r"^(?:heading[1-9]|标题[1-9]|[1-9])$")
STRICT_HEADING_TEXT_RE = re.compile(r"^(?:[一二三四五六七八九十]+[、.．]|第[一二三四五六七八九十0-9]+章)")
HEADING_TRAILING_PUNCT_RE = re.compile(r"[。，、；：！,.;:!]$")
# Font-family heuristics (Heiti vs Songti) for the heading/body font check.
HEI_FONT_RE = re.compile(r"黑体|hei|yahei|heiti", re.I)
SONG_FONT_RE = re.compile(r"宋|song|simsun|nsimsun|stsong", re.I)
# Caption markers used for caption-position and body-font checks.
TABLE_CAPTION_RE = re.compile(r"^表\s*\d")
FIGURE_CAPTION_RE = re.compile(r"^图\s*\d")
# Default Word hyperlink/visited-link colors that should not be flagged as stray color.
HYPERLINK_COLOR_DEFAULTS = {"0563C1", "954F72", "0000FF", "0000EE", "1155CC"}
# High-precision "looks like a heading" markers for the heading-without-style check:
# Chinese ordinals (一、), chapter/section words (第3章/第二节), or multilevel numbers (1.1 / 1.1.1).
LOOKS_LIKE_HEADING_RE = re.compile(
    r"^(?:[一二三四五六七八九十]+[、.．]|第[一二三四五六七八九十0-9]+[章节]|\d+\.\d+(?:\.\d+)?[ \t　])"
)
# Vertical / inner-vertical border tags: a three-line table must not have these.
VERTICAL_BORDER_TAGS = ("left", "right", "insideV", "start", "end")
CN_BIB_HEADING_RE = re.compile(r"^(?:第?[一二三四五六七八九十0-9]+(?:章|节)?[、.．]?)?(参考文献|参考资料)$")
EN_BIB_HEADING_RE = re.compile(r"^(?:\d+[.)]?)?references$", re.I)
VISIBLE_LATEX_RE = re.compile(r"\\(?:frac|sqrt|sum|prod|int|begin|mathrm|text|sin|exp|lg|max)|[_^]\{")
TEXT_EQUATION_RE = re.compile(
    r"(?<![A-Za-z])(?:[A-Za-z\u0370-\u03ff][A-Za-z0-9_{}\\]*|[A-Z])\s*(?:=|≤|≥|≈|<|>)"
)
TEXT_SUBSCRIPT_RE = re.compile(r"\b[A-Za-z\u0370-\u03ff]{1,4}_[A-Za-z0-9{}\\]+")
# Generic formula-context cues. Strong definition cues intentionally include bare quantity symbols
# such as "式中 Q 为流量，N 为出力" because report rules require these symbols to be rendered as OMML.
SYMBOL_CONTEXT_RE = re.compile(r"(式中|其中|符号|变量|量符号|计算式|公式)")
# Generic quantity symbols. Domain-specific symbol lists are intentionally avoided so the auditor
# stays general-purpose and catches unknown engineering symbols instead of only hydraulics terms.
SUBSCRIPT_OR_SUPERSCRIPT_QUANTITY_RE = re.compile(r"(?<![A-Za-zͰ-Ͽ])[A-Za-zͰ-Ͽ][_^]\{?[A-Za-z0-9Ͱ-Ͽ]")
BARE_QUANTITY_SYMBOL_RE = re.compile(r"(?<![A-Za-zͰ-Ͽ])[A-Za-zͰ-Ͽ](?![A-Za-zͰ-Ͽ])")
# A bare symbol definition. Branch 1 (symbol BEFORE the verb) keeps the strong cues 为/表示/=
# ("Q 为流量", "x ="). Branch 2 (verb BEFORE the symbol) only uses the unambiguous 记为/表示为/定义为,
# so ordinary prose such as "为 A 方案" or "取 N 个" is no longer misread as a formula.
BARE_SYMBOL_DEFINITION_RE = re.compile(
    r"(?<![A-Za-zͰ-Ͽ])([A-Za-zͰ-Ͽ])(?![A-Za-zͰ-Ͽ])\s*(?:为|表示|=)"
    r"|(?:记为|表示为|定义为)\s*([A-Za-zͰ-Ͽ])(?![A-Za-zͰ-Ͽ])"
)

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


class AuditInputError(Exception):
    """Raised when the input file cannot be audited as a normal DOCX."""


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


def add_issue(
    issues: list[Issue],
    severity: str,
    code: str,
    message: str,
    counts: Optional[dict[tuple[str, str], int]] = None,
    limit: int = 8,
) -> None:
    if counts is None:
        if sum(1 for issue in issues if issue.code == code) < limit:
            issues.append(Issue(severity, code, message))
        return
    key = (severity, code)
    counts[key] = counts.get(key, 0) + 1
    if counts[key] <= limit:
        issues.append(Issue(severity, code, message))


def detect_language(text: str) -> tuple[str, int, int]:
    cjk_count = len(CJK_RE.findall(text))
    latin_count = len(LATIN_RE.findall(text))
    language = "zh" if cjk_count >= max(40, latin_count * 0.25) else "en"
    return language, cjk_count, latin_count


def paragraph_alignment(paragraph: ET.Element) -> Optional[str]:
    return w_attr(paragraph.find("w:pPr/w:jc", NS), "val")


def paragraph_style(paragraph: ET.Element) -> str:
    return w_attr(paragraph.find("w:pPr/w:pStyle", NS), "val") or ""


def run_is_superscript(run: ET.Element) -> bool:
    return w_attr(run.find("w:rPr/w:vertAlign", NS), "val") == "superscript"


def is_heading_style(style_id: str) -> bool:
    normalized = re.sub(r"[\s_-]+", "", style_id).lower()
    return bool(HEADING_STYLE_RE.match(normalized)) or normalized.startswith("heading")


def is_heading(paragraph: ET.Element, text: str) -> bool:
    return is_heading_style(paragraph_style(paragraph)) or bool(STRICT_HEADING_TEXT_RE.match(text))


def heading_level(paragraph: ET.Element, text: str) -> Optional[int]:
    style = re.sub(r"[\s_-]+", "", paragraph_style(paragraph)).lower()
    match = re.search(r"(?:heading|标题)([1-9])", style) or re.fullmatch(r"([1-9])", style)
    if match:
        return int(match.group(1))
    if STRICT_HEADING_TEXT_RE.match(text):
        return 1
    return None


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall(".//w:t", NS))


def run_eastasia_font(run: ET.Element) -> Optional[str]:
    return w_attr(run.find("w:rPr/w:rFonts", NS), "eastAsia")


def is_heading1_style(style_id: str) -> bool:
    normalized = re.sub(r"[\s_-]+", "", style_id).lower()
    return normalized in H1_STYLE_IDS or normalized.endswith("heading1")


def is_level1_heading(paragraph: ET.Element, text: str) -> bool:
    return is_heading1_style(paragraph_style(paragraph)) or bool(H1_TEXT_RE.match(text))


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


def audit_punctuation(
    paragraphs: list[ET.Element],
    language: str,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
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
                    counts,
                )
        else:
            if CHINESE_PUNCT_RE.search(cleaned):
                add_issue(
                    issues,
                    "FAIL",
                    "EN_PUNCT",
                    f"Paragraph {index} has Chinese punctuation in an English document: {sample}",
                    counts,
                )


def audit_abstract_keywords(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
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
                    counts,
                )


def audit_headings(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    previous_text = ""
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if is_level1_heading(paragraph, text):
            if paragraph_alignment(paragraph) == "center":
                add_issue(issues, "FAIL", "H1_CENTER", f"Level-1 heading is centered at paragraph {index}: {text[:80]}", counts)
            if index > 1 and previous_text and not has_spacing_before(paragraph):
                add_issue(
                    issues,
                    "WARN",
                    "H1_GAP",
                    f"Level-1 heading may need a blank visual gap before it at paragraph {index}: {text[:80]}",
                    counts,
                )
        if text and is_heading(paragraph, text) and HEADING_TRAILING_PUNCT_RE.search(text):
            add_issue(
                issues,
                "WARN",
                "HEADING_PUNCT",
                f"Heading ends with punctuation at paragraph {index}: {text[:80]}",
                counts,
            )
        if text:
            previous_text = text


def audit_fonts(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Flag the heading/body font swap (headings should be Heiti, body Songti).

    Only direct run-level `w:rFonts` are inspected; fonts inherited from styles are not
    visible in word/document.xml, so this is a WARN-level guardrail, not a full check.
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        heading = is_heading(paragraph, text)
        level = heading_level(paragraph, text)
        caption = bool(TABLE_CAPTION_RE.match(text) or FIGURE_CAPTION_RE.match(text))
        # Cover titles and other centered display lines are not body prose; skip the body-font check.
        centered = paragraph_alignment(paragraph) == "center"
        for run in paragraph.findall(".//w:r", NS):
            if not run_text(run).strip():
                continue
            font = run_eastasia_font(run)
            if not font:
                continue
            # Level 1-2 headings should be Heiti; level-3 headings are intentionally bold Songti.
            if level in (1, 2) and SONG_FONT_RE.search(font) and not HEI_FONT_RE.search(font):
                add_issue(
                    issues,
                    "WARN",
                    "HEADING_FONT",
                    f"Level-1/2 heading paragraph {index} uses a Songti font '{font}'; they should be Heiti.",
                    counts,
                )
                break
            if not heading and not caption and not centered and HEI_FONT_RE.search(font) and not SONG_FONT_RE.search(font):
                add_issue(
                    issues,
                    "WARN",
                    "BODY_FONT",
                    f"Body paragraph {index} uses a Heiti font '{font}'; body text should be Songti.",
                    counts,
                )
                break


def audit_captions(root: ET.Element, issues: list[Issue], counts: Optional[dict[tuple[str, str], int]] = None) -> None:
    """Heuristic check: table captions go above tables, figure captions go below figures."""
    body = root.find("w:body", NS)
    if body is None:
        return
    items: list[str] = []
    for child in list(body):
        if child.tag == f"{W}tbl":
            items.append("table")
        elif child.tag == f"{W}p":
            text = paragraph_text(child).strip()
            if TABLE_CAPTION_RE.match(text):
                items.append("tcap")
            elif FIGURE_CAPTION_RE.match(text):
                items.append("fcap")
            elif (
                child.find(".//w:drawing", NS) is not None
                or child.find(".//w:pict", NS) is not None
                or child.find(".//w:object", NS) is not None
            ):
                items.append("fig")
            else:
                items.append("para")
    for i, kind in enumerate(items):
        prev = items[i - 1] if i > 0 else None
        nxt = items[i + 1] if i + 1 < len(items) else None
        if kind == "tcap" and nxt != "table" and prev == "table":
            add_issue(
                issues,
                "WARN",
                "CAPTION_POSITION",
                "A table caption appears below its table; table captions belong above the table.",
                counts,
            )
        elif kind == "fcap" and prev != "fig" and nxt == "fig":
            add_issue(
                issues,
                "WARN",
                "CAPTION_POSITION",
                "A figure caption appears above its figure; figure captions belong below the figure.",
                counts,
            )


def audit_tables(root: ET.Element, issues: list[Issue], counts: Optional[dict[tuple[str, str], int]] = None) -> None:
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        text_sizes: set[str] = set()
        for run in table.findall(".//w:r", NS):
            if not paragraph_text(run).strip():
                continue
            size = w_attr(run.find("w:rPr/w:sz", NS), "val")
            size_cs = w_attr(run.find("w:rPr/w:szCs", NS), "val")
            if size:
                text_sizes.add(size)
            for value in (size, size_cs):
                # Table text should be 五号 (w:sz=21); 小四 (24) is also accepted. Anything
                # else (especially larger than the body) is flagged.
                if value not in (None, "21", "24"):
                    add_issue(
                        issues,
                        "FAIL",
                        "TABLE_SIZE",
                        f"Table {table_index} has direct font size {value}; expected 五号/10.5 pt (w:sz=21; 小四/24 also accepted).",
                        counts,
                    )
                    break
        # OMML formula symbols carry their own size in m:r/w:rPr/w:sz. An unsized math run
        # inherits the body default (小四), so inside a 五号 table the symbols render larger
        # than the surrounding cell text. Plain w:r scanning above never sees these m:r runs.
        table_is_wuhao = "21" in text_sizes
        formula_flagged = False
        for math_run in table.findall(".//m:r", NS):
            if not "".join(node.text or "" for node in math_run.findall("m:t", NS)).strip():
                continue
            math_size = w_attr(math_run.find("w:rPr/w:sz", NS), "val")
            if math_size not in (None, "21", "24"):
                add_issue(
                    issues,
                    "FAIL",
                    "TABLE_SIZE",
                    f"Table {table_index} has a formula symbol at font size {math_size}; expected 五号/10.5 pt (w:sz=21; 小四/24 also accepted).",
                    counts,
                )
            elif table_is_wuhao and math_size != "21" and not formula_flagged:
                add_issue(
                    issues,
                    "WARN",
                    "TABLE_FORMULA_SIZE",
                    f"Table {table_index} has OMML/formula symbols larger than its 五号 (10.5 pt) text; "
                    "set in-table formulas to the table text size (w:sz=21), not the larger body 小四.",
                    counts,
                )
                formula_flagged = True


def audit_color(root: ET.Element, issues: list[Issue], counts: Optional[dict[tuple[str, str], int]] = None) -> None:
    # Colors inside hyperlinks are expected (blue link text), so exclude them.
    hyperlink_colors = {id(c) for hl in root.findall(".//w:hyperlink", NS) for c in hl.findall(".//w:color", NS)}
    for color in root.findall(".//w:color", NS):
        if id(color) in hyperlink_colors:
            continue
        if color.get(f"{W}themeColor"):
            # Theme colors are template-driven; treat as intentional.
            continue
        value = (w_attr(color, "val") or "").upper()
        if value and value not in ("000000", "AUTO", "C00000") and value not in HYPERLINK_COLOR_DEFAULTS:
            add_issue(
                issues,
                "WARN",
                "COLOR",
                f"Found direct font color {value}; verify it is intentional. Body/captions/tables/references should be black.",
                counts,
            )


def audit_heading_styles(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn when a paragraph looks like a heading but uses no Word heading style.

    Such pseudo-headings do not enter the navigation pane or an auto-generated TOC.
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if not text or len(text) > 40:
            continue
        if LOOKS_LIKE_HEADING_RE.match(text) and not is_heading_style(paragraph_style(paragraph)):
            add_issue(
                issues,
                "WARN",
                "HEADING_NO_STYLE",
                f"Paragraph {index} looks like a heading but uses no Word heading style (won't enter the TOC): {text[:60]}",
                counts,
            )


def border_is_visible(border: Optional[ET.Element]) -> bool:
    if border is None:
        return False
    return (w_attr(border, "val") or "").lower() not in ("", "none", "nil")


def audit_table_borders(root: ET.Element, issues: list[Issue], counts: Optional[dict[tuple[str, str], int]] = None) -> None:
    """Warn when a table has visible vertical/inner borders instead of a three-line layout."""
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        sources = [table.find("w:tblPr/w:tblBorders", NS)]
        sources.extend(table.findall(".//w:tc/w:tcPr/w:tcBorders", NS))
        for borders in sources:
            if borders is None:
                continue
            if any(border_is_visible(borders.find(f"w:{tag}", NS)) for tag in VERTICAL_BORDER_TAGS):
                add_issue(
                    issues,
                    "WARN",
                    "TABLE_BORDERS",
                    f"Table {table_index} has vertical/inner borders; use a white three-line table (top, header, bottom rules only).",
                    counts,
                )
                break


def audit_citations(
    document_xml: str,
    body_text: str,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    if FULLWIDTH_CITATION_RE.search(body_text):
        add_issue(issues, "FAIL", "CITATION_BRACKETS", "Body citations use full-width or non-ASCII brackets.", counts)
    if CITATION_RE.search(body_text) and "REF ref_" not in document_xml:
        add_issue(
            issues,
            "FAIL",
            "CITATION_FIELDS",
            "Body citation patterns exist, but no Word REF ref_### cross-reference fields were found.",
            counts,
        )


def audit_bare_citations(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn about superscript citation numbers that dropped their brackets.

    A citation should be the whole bracketed token (e.g. ``[1]``). A bare superscript
    number such as ``1`` is reported as a WARN, not a FAIL, because it might instead be
    an exponent. To limit false positives, a number whose base is alphanumeric
    (an exponent like ``m`` + ``2`` or ``10`` + ``5``) is not flagged.
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        segments: list[tuple[str, bool]] = []
        for run in paragraph.findall(".//w:r", NS):
            text = "".join(node.text or "" for node in run.findall(".//w:t", NS))
            if text:
                segments.append((text, run_is_superscript(run)))
        preceding = ""
        position = 0
        while position < len(segments):
            text, superscript = segments[position]
            if not superscript:
                preceding = text[-1]
                position += 1
                continue
            group = text
            cursor = position + 1
            while cursor < len(segments) and segments[cursor][1]:
                group += segments[cursor][0]
                cursor += 1
            token = group.strip()
            base_is_exponent = bool(preceding) and preceding.isascii() and preceding.isalnum()
            if (
                SUPERSCRIPT_CITATION_RE.match(token)
                and "[" not in group
                and "]" not in group
                and not base_is_exponent
            ):
                add_issue(
                    issues,
                    "WARN",
                    "CITATION_NO_BRACKETS",
                    f"Paragraph {index} has a bare superscript number '{token}' with no brackets; "
                    "a citation should be the whole bracketed [n], not a bare number.",
                    counts,
                )
            preceding = group[-1] if group else preceding
            position = cursor


def has_omml(paragraph: ET.Element) -> bool:
    return paragraph.find(".//m:oMath", NS) is not None or paragraph.find(".//m:oMathPara", NS) is not None


def audit_formulas(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        cleaned = strip_protected(text)
        sample = text.replace("\n", " ")[:100]
        visible_latex = VISIBLE_LATEX_RE.search(cleaned)
        text_equation = TEXT_EQUATION_RE.search(cleaned)
        text_subscript = TEXT_SUBSCRIPT_RE.search(cleaned)
        subscript_symbol = SUBSCRIPT_OR_SUPERSCRIPT_QUANTITY_RE.search(cleaned)
        definition_symbol = BARE_SYMBOL_DEFINITION_RE.search(cleaned)
        context_symbol = SYMBOL_CONTEXT_RE.search(cleaned) and BARE_QUANTITY_SYMBOL_RE.search(cleaned)
        if visible_latex:
            add_issue(
                issues,
                "FAIL",
                "VISIBLE_LATEX",
                f"Paragraph {index} contains visible LaTeX source instead of rendered Word OMML: {sample}",
                counts,
            )
        elif text_equation or text_subscript or subscript_symbol or definition_symbol or context_symbol:
            severity = "WARN" if has_omml(paragraph) else "FAIL"
            add_issue(
                issues,
                severity,
                "FORMULA_TEXT",
                f"Paragraph {index} has likely formula/quantity-symbol text that should be LaTeX-rendered OMML: {sample}",
                counts,
            )


def load_document_xml(docx_path: Path) -> str:
    try:
        with zipfile.ZipFile(docx_path) as archive:
            try:
                return archive.read("word/document.xml").decode("utf-8")
            except KeyError as exc:
                raise AuditInputError(
                    "This file does not contain word/document.xml; it may be encrypted, damaged, or not a normal DOCX."
                ) from exc
    except zipfile.BadZipFile as exc:
        raise AuditInputError("Input is not a valid DOCX zip package. Save .doc files as .docx before auditing.") from exc
    except RuntimeError as exc:
        raise AuditInputError(f"Could not read the DOCX package: {exc}") from exc


def is_bibliography_heading(text: str) -> bool:
    stripped = re.sub(r"\s+", "", text.strip())
    return bool(CN_BIB_HEADING_RE.match(stripped) or EN_BIB_HEADING_RE.match(stripped))


def summarize_issues(issues: list[Issue], counts: dict[tuple[str, str], int]) -> tuple[dict[str, int], list[dict[str, object]]]:
    if not counts:
        counts.update(Counter((issue.severity, issue.code) for issue in issues))
    printed = Counter((issue.severity, issue.code) for issue in issues)
    omitted = []
    summary = {"FAIL": 0, "WARN": 0, "printed": len(issues), "omitted": 0}
    for (severity, code), total in sorted(counts.items()):
        summary[severity] = summary.get(severity, 0) + total
        skipped = total - printed[(severity, code)]
        if skipped > 0:
            summary["omitted"] += skipped
            omitted.append({"severity": severity, "code": code, "omitted": skipped})
    return summary, omitted


def print_json_result(result: dict) -> None:
    print(json.dumps(result, ensure_ascii=False, indent=2))


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit a DOCX against Allen's Word report formatting guardrails.")
    parser.add_argument("docx", type=Path, help="Path to the DOCX file to audit.")
    parser.add_argument("--language", choices=("zh", "en"), help="Override dominant document language detection.")
    parser.add_argument("--json", action="store_true", dest="json_output", help="Emit machine-readable JSON output.")
    args = parser.parse_args()

    if not args.docx.exists():
        message = f"File not found: {args.docx}"
        if args.json_output:
            print_json_result({"status": "error", "error": message, "scope": AUDIT_SCOPE})
        else:
            print(f"ERROR: {message}", file=sys.stderr)
        return 2

    try:
        document_xml = load_document_xml(args.docx)
        root = ET.fromstring(document_xml)
    except (AuditInputError, ET.ParseError) as exc:
        message = str(exc)
        if args.json_output:
            print_json_result({"status": "error", "error": message, "scope": AUDIT_SCOPE})
        else:
            print(f"ERROR: {message}", file=sys.stderr)
        return 2

    all_paragraphs = root.findall(".//w:body//w:p", NS)
    # Paragraph-based checks audit the main story only; table-cell paragraphs are excluded because
    # numbers/units/short fragments in cells produced many heading/punctuation/formula false positives.
    table_paragraph_ids = {id(p) for table in root.findall(".//w:tbl", NS) for p in table.findall(".//w:p", NS)}
    paragraphs = [p for p in all_paragraphs if id(p) not in table_paragraph_ids]
    texts = [paragraph_text(paragraph) for paragraph in paragraphs]
    bib_index = next((index for index, text in enumerate(texts) if is_bibliography_heading(text)), len(texts))
    body_text = "\n".join(texts[:bib_index])
    detected_language, cjk_count, latin_count = detect_language(body_text)
    language = args.language or detected_language

    issues: list[Issue] = []
    issue_counts: dict[tuple[str, str], int] = {}
    audit_punctuation(paragraphs[:bib_index], language, issues, issue_counts)
    audit_abstract_keywords(paragraphs, issues, issue_counts)
    audit_headings(paragraphs, issues, issue_counts)
    audit_heading_styles(paragraphs[:bib_index], issues, issue_counts)
    audit_fonts(paragraphs[:bib_index], issues, issue_counts)
    audit_tables(root, issues, issue_counts)
    audit_table_borders(root, issues, issue_counts)
    audit_color(root, issues, issue_counts)
    audit_captions(root, issues, issue_counts)
    audit_citations(document_xml, body_text, issues, issue_counts)
    audit_bare_citations(paragraphs[:bib_index], issues, issue_counts)
    audit_formulas(paragraphs[:bib_index], issues, issue_counts)

    omml_count = len(root.findall(".//m:oMath", NS)) + len(root.findall(".//m:oMathPara", NS))
    summary, omitted = summarize_issues(issues, issue_counts)
    status = "fail" if summary["FAIL"] else "pass"

    if args.json_output:
        print_json_result(
            {
                "status": status,
                "scope": AUDIT_SCOPE,
                "language": language,
                "stats": {"cjk_chars": cjk_count, "latin_letters": latin_count, "omml_objects": omml_count},
                "summary": {
                    "fail": summary["FAIL"],
                    "warn": summary["WARN"],
                    "printed": summary["printed"],
                    "omitted": summary["omitted"],
                },
                "issues": [issue.__dict__ for issue in issues],
                "omitted_by_code": omitted,
            }
        )
        return 1 if summary["FAIL"] else 0

    print(f"Document language: {language} (CJK chars={cjk_count}, Latin letters={latin_count}, OMML objects={omml_count})")
    print(f"Audit scope: {AUDIT_SCOPE}")
    if not issues:
        print("PASS: no machine-detected guardrail issues.")
        print("Summary: FAIL=0 WARN=0 printed=0 omitted=0")
        return 0

    for severity in ("FAIL", "WARN"):
        for issue in issues:
            if issue.severity == severity:
                print(f"{issue.severity} {issue.code}: {issue.message}")

    for item in omitted:
        print(f"OMITTED {item['severity']} {item['code']}: {item['omitted']} additional issue(s) not shown.")

    print(f"Summary: FAIL={summary['FAIL']} WARN={summary['WARN']} printed={summary['printed']} omitted={summary['omitted']}")
    return 1 if summary["FAIL"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
