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
M = f"{{{M_NS}}}"
NS = {"w": W_NS, "m": M_NS}
AUDIT_SCOPE = (
    "Main document story only: word/document.xml (plus word/styles.xml for table "
    "style shading/borders). Headers, footers, footnotes, endnotes, comments, and "
    "separate embedded parts are not audited by this script. Table-cell text is "
    "checked for font size, borders, and shading only, not punctuation/heading/formula."
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
# Any digit or operator inside a formula run: if the run is italic, these render slanted.
FORMULA_DIGIT_OPERATOR_RE = re.compile(r"[\d+\-*/=<>≤≥≈×·∙()（）]")
# Two or more adjacent Latin letters in a formula: a unit, function name, or multi-letter
# coefficient — never one italic variable.
MULTILETTER_RE = re.compile(r"[A-Za-z]{2,}")
# An italic plain-text run that visibly is an equation or a value-with-unit: manual italic math.
ITALIC_TEXT_MATH_RE = re.compile(
    r"[=≤≥≈<>]\s*\d|\d\s*[=≤≥≈<>]|\d\s*(?:km|mm|cm|kg|kN|kPa|MPa|kW|MW|kV|Hz|min)\b|\d\s*m[²³]"
)
# A number glued directly to a multi-letter unit (or m²/m³): needs a half-width space.
UNIT_NOSPACE_RE = re.compile(r"\d(?:km|mm|cm|kg|kN|kPa|MPa|kW|MW|kV|Hz|min)(?![A-Za-z])|\dm(?=[²³/])")
# A space before %, °, ℃ — these attach directly to the number.
SPACE_BEFORE_ATTACHED_UNIT_RE = re.compile(r"\d[  　\t]+(?:%|℃|°C|°(?![CF]))")
# A caption number such as 表 1-1 / 图 2.3 at the start of a caption paragraph.
CAPTION_NUMBER_RE = re.compile(r"^(图|表)\s*(\d+(?:[-–—.]\d+)?)")
# A display-equation number such as (3-3) or (式 3-3).
EQUATION_NUMBER_RE = re.compile(r"[(（]\s*(?:式\s*)?\d+\s*[-–—]\s*\d+\s*[)）]")
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

# Non-numeric protected tokens (URLs, DOIs, emails, file names, citations, code).
TOKEN_PROTECTED_PATTERNS = [
    re.compile(r"https?://\S+", re.I),
    re.compile(r"\bdoi\s*:\s*\S+", re.I),
    re.compile(r"\b10\.\d{4,9}/\S+", re.I),
    re.compile(r"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b"),
    re.compile(r"\S+\.(?:docx?|xlsx?|pptx?|pdf|csv|txt|dat|py|r|m)\b", re.I),
    re.compile(r"\[[0-9,\-\s]+\]"),
    re.compile(r"\bref_\d+\b", re.I),
    re.compile(r"`[^`]+`"),
]

# Full protection additionally hides plain numbers (decimals, thousands separators);
# used by the punctuation/formula checks, but NOT by checks that need the numbers
# themselves (number-unit spacing, manual italic math).
PROTECTED_PATTERNS = TOKEN_PROTECTED_PATTERNS + [
    re.compile(r"\b\d+\.\d+\b"),
    re.compile(r"\b\d{1,3}(?:,\d{3})+(?:\.\d+)?\b"),
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


def m_attr(element: Optional[ET.Element], name: str) -> Optional[str]:
    if element is None:
        return None
    return element.get(f"{M}{name}")


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def strip_protected(text: str) -> str:
    cleaned = text
    for pattern in PROTECTED_PATTERNS:
        cleaned = pattern.sub(" ", cleaned)
    return cleaned


def strip_protected_tokens(text: str) -> str:
    """Like strip_protected, but keeps plain numbers (for number-sensitive checks)."""
    cleaned = text
    for pattern in TOKEN_PROTECTED_PATTERNS:
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


def run_is_italic(run: ET.Element) -> bool:
    italic = run.find("w:rPr/w:i", NS)
    if italic is not None and (w_attr(italic, "val") or "true").lower() not in ("false", "0", "off"):
        return True
    # OMML style uses the m:val attribute (m:sty m:val="i"), not w:val.
    sty = run.find("m:rPr/m:sty", NS)
    return (m_attr(sty, "val") or w_attr(sty, "val")) in ("i", "bi")


def math_run_is_upright(run: ET.Element) -> bool:
    """True when an m:r carries an explicit upright style (m:sty p/b or m:nor).

    OMML letters render in *math italic by default*: the absence of any italic
    marker does NOT mean upright. Units, function names, and explanatory
    subscripts must carry m:sty val="p" (what Pandoc emits for \\mathrm/\\text)
    or m:nor to actually display upright.
    """
    if run.find("m:rPr/m:nor", NS) is not None:
        return True
    sty = run.find("m:rPr/m:sty", NS)
    return (m_attr(sty, "val") or w_attr(sty, "val")) in ("p", "b")


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
        # Include OMML math runs (m:r) so in-table formulas/symbols are size-checked too,
        # not just ordinary w:r text. Formula size is carried on w:rPr/w:sz inside m:r.
        for run in table.findall(".//w:r", NS) + table.findall(".//m:r", NS):
            text = "".join(node.text or "" for node in run.findall(".//w:t", NS))
            text += "".join(node.text or "" for node in run.findall(".//m:t", NS))
            if not text.strip():
                continue
            size = w_attr(run.find("w:rPr/w:sz", NS), "val")
            size_cs = w_attr(run.find("w:rPr/w:szCs", NS), "val")
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


def border_size(border: Optional[ET.Element]) -> Optional[int]:
    value = w_attr(border, "sz")
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def shading_is_visible(shd: Optional[ET.Element]) -> bool:
    """True when a w:shd element actually paints a non-white background."""
    if shd is None:
        return False
    val = (w_attr(shd, "val") or "clear").lower()
    fill = (w_attr(shd, "fill") or "auto").upper()
    if val not in ("clear", "nil", "none"):
        return True
    return fill not in ("AUTO", "FFFFFF")


def build_style_map(styles_root: Optional[ET.Element]) -> dict[str, ET.Element]:
    if styles_root is None:
        return {}
    styles = {}
    for style in styles_root.findall("w:style", NS):
        style_id = style.get(f"{W}styleId")
        if style_id:
            styles[style_id] = style
    return styles


def style_chain(style_id: str, style_map: dict[str, ET.Element]) -> list[ET.Element]:
    """Return the style and its basedOn ancestors (cycle-safe, depth-limited)."""
    chain: list[ET.Element] = []
    seen: set[str] = set()
    current: Optional[str] = style_id
    while current and current not in seen and len(chain) < 8:
        seen.add(current)
        style = style_map.get(current)
        if style is None:
            break
        chain.append(style)
        current = w_attr(style.find("w:basedOn", NS), "val")
    return chain


def table_style_id(table: ET.Element) -> Optional[str]:
    return w_attr(table.find("w:tblPr/w:tblStyle", NS), "val")


def audit_table_shading(
    root: ET.Element,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
    styles_root: Optional[ET.Element] = None,
) -> None:
    """A three-line table must be white: no cell/table shading, direct or style-driven.

    Header shading usually comes from the table style's firstRow conditional format
    (in styles.xml), which is invisible in word/document.xml — so the referenced
    table style's chain is inspected too when styles.xml is available.
    """
    style_map = build_style_map(styles_root)
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        if any(shading_is_visible(shd) for shd in table.findall(".//w:shd", NS)):
            add_issue(
                issues,
                "FAIL",
                "TABLE_SHADING",
                f"Table {table_index} has cell/table shading; a three-line table must be all white "
                "(clear every w:shd to val=clear fill=auto).",
                counts,
            )
            continue
        style_id = table_style_id(table)
        if not style_id:
            continue
        for style in style_chain(style_id, style_map):
            if any(shading_is_visible(shd) for shd in style.findall(".//w:shd", NS)):
                add_issue(
                    issues,
                    "FAIL",
                    "TABLE_SHADING",
                    f"Table {table_index} references table style '{style_id}' whose definition adds shading "
                    "(often firstRow header shading); remove the style reference, zero the w:tblLook flags, "
                    "and set explicit borders instead.",
                    counts,
                )
                break


def resolve_table_borders(
    table: ET.Element, style_map: dict[str, ET.Element]
) -> Optional[ET.Element]:
    borders = table.find("w:tblPr/w:tblBorders", NS)
    if borders is not None:
        return borders
    style_id = table_style_id(table)
    if style_id:
        for style in style_chain(style_id, style_map):
            borders = style.find("w:tblPr/w:tblBorders", NS)
            if borders is not None:
                return borders
    return None


def audit_table_rules(
    root: ET.Element,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
    styles_root: Optional[ET.Element] = None,
) -> None:
    """Check the three horizontal rules: visible thick top/bottom, thinner header rule.

    Catches the classic "no top/bottom rule, thick middle line" failure and tables
    that were left borderless instead of getting explicit three-line borders.
    """
    style_map = build_style_map(styles_root)
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        borders = resolve_table_borders(table, style_map)
        if borders is None:
            add_issue(
                issues,
                "WARN",
                "TABLE_RULES",
                f"Table {table_index} has no explicit table borders; a three-line table needs visible thick "
                "top/bottom rules (w:sz≈12) and a thin header rule (w:sz≈6) set explicitly.",
                counts,
            )
            continue
        top = borders.find("w:top", NS)
        bottom = borders.find("w:bottom", NS)
        if not border_is_visible(top) or not border_is_visible(bottom):
            add_issue(
                issues,
                "WARN",
                "TABLE_RULES",
                f"Table {table_index} is missing a visible top or bottom rule; both must exist and be the "
                "thick rules (w:sz≈12) of the three-line table.",
                counts,
            )
        if border_is_visible(borders.find("w:insideH", NS)):
            add_issue(
                issues,
                "WARN",
                "TABLE_RULES",
                f"Table {table_index} has row-to-row horizontal borders (insideH); only the header rule may "
                "sit between the top and bottom rules.",
                counts,
            )
        top_size = border_size(top)
        if top_size is None:
            continue
        rows = table.findall("w:tr", NS)
        if not rows:
            continue
        header_sizes = []
        for cell_borders in rows[0].findall("w:tc/w:tcPr/w:tcBorders", NS):
            bottom_rule = cell_borders.find("w:bottom", NS)
            if border_is_visible(bottom_rule):
                size = border_size(bottom_rule)
                if size is not None:
                    header_sizes.append(size)
        if header_sizes and max(header_sizes) >= top_size:
            add_issue(
                issues,
                "WARN",
                "TABLE_RULES",
                f"Table {table_index}'s header rule is not thinner than its top/bottom rules; expected header "
                "w:sz≈6 vs top/bottom w:sz≈12.",
                counts,
            )


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


def audit_math_duplication(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """FAIL when an OMML object's text also survives as plain text in the same paragraph.

    This is the append-instead-of-replace failure: the model keeps the original
    plain-text token in the prose and appends rendered math (typically at the end
    of the paragraph) instead of swapping it in place.
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        maths = paragraph.findall(".//m:oMath", NS)
        if not maths:
            continue
        plain = re.sub(r"\s+", "", paragraph_text(paragraph))
        if not plain:
            continue
        for math in maths:
            math_text = re.sub(
                r"\s+", "", "".join(node.text or "" for node in math.findall(".//m:t", NS))
            )
            # Require a few characters so a lone symbol shared with unrelated prose
            # (e.g. an OMML "Q" plus the letter Q elsewhere) is not misflagged.
            if len(math_text) >= 4 and math_text in plain:
                add_issue(
                    issues,
                    "FAIL",
                    "MATH_DUPLICATE",
                    f"Paragraph {index} contains rendered math '{math_text[:30]}' that also remains as plain "
                    "text; the OMML must replace the original text in place, not be appended next to it.",
                    counts,
                )
                break


def audit_formula_digit_italics(root: ET.Element, issues: list[Issue], counts: Optional[dict[tuple[str, str], int]] = None) -> None:
    """Warn when a number/operator inside a formula is italic (only variables should be italic).

    Usually caused by blanket-italicizing the whole equation, which also slants digits.
    A mixed run such as an italic 'F=44.5' is flagged too — its digits are italic even
    though the run also contains a letter.
    """
    for math in root.findall(".//m:oMath", NS):
        for run in math.findall(".//m:r", NS):
            text = "".join(node.text or "" for node in run.findall(".//m:t", NS))
            if not text.strip():
                continue
            if run_is_italic(run) and FORMULA_DIGIT_OPERATOR_RE.search(text):
                add_issue(
                    issues,
                    "WARN",
                    "FORMULA_DIGIT_ITALIC",
                    f"A number/operator in a formula is italic ('{text.strip()[:20]}'); digits and operators must be upright, only variables italic.",
                    counts,
                )
                break


def audit_formula_multiletter_italics(
    root: ET.Element,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn about a multi-letter OMML run left at the default math-italic style.

    Two or more adjacent Latin letters in a formula are never a single italic
    variable: they are either a unit ('km', 'kW'), a function name ('max'), or a
    multi-letter coefficient that must be one variable + upright subscript
    ('C_I', not 'CI'). All of these need an explicit upright style — OMML letters
    are italic by default, so 'no italic marker' still renders slanted.
    """
    for math in root.findall(".//m:oMath", NS):
        for run in math.findall(".//m:r", NS):
            text = "".join(node.text or "" for node in run.findall(".//m:t", NS))
            if not MULTILETTER_RE.search(text):
                continue
            if not math_run_is_upright(run):
                add_issue(
                    issues,
                    "WARN",
                    "FORMULA_MULTILETTER_ITALIC",
                    f"Formula run '{text.strip()[:20]}' has 2+ adjacent letters at OMML's default italic; "
                    "units/functions need an explicit upright style (m:sty val=\"p\", LaTeX \\mathrm), and a "
                    "multi-letter coefficient must become one variable + upright subscript.",
                    counts,
                )
                break


def audit_manual_italic_math(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """FAIL on a manually italicized plain-text pseudo-formula.

    The observed failure mode: instead of OMML, the model writes 'F = 44.5 km²'
    as an ordinary italic text run — which both fakes the equation and slants
    the digits/units. Only ordinary w:r runs are inspected (OMML has its own checks).
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        for run in paragraph.findall("w:r", NS):
            text = run_text(run)
            if not text.strip() or not run_is_italic(run):
                continue
            if ITALIC_TEXT_MATH_RE.search(strip_protected_tokens(text)):
                add_issue(
                    issues,
                    "FAIL",
                    "MANUAL_ITALIC_MATH",
                    f"Paragraph {index} has an italicized plain-text formula ('{text.strip()[:40]}'); this must "
                    "be a LaTeX-rendered OMML object, and italics never apply to digits/units anyway.",
                    counts,
                )
                break


def audit_caption_alignment(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn when a 表/图 caption paragraph is not centered.

    A caption using a named style with no direct alignment is left alone (the
    style may center it); direct-formatted captions must carry jc=center.
    """
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph).strip()
        if not (TABLE_CAPTION_RE.match(text) or FIGURE_CAPTION_RE.match(text)):
            continue
        alignment = paragraph_alignment(paragraph)
        if alignment == "center":
            continue
        if alignment is None and paragraph_style(paragraph):
            continue
        add_issue(
            issues,
            "WARN",
            "CAPTION_ALIGN",
            f"Caption paragraph {index} is not centered: {text[:60]}",
            counts,
        )


def audit_number_unit_spacing(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn on number–unit spacing errors: '20km' needs a space, '50 %' must attach."""
    for index, paragraph in enumerate(paragraphs, start=1):
        text = paragraph_text(paragraph)
        if not text.strip():
            continue
        cleaned = strip_protected_tokens(text)
        sample = text.strip().replace("\n", " ")[:80]
        if UNIT_NOSPACE_RE.search(cleaned):
            add_issue(
                issues,
                "WARN",
                "NUMBER_UNIT_SPACING",
                f"Paragraph {index} has a number glued to its unit (e.g. '20km'); use a half-width "
                f"space: '20 km': {sample}",
                counts,
            )
        if SPACE_BEFORE_ATTACHED_UNIT_RE.search(cleaned):
            add_issue(
                issues,
                "WARN",
                "NUMBER_UNIT_SPACING",
                f"Paragraph {index} has a space before %, °, or ℃; these attach directly (50%, 25℃): {sample}",
                counts,
            )


def audit_float_order(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn when a figure/table appears before its first in-text reference.

    Only the unambiguous inversion is flagged: the caption has no mention earlier
    in the document but IS mentioned later ("先出现后引用"). A caption with no
    mention at all is left to human review.
    """
    texts = [paragraph_text(p) for p in paragraphs]
    for index, text in enumerate(texts):
        match = CAPTION_NUMBER_RE.match(text.strip())
        if not match:
            continue
        kind, number = match.group(1), match.group(2)
        token = re.compile(rf"{kind}\s*{re.escape(number)}(?![\d.\-–—])")
        mentioned_before = any(token.search(t) for t in texts[:index])
        mentioned_after = any(token.search(t) for t in texts[index + 1 :])
        if not mentioned_before and mentioned_after:
            add_issue(
                issues,
                "WARN",
                "FLOAT_ORDER",
                f"{kind} {number} appears before its first in-text reference; figures/tables must be "
                "referenced first, then placed nearby after the reference.",
                counts,
            )


def audit_table_header_repeat(
    root: ET.Element,
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn when a multi-row table's header row is not set to repeat across pages."""
    for table_index, table in enumerate(root.findall(".//w:tbl", NS), start=1):
        rows = table.findall("w:tr", NS)
        if len(rows) < 2:
            continue
        header = rows[0].find("w:trPr/w:tblHeader", NS)
        repeats = header is not None and (w_attr(header, "val") or "true").lower() not in ("false", "0", "off")
        if not repeats:
            add_issue(
                issues,
                "WARN",
                "TABLE_HEADER_REPEAT",
                f"Table {table_index}'s header row is not marked to repeat across page breaks "
                "(w:trPr/w:tblHeader on the first row).",
                counts,
            )


def audit_equation_numbers(
    paragraphs: list[ET.Element],
    issues: list[Issue],
    counts: Optional[dict[tuple[str, str], int]] = None,
) -> None:
    """Warn when a numbered display formula is centered (the number should be right-aligned)."""
    for index, paragraph in enumerate(paragraphs, start=1):
        if paragraph.find(".//m:oMath", NS) is None and paragraph.find(".//m:oMathPara", NS) is None:
            continue
        text = paragraph_text(paragraph)
        if EQUATION_NUMBER_RE.search(text) and paragraph_alignment(paragraph) == "center":
            add_issue(
                issues,
                "WARN",
                "EQUATION_NUMBER_CENTER",
                "Paragraph {0} centers a numbered formula; the number should be right-aligned "
                "(left-aligned paragraph with a center tab for the equation and a right tab for the number): {1}".format(
                    index, text.strip()[:50]
                ),
                counts,
            )


def load_part(docx_path: Path, name: str) -> Optional[str]:
    """Read an optional part (e.g. word/styles.xml) from the DOCX; None when absent."""
    try:
        with zipfile.ZipFile(docx_path) as archive:
            try:
                return archive.read(name).decode("utf-8")
            except KeyError:
                return None
    except (zipfile.BadZipFile, RuntimeError):
        return None


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

    styles_root: Optional[ET.Element] = None
    styles_xml = load_part(args.docx, "word/styles.xml")
    if styles_xml:
        try:
            styles_root = ET.fromstring(styles_xml)
        except ET.ParseError:
            styles_root = None

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
    audit_table_shading(root, issues, issue_counts, styles_root)
    audit_table_rules(root, issues, issue_counts, styles_root)
    audit_table_header_repeat(root, issues, issue_counts)
    audit_color(root, issues, issue_counts)
    audit_captions(root, issues, issue_counts)
    audit_caption_alignment(paragraphs, issues, issue_counts)
    audit_float_order(paragraphs, issues, issue_counts)
    audit_citations(document_xml, body_text, issues, issue_counts)
    audit_bare_citations(paragraphs[:bib_index], issues, issue_counts)
    audit_formulas(paragraphs[:bib_index], issues, issue_counts)
    audit_math_duplication(paragraphs, issues, issue_counts)
    audit_manual_italic_math(paragraphs[:bib_index], issues, issue_counts)
    audit_formula_digit_italics(root, issues, issue_counts)
    audit_formula_multiletter_italics(root, issues, issue_counts)
    audit_number_unit_spacing(paragraphs[:bib_index], issues, issue_counts)
    audit_equation_numbers(paragraphs[:bib_index], issues, issue_counts)

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
