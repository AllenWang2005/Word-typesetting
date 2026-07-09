#!/usr/bin/env python3
"""Apply safe, mechanical formatting fixes to a DOCX.

This is the conservative companion to ``audit_docx_format.py``. It only performs
fixes that are unambiguous and low risk, working on the raw ``word/document.xml``
text so every other byte of the package is preserved.

Default fixes (always applied):

1. Full-width / lenticular citation brackets around numbers become ASCII:
   ``［1］`` / ``【1,2】`` -> ``[1]`` / ``[1,2]``.
2. ASCII sentence punctuation sitting *between two CJK characters* becomes its
   full-width form: ``中文,中文`` -> ``中文，中文``. Decimals (``3.14``), URLs,
   citation brackets, and any punctuation next to a tag boundary are left alone,
   because the lookbehind/lookahead require a CJK character on both sides.

Opt-in fixes:

``--units``          insert the half-width space between a number and its unit
                     (``20km`` -> ``20 km``) and remove the space before ``%`` /
                     ``°`` / ``℃`` (``50 %`` -> ``50%``), inside w:t text only.
``--tables``         white three-line hygiene: clear every ``w:shd`` inside
                      tables to ``clear/auto``, zero the ``w:tblLook``
                      conditional-formatting flags, clear row-level exception
                      borders (``w:tblPrEx/w:tblBorders``), and mark the first row of
                      every multi-row table to repeat across pages
                      (``w:tblHeader``). Borders are NOT touched — set them per
                      ``references/three-line-table-ooxml.md``.
                      Row-level exception borders are cleared because they are
                      inherited grid artifacts, not intentional three-line rules.
                      OMML formula runs inside normal data tables are stamped to
                      五号/10.5 pt (``w:sz=21``); body formulas stay 小四/12 pt.
``--update-fields``  set ``w:updateFields`` in word/settings.xml so MS Word
                     (the canonical renderer) refreshes TOC/REF fields on open.
``--all``            all of the above.

It does NOT touch fonts, styles, formula semantics, or citation cross-references
— those need judgement and belong to the model + the main standard (use
``scripts/replace_math.py`` for formula conversion). The only formula formatting
it changes is the safe table-context size stamp above. By default it writes a
new ``*.normalized.docx`` file and leaves the input untouched.

Known limitation: the punctuation fix works on raw XML, so a CJK character and
its punctuation split across two runs are not matched.

Usage:
    python scripts/normalize_docx.py report.docx                 # -> report.normalized.docx
    python scripts/normalize_docx.py report.docx --all -o fixed.docx
    python scripts/normalize_docx.py report.docx --tables --update-fields --in-place
"""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path

CJK = r"㐀-䶿一-鿿"
FULLWIDTH = {",": "，", ".": "。", ";": "；", ":": "：", "?": "？", "!": "！"}
BRACKET_NUMBER = r"\d(?:[\d,，\-–—\s]*\d)?"

WT_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)")
TBL_RE = re.compile(r"<w:tbl>.*?</w:tbl>", re.S)
SHD_RE = re.compile(r"<w:shd\b[^>]*/>")
TBLLOOK_RE = re.compile(r"<w:tblLook\b[^>]*/>")
TBLPREX_RE = re.compile(r"<w:tblPrEx\b[^>]*>.*?</w:tblPrEx>", re.S)
TBLPREX_BORDERS_RE = re.compile(r"<w:tblBorders\b[^>]*>.*?</w:tblBorders>", re.S)
TR_OPEN_RE = re.compile(r"<w:tr(?:\s[^>]*)?>")
TC_OPEN_RE = re.compile(r"<w:tc(?:\s[^>]*)?>")
MATH_RUN_RE = re.compile(r"<m:r\b[^>]*>.*?</m:r>", re.S)
WRPR_RE = re.compile(r"<w:rPr\b[^>]*>.*?</w:rPr>", re.S)
WSZ_RE = re.compile(r"<w:sz\b[^>]*/>")
WSZCS_RE = re.compile(r"<w:szCs\b[^>]*/>")
MR_OPEN_RE = re.compile(r"<m:r\b[^>]*>")
MRPR_RE = re.compile(r"<m:rPr\b[^>]*>.*?</m:rPr>", re.S)
TAG_RE = re.compile(r"<[^>]+>")
EQUATION_NUMBER_RE = re.compile(r"[(（]\s*(?:式\s*)?\d+\s*[-–—]\s*\d+\s*[)）]")
UNIT_GLUE_RE = re.compile(r"(\d)(km|mm|cm|kg|kN|kPa|MPa|kW|MW|kV|Hz|min)(?![A-Za-z])")
UNIT_GLUE_M_RE = re.compile(r"(\d)(m)(?=[²³/])")
SPACE_BEFORE_UNIT_RE = re.compile(r"(\d)[  　\t]+(%|℃|°C|°(?![CF]))")

CLEAR_SHD = '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'
ZERO_TBLLOOK = (
    '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
    'w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'
)
TABLE_FORMULA_SZ = '<w:sz w:val="21"/>'
TABLE_FORMULA_SZCS = '<w:szCs w:val="21"/>'


def normalize_document_xml(xml: str) -> tuple[str, dict[str, int]]:
    """The two always-on fixes: citation brackets and CJK punctuation."""
    counts = {"citation_brackets": 0, "cjk_punctuation": 0}

    def repl_bracket(match: "re.Match[str]") -> str:
        counts["citation_brackets"] += 1
        return "[" + match.group(1) + "]"

    xml = re.sub(rf"［\s*({BRACKET_NUMBER})\s*］", repl_bracket, xml)
    xml = re.sub(rf"【\s*({BRACKET_NUMBER})\s*】", repl_bracket, xml)

    def repl_punct(match: "re.Match[str]") -> str:
        counts["cjk_punctuation"] += 1
        return match.group(1) + FULLWIDTH[match.group(2)]

    xml = re.sub(rf"([{CJK}])([,.;:?!])(?=[{CJK}])", repl_punct, xml)
    return xml, counts


def fix_unit_spacing(xml: str) -> tuple[str, int]:
    """Fix number-unit spacing inside w:t text nodes only."""
    fixed = 0

    def process(match: "re.Match[str]") -> str:
        nonlocal fixed
        text = match.group(2)
        text, n1 = UNIT_GLUE_RE.subn(r"\1 \2", text)
        text, n2 = UNIT_GLUE_M_RE.subn(r"\1 \2", text)
        text, n3 = SPACE_BEFORE_UNIT_RE.subn(r"\1\2", text)
        fixed += n1 + n2 + n3
        return match.group(1) + text + match.group(3)

    return WT_RE.sub(process, xml), fixed


def fix_table_hygiene(xml: str) -> tuple[str, dict[str, int]]:
    """Clear in-table shading, zero tblLook flags, repeat the header row."""
    counts = {
        "shading_cleared": 0,
        "tbllook_zeroed": 0,
        "row_exception_borders_cleared": 0,
        "header_repeat_added": 0,
        "table_formula_size_fixed": 0,
    }

    def is_formula_layout_table(block: str) -> bool:
        # 1x2/1x3 equation-layout helper tables are not data tables; their
        # display formulas should remain body-sized.
        if len(TR_OPEN_RE.findall(block)) != 1 or len(TC_OPEN_RE.findall(block)) not in (2, 3):
            return False
        if "<m:oMath" not in block:
            return False
        text = TAG_RE.sub("", block)
        return bool(EQUATION_NUMBER_RE.search(text))

    def fix_math_run_size(match: "re.Match[str]") -> str:
        run = match.group(0)
        original = run

        def fix_rpr(match: "re.Match[str]") -> str:
            rpr = match.group(0)
            if WSZ_RE.search(rpr):
                rpr = WSZ_RE.sub(TABLE_FORMULA_SZ, rpr, count=1)
            else:
                rpr = rpr.replace("</w:rPr>", TABLE_FORMULA_SZ + "</w:rPr>", 1)
            if WSZCS_RE.search(rpr):
                rpr = WSZCS_RE.sub(TABLE_FORMULA_SZCS, rpr, count=1)
            else:
                rpr = rpr.replace("</w:rPr>", TABLE_FORMULA_SZCS + "</w:rPr>", 1)
            return rpr

        if WRPR_RE.search(run):
            run = WRPR_RE.sub(fix_rpr, run, count=1)
        else:
            insert = f"<w:rPr>{TABLE_FORMULA_SZ}{TABLE_FORMULA_SZCS}</w:rPr>"
            m_rpr = MRPR_RE.search(run)
            if m_rpr:
                run = run[: m_rpr.end()] + insert + run[m_rpr.end() :]
            else:
                opener = MR_OPEN_RE.search(run)
                if opener:
                    run = run[: opener.end()] + insert + run[opener.end() :]
        if run != original:
            counts["table_formula_size_fixed"] += 1
        return run

    def process_table(match: "re.Match[str]") -> str:
        block = match.group(0)

        def clear_shd(m: "re.Match[str]") -> str:
            if m.group(0) == CLEAR_SHD:
                return m.group(0)
            counts["shading_cleared"] += 1
            return CLEAR_SHD

        def zero_look(m: "re.Match[str]") -> str:
            if m.group(0) == ZERO_TBLLOOK:
                return m.group(0)
            counts["tbllook_zeroed"] += 1
            return ZERO_TBLLOOK

        block = SHD_RE.sub(clear_shd, block)
        block = TBLLOOK_RE.sub(zero_look, block)
        if not is_formula_layout_table(block):
            block = MATH_RUN_RE.sub(fix_math_run_size, block)

        def clear_row_exception_borders(m: "re.Match[str]") -> str:
            cleaned, removed = TBLPREX_BORDERS_RE.subn("", m.group(0))
            counts["row_exception_borders_cleared"] += removed
            return cleaned

        block = TBLPREX_RE.sub(clear_row_exception_borders, block)
        rows = TR_OPEN_RE.findall(block)
        if len(rows) >= 2:
            first_tr = TR_OPEN_RE.search(block)
            first_tr_end = block.find("</w:tr>", first_tr.end())
            first_row = block[first_tr.end() : first_tr_end]
            if "<w:tblHeader" not in first_row:
                if first_row.lstrip().startswith("<w:trPr>"):
                    insert_at = block.find("<w:trPr>", first_tr.end()) + len("<w:trPr>")
                    block = block[:insert_at] + "<w:tblHeader/>" + block[insert_at:]
                else:
                    block = (
                        block[: first_tr.end()]
                        + "<w:trPr><w:tblHeader/></w:trPr>"
                        + block[first_tr.end() :]
                    )
                counts["header_repeat_added"] += 1
        return block

    return TBL_RE.sub(process_table, xml), counts


def fix_settings_update_fields(settings_xml: str) -> tuple[str, bool]:
    """Ensure w:settings contains w:updateFields so Word refreshes fields on open."""
    if "updateFields" in settings_xml:
        return settings_xml, False
    match = re.search(r"<w:settings\b[^>]*>", settings_xml)
    if match is None:
        return settings_xml, False
    insert_at = match.end()
    return (
        settings_xml[:insert_at] + '<w:updateFields w:val="true"/>' + settings_xml[insert_at:],
        True,
    )


def normalize_docx(
    src: Path, dst: Path, units: bool = False, tables: bool = False, update_fields: bool = False
) -> dict[str, int]:
    with zipfile.ZipFile(src) as archive:
        names = archive.namelist()
        if "word/document.xml" not in names:
            raise ValueError("Input does not contain word/document.xml; not a normal DOCX.")
        data = {name: archive.read(name) for name in names}

    xml, counts = normalize_document_xml(data["word/document.xml"].decode("utf-8"))
    if units:
        xml, fixed = fix_unit_spacing(xml)
        counts["unit_spacing"] = fixed
    if tables:
        xml, table_counts = fix_table_hygiene(xml)
        counts.update(table_counts)
    if update_fields and "word/settings.xml" in data:
        settings, changed = fix_settings_update_fields(data["word/settings.xml"].decode("utf-8"))
        data["word/settings.xml"] = settings.encode("utf-8")
        counts["update_fields_set"] = int(changed)
    data["word/document.xml"] = xml.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.writestr(name, data[name])
    return counts


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply safe mechanical formatting fixes to a DOCX.")
    parser.add_argument("docx", type=Path, help="Path to the DOCX file to normalize.")
    parser.add_argument("-o", "--output", type=Path, help="Output path (default: <name>.normalized.docx).")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input file.")
    parser.add_argument("--units", action="store_true", help="Fix number-unit spacing (20km -> 20 km; 50 %% -> 50%%).")
    parser.add_argument("--tables", action="store_true", help="Clear in-table shading, zero tblLook, repeat header rows.")
    parser.add_argument("--update-fields", action="store_true", help="Set w:updateFields so Word refreshes TOC/REF on open.")
    parser.add_argument("--all", action="store_true", help="Enable every opt-in fix.")
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        return 2

    units = args.units or args.all
    tables = args.tables or args.all
    update_fields = args.update_fields or args.all

    if args.in_place:
        destination = args.docx
    elif args.output:
        destination = args.output
    else:
        destination = args.docx.with_suffix(".normalized.docx")

    try:
        if args.in_place:
            temp = args.docx.with_suffix(".normalize.tmp")
            counts = normalize_docx(args.docx, temp, units, tables, update_fields)
            temp.replace(args.docx)
        else:
            counts = normalize_docx(args.docx, destination, units, tables, update_fields)
    except (zipfile.BadZipFile, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(f"Wrote {destination}")
    for key, value in counts.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
