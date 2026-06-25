#!/usr/bin/env python3
"""Apply safe, mechanical formatting fixes to a DOCX.

This is the conservative companion to ``audit_docx_format.py``. It only performs
fixes that are unambiguous and low risk, working on the raw ``word/document.xml``
text so every other byte of the package is preserved:

1. Full-width / lenticular citation brackets around numbers become ASCII:
   ``［1］`` / ``【1,2】`` -> ``[1]`` / ``[1,2]``.
2. ASCII sentence punctuation sitting *between two CJK characters* becomes its
   full-width form: ``中文,中文`` -> ``中文，中文``. Decimals (``3.14``), URLs,
   citation brackets, and any punctuation next to a tag boundary are left alone,
   because the lookbehind/lookahead require a CJK character on both sides.

It does NOT touch fonts, styles, formulas, or citations cross-references — those
need judgement and belong to the model + the main standard. By default it writes
a new ``*.normalized.docx`` file and leaves the input untouched.

Usage:
    python scripts/normalize_docx.py report.docx                 # -> report.normalized.docx
    python scripts/normalize_docx.py report.docx -o fixed.docx
    python scripts/normalize_docx.py report.docx --in-place
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


def normalize_document_xml(xml: str) -> tuple[str, dict[str, int]]:
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


def normalize_docx(src: Path, dst: Path) -> dict[str, int]:
    with zipfile.ZipFile(src) as archive:
        names = archive.namelist()
        if "word/document.xml" not in names:
            raise ValueError("Input does not contain word/document.xml; not a normal DOCX.")
        data = {name: archive.read(name) for name in names}

    new_xml, counts = normalize_document_xml(data["word/document.xml"].decode("utf-8"))
    data["word/document.xml"] = new_xml.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.writestr(name, data[name])
    return counts


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply safe mechanical formatting fixes to a DOCX.")
    parser.add_argument("docx", type=Path, help="Path to the DOCX file to normalize.")
    parser.add_argument("-o", "--output", type=Path, help="Output path (default: <name>.normalized.docx).")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input file.")
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"ERROR: file not found: {args.docx}", file=sys.stderr)
        return 2

    if args.in_place:
        destination = args.docx
    elif args.output:
        destination = args.output
    else:
        destination = args.docx.with_suffix(".normalized.docx")

    try:
        if args.in_place:
            temp = args.docx.with_suffix(".normalize.tmp")
            counts = normalize_docx(args.docx, temp)
            temp.replace(args.docx)
        else:
            counts = normalize_docx(args.docx, destination)
    except (zipfile.BadZipFile, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(f"Wrote {destination}")
    print(f"Citation brackets normalized: {counts['citation_brackets']}")
    print(f"CJK punctuation normalized: {counts['cjk_punctuation']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
