#!/usr/bin/env python3
"""Replace plain-text math tokens in a DOCX with native OMML equations, in place.

This is the deterministic tool for the formula workflow's In-Place Replacement
Contract: instead of hand-editing OOXML (the error mode that produced appended
italic pseudo-formulas), the model prepares a JSON registry and this script does
the exact-position splice, preserving every character around each token.

Registry: a JSON list, one entry per distinct token:

    [
      {"find": "F=44.5 km²", "latex": "F = 44.5\\\\,\\\\mathrm{km^2}"},
      {"find": "Q_p=W/T", "latex": "Q_p = W/T", "sz": 21},
      {"find": "W = 0.278KFP", "latex": "W = 0.278\\\\,K F P",
       "display": true, "number": "(3-1)"}
    ]

Entry fields:
    find     exact text to locate (required). All occurrences are replaced.
    latex    LaTeX source; converted with Pandoc (must be installed).
    omml     pre-converted <m:oMath> XML; used instead of latex when present,
             so the script also works without Pandoc.
    display  the token must be an entire paragraph; it is rebuilt as a display
             equation line (left-aligned, center tab for the equation, right
             tab for the number). Default false (inline).
    number   display only: right-aligned equation number text, e.g. "(3-1)".
    sz       font size in half-points stamped onto the equation runs
             (24 = 小四 body, 21 = 五号 table). Default: --default-sz (24).

Usage:
    python scripts/replace_math.py report.docx registry.json -o out.docx
    python scripts/replace_math.py report.docx registry.json --in-place
    python scripts/replace_math.py --convert 'Q = \\frac{W}{T}'   # print OMML

The output summary lists replaced counts and any token that was not found or
that still remains as plain text — treat leftovers as a failed pass.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = f"{{{W_NS}}}"
M = f"{{{M_NS}}}"
NS = {"w": W_NS, "m": M_NS}

XMLNS_RE = re.compile(r'xmlns:([A-Za-z0-9_]+)="([^"]+)"')
ROOT_TAG_RE = re.compile(r"<w:document\b[^>]*>")
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# A4 with ~2.5 cm margins, in twips, used when the document has no sectPr.
DEFAULT_PAGE_W, DEFAULT_MARGIN = 11906, 1417


class ReplaceMathError(Exception):
    pass


def find_pandoc() -> Optional[str]:
    """Locate pandoc on PATH or in the usual install spots (brew, ~/.local/bin)."""
    found = shutil.which("pandoc")
    if found:
        return found
    for candidate in (
        Path.home() / ".local/bin/pandoc",
        Path("/opt/homebrew/bin/pandoc"),
        Path("/usr/local/bin/pandoc"),
    ):
        if candidate.is_file():
            return str(candidate)
    return None


def register_namespaces(xml: str) -> None:
    """Register every prefix declared in the source so ET keeps them on output."""
    for prefix, uri in XMLNS_RE.findall(xml):
        ET.register_namespace(prefix, uri)


def latex_to_omml(latex_fragments: list[tuple[str, bool]]) -> list[ET.Element]:
    """Convert LaTeX fragments to <m:oMath> elements via Pandoc (one batch call).

    Each fragment is (latex, display). Returns oMath elements in input order.
    """
    if not latex_fragments:
        return []
    pandoc = find_pandoc()
    if pandoc is None:
        raise ReplaceMathError(
            "Pandoc is required to convert LaTeX (brew install pandoc / apt install pandoc), "
            "or provide pre-converted OMML via the registry's \"omml\" field."
        )
    lines = []
    for latex, display in latex_fragments:
        wrapped = f"$${latex}$$" if display else f"${latex}$"
        lines.append(wrapped)
        lines.append("")
    with tempfile.TemporaryDirectory() as folder:
        out = Path(folder) / "math.docx"
        result = subprocess.run(
            [pandoc, "-f", "markdown+tex_math_dollars", "-t", "docx", "-o", str(out)],
            input="\n".join(lines).encode("utf-8"),
            capture_output=True,
        )
        if result.returncode != 0:
            raise ReplaceMathError(f"Pandoc failed: {result.stderr.decode('utf-8', 'replace')[:400]}")
        with zipfile.ZipFile(out) as archive:
            document = archive.read("word/document.xml").decode("utf-8")
    root = ET.fromstring(document)
    maths = root.findall(".//m:oMath", NS)
    if len(maths) != len(latex_fragments):
        raise ReplaceMathError(
            f"Expected {len(latex_fragments)} converted equations, got {len(maths)}; "
            "check the LaTeX sources (one formula per registry entry)."
        )
    return [copy.deepcopy(math) for math in maths]


def stamp_font_size(omath: ET.Element, sz: int) -> None:
    """Set w:sz/w:szCs on every math run so the equation matches its context size."""
    for run in omath.iter(f"{M}r"):
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(f"{W}rPr")
            m_rpr = run.find("m:rPr", NS)
            run.insert(list(run).index(m_rpr) + 1 if m_rpr is not None else 0, rpr)
        for tag in ("sz", "szCs"):
            node = rpr.find(f"w:{tag}", NS)
            if node is None:
                node = ET.SubElement(rpr, f"{W}{tag}")
            node.set(f"{W}val", str(sz))


def run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.findall("w:t", NS))


def make_text_run(template: ET.Element, text: str) -> ET.Element:
    """A new w:r with the template's run properties and the given literal text."""
    run = ET.Element(f"{W}r")
    rpr = template.find("w:rPr", NS)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = ET.SubElement(run, f"{W}t")
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    return run


def paragraph_plain_text(paragraph: ET.Element) -> str:
    """Text of ordinary w:t nodes (math m:t is separate)."""
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def replace_inline(paragraph: ET.Element, token: str, omath: ET.Element) -> int:
    """Replace every occurrence of token in the paragraph's direct runs, in place."""
    replaced = 0
    while True:
        runs = [child for child in paragraph if child.tag == f"{W}r"]
        offsets: list[tuple[ET.Element, int, int]] = []
        cursor = 0
        for run in runs:
            text = run_text(run)
            offsets.append((run, cursor, cursor + len(text)))
            cursor += len(text)
        full = "".join(run_text(run) for run in runs)
        position = full.find(token)
        if position == -1:
            return replaced
        end = position + len(token)
        span = [
            (run, start, stop)
            for run, start, stop in offsets
            if stop > position and start < end and stop > start
        ]
        if not span:
            return replaced
        first_run, first_start, _ = span[0]
        last_run, last_start, _ = span[-1]
        prefix = run_text(first_run)[: position - first_start]
        suffix = run_text(last_run)[end - last_start :]
        children = list(paragraph)
        insert_at = children.index(first_run)
        # Remove every element between (and including) the first and last spanned run.
        for child in children[insert_at : children.index(last_run) + 1]:
            paragraph.remove(child)
        new_nodes: list[ET.Element] = []
        if prefix:
            new_nodes.append(make_text_run(first_run, prefix))
        new_nodes.append(copy.deepcopy(omath))
        if suffix:
            new_nodes.append(make_text_run(last_run, suffix))
        for offset, node in enumerate(new_nodes):
            paragraph.insert(insert_at + offset, node)
        replaced += 1


def content_width(root: ET.Element) -> int:
    sect = root.find(".//w:body/w:sectPr", NS)
    page_w, left, right = DEFAULT_PAGE_W, DEFAULT_MARGIN, DEFAULT_MARGIN
    if sect is not None:
        size = sect.find("w:pgSz", NS)
        margin = sect.find("w:pgMar", NS)
        try:
            if size is not None:
                page_w = int(size.get(f"{W}w", page_w))
            if margin is not None:
                left = int(margin.get(f"{W}left", left))
                right = int(margin.get(f"{W}right", right))
        except ValueError:
            pass
    return max(page_w - left - right, 2000)


def make_display_paragraph(
    paragraph: ET.Element, omath: ET.Element, number: Optional[str], width: int
) -> None:
    """Rebuild the paragraph as: [center tab] equation [right tab] number.

    Left-aligned with explicit tab stops so the number is right-aligned without
    centering the whole line (the EQUATION_NUMBER_CENTER failure).
    """
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(f"{W}pPr")
        paragraph.insert(0, ppr)
    for tag in ("w:tabs", "w:jc"):
        node = ppr.find(tag, NS)
        if node is not None:
            ppr.remove(node)
    tabs = ET.Element(f"{W}tabs")
    for val, pos in (("center", width // 2), ("right", width)):
        tab = ET.SubElement(tabs, f"{W}tab")
        tab.set(f"{W}val", val)
        tab.set(f"{W}pos", str(pos))
    style = ppr.find("w:pStyle", NS)
    ppr.insert(list(ppr).index(style) + 1 if style is not None else 0, tabs)
    jc = ET.SubElement(ppr, f"{W}jc")
    jc.set(f"{W}val", "left")
    for child in [c for c in paragraph if c.tag != f"{W}pPr"]:
        paragraph.remove(child)

    def tab_run() -> ET.Element:
        run = ET.Element(f"{W}r")
        ET.SubElement(run, f"{W}tab")
        return run

    paragraph.append(tab_run())
    paragraph.append(copy.deepcopy(omath))
    if number:
        paragraph.append(tab_run())
        number_run = ET.Element(f"{W}r")
        node = ET.SubElement(number_run, f"{W}t")
        node.text = number
        paragraph.append(number_run)


def apply_registry(document_xml: str, registry: list[dict], default_sz: int) -> tuple[str, dict]:
    register_namespaces(document_xml)
    root_tag_match = ROOT_TAG_RE.search(document_xml)
    if root_tag_match is None:
        raise ReplaceMathError("word/document.xml has no <w:document> root.")
    root = ET.fromstring(document_xml)
    width = content_width(root)

    # Convert every LaTeX entry in one Pandoc batch.
    latex_entries = [e for e in registry if "omml" not in e]
    maths = latex_to_omml([(e["latex"], bool(e.get("display"))) for e in latex_entries])
    for entry, math in zip(latex_entries, maths):
        entry["_omath"] = math
    for entry in registry:
        if "omml" in entry:
            fragment = ET.fromstring(entry["omml"])
            entry["_omath"] = fragment if fragment.tag == f"{M}oMath" else fragment.find("m:oMath", NS)
            if entry["_omath"] is None:
                raise ReplaceMathError(f"Entry '{entry['find']}': omml must contain an <m:oMath> element.")
        stamp_font_size(entry["_omath"], int(entry.get("sz", default_sz)))

    counts = {entry["find"]: 0 for entry in registry}
    paragraphs = root.findall(".//w:p", NS)
    for entry in registry:
        token = entry["find"]
        for paragraph in paragraphs:
            if bool(entry.get("display")):
                if paragraph_plain_text(paragraph).strip() == token.strip() and paragraph.find(".//m:oMath", NS) is None:
                    make_display_paragraph(paragraph, entry["_omath"], entry.get("number"), width)
                    counts[token] += 1
            else:
                counts[token] += replace_inline(paragraph, token, entry["_omath"])

    leftovers = []
    for entry in registry:
        token = entry["find"]
        if any(token in paragraph_plain_text(p) for p in root.findall(".//w:p", NS)):
            leftovers.append(token)

    serialized = ET.tostring(root, encoding="unicode")
    # ET may drop namespace declarations that only mc:Ignorable needs; restore the
    # original root opening tag verbatim (root attributes were not modified).
    serialized = ROOT_TAG_RE.sub(lambda _: root_tag_match.group(0), serialized, count=1)
    report = {
        "replaced": counts,
        "not_found": [token for token, n in counts.items() if n == 0],
        "still_plain_text": leftovers,
    }
    return XML_DECL + serialized, report


def rewrite_docx(src: Path, dst: Path, new_document_xml: str) -> None:
    with zipfile.ZipFile(src) as archive:
        names = archive.namelist()
        data = {name: archive.read(name) for name in names}
    data["word/document.xml"] = new_document_xml.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            archive.writestr(name, data[name])


def main() -> int:
    parser = argparse.ArgumentParser(description="Replace plain-text math in a DOCX with OMML, in place.")
    parser.add_argument("docx", nargs="?", type=Path, help="Path to the DOCX file.")
    parser.add_argument("registry", nargs="?", type=Path, help="Path to the JSON formula registry.")
    parser.add_argument("-o", "--output", type=Path, help="Output path (default: <name>.math.docx).")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input file.")
    parser.add_argument("--default-sz", type=int, default=24, help="Default half-point size (24=小四, 21=五号).")
    parser.add_argument("--convert", metavar="LATEX", help="Just convert one LaTeX fragment and print the OMML XML.")
    parser.add_argument("--display", action="store_true", help="With --convert: treat as display math.")
    args = parser.parse_args()

    try:
        if args.convert:
            (math,) = latex_to_omml([(args.convert, args.display)])
            register_namespaces(f'xmlns:m="{M_NS}" xmlns:w="{W_NS}"')
            print(ET.tostring(math, encoding="unicode"))
            return 0
        if not args.docx or not args.registry:
            parser.error("docx and registry are required (or use --convert).")
        registry = json.loads(args.registry.read_text(encoding="utf-8"))
        if not isinstance(registry, list) or not all("find" in e and ("latex" in e or "omml" in e) for e in registry):
            raise ReplaceMathError('Registry must be a JSON list of {"find": ..., "latex"|"omml": ...} entries.')
        with zipfile.ZipFile(args.docx) as archive:
            document_xml = archive.read("word/document.xml").decode("utf-8")
        new_xml, report = apply_registry(document_xml, registry, args.default_sz)
        destination = args.docx if args.in_place else (args.output or args.docx.with_suffix(".math.docx"))
        if args.in_place:
            temp = args.docx.with_suffix(".math.tmp")
            rewrite_docx(args.docx, temp, new_xml)
            temp.replace(args.docx)
        else:
            rewrite_docx(args.docx, destination, new_xml)
        print(json.dumps({"output": str(destination), **report}, ensure_ascii=False, indent=2))
        return 1 if (report["not_found"] or report["still_plain_text"]) else 0
    except (ReplaceMathError, zipfile.BadZipFile, KeyError, json.JSONDecodeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
