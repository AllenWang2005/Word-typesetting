# Word Typesetting

Allen's Codex skill for formal Chinese Word report formatting.

This repository packages a reusable Codex skill that applies a consistent Word/DOCX formatting standard for course designs, internship reports, thesis-style documents, and engineering calculation reports.

## What This Skill Covers

- Chinese body text in Songti/SimSun and English/numbers in Times New Roman, with a full font-size hierarchy (headings in Heiti, body in Songti)
- Dominant-language punctuation normalization for Chinese or English documents, including special marks (ellipsis `……`, em dash `——`, book-title `《 》`), while preserving English punctuation inside English phrases, URLs, formulas, code, decimals, and citation brackets
- Document structure and page setup: section order, A4 page margins/gutter, headers/footers, Roman-vs-Arabic page numbering via section breaks, and an auto-generated table of contents
- Left-aligned Word heading styles with multilevel numbering bound to heading styles, plus widow/orphan control
- White three-line tables (top/bottom rules ~1.5 pt, header rule ~0.75 pt), table text at small-four/12 pt and centered, captions above tables, header rows repeated across page breaks
- Figures with captions below, by-chapter numbering (`图 1-1` / `表 2-3` / `式(1-1)`), and the "reference before appearance" rule
- Numbers and units per GB 3100/3101/3102 and GB/T 15835
- Formula discovery plus LaTeX-authored formulas, inline variables, and quantity symbols rendered into native Word OMML
- Superscript, field-backed reference citations linked to bibliography entries, with bibliography entries formatted per GB/T 7714
- Appendix code blocks with red comments
- Pre-delivery checks for layout, punctuation, formulas, tables, figures, citations, and code appendices

## Repository Structure

```text
.
├── SKILL.md
├── LICENSE
├── agents/
│   └── openai.yaml
├── references/
│   ├── 20260625-formatting-standard.md
│   ├── 20260625-document-structure-and-page-setup.md
│   ├── 20260625-latex-omml-formula-workflow.md
│   ├── 20260625-reference-style-gbt7714.md
│   └── 20260625-citation-crossrefs-ooxml.md
├── scripts/
│   └── audit_docx_format.py
├── examples/
│   ├── README.md
│   ├── make_sample.py
│   └── sample-audit-output.txt
└── tests/
    └── test_audit_docx_format.py
```

Reference documents are date-prefixed (`YYYYMMDD-`) so standard revisions are easy to track. When the standard changes, add new dated reference files and update the paths in `SKILL.md`.

## Install

Install this skill from the GitHub repository:

```text
AllenWang2005/Word-typesetting
```

If installing manually, copy the repository folder into your Codex skills directory:

```text
~/.codex/skills/word-report-formatting          # macOS / Linux
%USERPROFILE%\.codex\skills\word-report-formatting   # Windows
```

## Use

Ask Codex to use the skill when creating or polishing a Word report:

```text
Use $word-report-formatting to format this course design report.
```

The skill should also trigger naturally for tasks involving formal Chinese Word reports, three-line tables, OMML formulas, figure/table captions, or appendix code formatting.

## Main Standard

The full formatting checklist lives in:

```text
references/20260625-formatting-standard.md
```

Core rule for formulas:

```text
Formulas, inline variables, mathematical objects, and quantity symbols must
be authored as LaTeX and rendered into native Word OMML. Manual italic text
is not enough. Variables are italic; digits, operators, units, function names,
constants, explanatory text, and explanatory subscripts are upright.
```

For DOCX edits, the skill can run a lightweight audit after formatting:

```text
python scripts/audit_docx_format.py path/to/report.docx
```

The script flags common failures such as punctuation mismatches, abstract/keyword indentation, centered level-1 headings, direct non-12 pt table text, likely plain-text formulas or quantity symbols, non-ASCII citation brackets, and missing `REF ref_###` citation fields. It is intentionally domain-neutral: it does not hard-code any field-specific symbol list.

## Examples

See `examples/` for a runnable script (`make_sample.py`) that builds a deliberately non-compliant DOCX with python-docx, plus the captured audit output (`sample-audit-output.txt`) showing what the auditor reports.

## Tests

Run the test suite (standard library only, no extra dependencies):

```text
python -m unittest discover -s tests -v
```

CI runs `py_compile` plus these tests on every push (see `.github/workflows/ci.yml`).

## Maintenance Notes

- Keep `SKILL.md` concise so it loads quickly when the skill triggers.
- Put detailed formatting rules in the dated `references/` files.
- Keep `scripts/audit_docx_format.py` conservative: use `FAIL` only for machine-checkable violations and `WARN` for items that require visual review.
- Do not store private information, credentials, or one-off chat history in this repository.
- When the formatting standard changes, add a new dated reference file (or update an existing one) and refresh the short summary and paths in `SKILL.md`.
