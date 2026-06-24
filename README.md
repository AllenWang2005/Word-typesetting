# Word Typesetting

Allen's Codex skill for formal Chinese Word report formatting.

This repository packages a reusable Codex skill that applies a consistent Word/DOCX formatting standard for course designs, internship reports, thesis-style documents, and engineering calculation reports.

## What This Skill Covers

- Chinese body text in Songti/SimSun and English/numbers in Times New Roman
- Dominant-language punctuation normalization for Chinese or English documents, while preserving English punctuation inside English phrases, URLs, formulas, code, decimals, and citation brackets
- Left-aligned Word heading styles and outline levels for navigation and automatic TOC
- Left-aligned abstract/keywords, black body text, and centered figure/table captions
- White three-line tables without full grids or vertical lines, with table text at small-four/12 pt and nonconforming original tables normalized
- Formula discovery plus LaTeX-authored formulas, inline variables, and quantity symbols rendered into native Word OMML
- Superscript, field-backed reference citations linked to bibliography entries
- Consistent figure/table captions and numbering
- Appendix code blocks with red comments
- Pre-delivery checks for layout, punctuation, formulas, tables, figures, citations, and code appendices

## Repository Structure

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
    ├── formatting-standard.md
    ├── latex-omml-formula-workflow.md
    └── citation-crossrefs-ooxml.md
└── scripts/
    └── audit_docx_format.py
```

## Install

Install this skill from the GitHub repository:

```text
AllenWang2005/Word-typesetting
```

If installing manually, copy the repository folder into your Codex skills directory, for example:

```text
C:\Users\25102\.codex\skills\word-report-formatting
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
references/formatting-standard.md
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

The script flags common failures such as punctuation mismatches, abstract/keyword indentation, centered level-1 headings, direct non-12 pt table text, likely plain-text formulas or quantity symbols, non-ASCII citation brackets, and missing `REF ref_###` citation fields.

## Maintenance Notes

- Keep `SKILL.md` concise so it loads quickly when the skill triggers.
- Put detailed formatting rules in `references/formatting-standard.md`.
- Keep `scripts/audit_docx_format.py` conservative: use `FAIL` only for machine-checkable violations and `WARN` for items that require visual review.
- Do not store private information, credentials, or one-off chat history in this repository.
- When the formatting standard changes, update both `references/formatting-standard.md` and the short summary in `SKILL.md` if needed.
