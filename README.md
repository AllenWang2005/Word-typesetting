# Word Typesetting

Allen's Codex skill for formal Chinese Word report formatting.

This repository packages a reusable Codex skill that applies a consistent Word/DOCX formatting standard for course designs, internship reports, thesis-style documents, and engineering calculation reports.

## What This Skill Covers

- Chinese body text in Songti/SimSun and English/numbers in Times New Roman
- Left-aligned Word heading styles and outline levels for navigation and automatic TOC
- Left-aligned abstract/keywords, black body text, and centered figure/table captions
- White three-line tables without full grids or vertical lines, with nonconforming original tables normalized
- Native Word OMML formulas with correct italic/upright rules
- Superscript, field-backed reference citations linked to bibliography entries
- Consistent figure/table captions and numbering
- Appendix code blocks with red comments
- Pre-delivery checks for layout, formulas, tables, figures, and code appendices

## Repository Structure

```text
.
├── SKILL.md
├── agents/
│   └── openai.yaml
└── references/
    ├── formatting-standard.md
    └── citation-crossrefs-ooxml.md
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
Variables are italic. Digits, operators, units, function names, constants,
explanatory text, and explanatory subscripts are upright.
```

## Maintenance Notes

- Keep `SKILL.md` concise so it loads quickly when the skill triggers.
- Put detailed formatting rules in `references/formatting-standard.md`.
- Do not store private information, credentials, or one-off chat history in this repository.
- When the formatting standard changes, update both `references/formatting-standard.md` and the short summary in `SKILL.md` if needed.
