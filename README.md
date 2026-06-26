# Word Typesetting

**English** · [简体中文](README.zh-CN.md)

[![Release](https://img.shields.io/github/v/release/AllenWang2005/Word-typesetting?label=release)](https://github.com/AllenWang2005/Word-typesetting/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![CI](https://github.com/AllenWang2005/Word-typesetting/actions/workflows/ci.yml/badge.svg)](https://github.com/AllenWang2005/Word-typesetting/actions/workflows/ci.yml)
[![Stars](https://img.shields.io/github/stars/AllenWang2005/Word-typesetting?style=flat)](https://github.com/AllenWang2005/Word-typesetting/stargazers)
![Skill](https://img.shields.io/badge/skill-Codex%20%7C%20Claude%20Code-7c3aed)

A reusable **Codex / Claude Code skill** that formats formal Chinese Word (`.docx`)
reports to a consistent academic standard — course designs, internship reports,
thesis-style documents, and engineering calculation reports.

It is two things working together:

1. **A written standard** (`SKILL.md` + `references/`) that tells the model exactly
   how a compliant report should look — typography, document structure, tables,
   figures, formulas, citations, numbers/units.
2. **Tooling** (`scripts/`) that *checks* a finished `.docx` against the machine-checkable
   parts of that standard, and *auto-fixes* the safest issues.

The model does the actual editing (with the general DOCX tooling); this skill supplies
the rules, a completion checklist, and a guardrail so nothing obvious slips through.

---

## What it can do

**Typography & language**
- Chinese body text in Songti/SimSun, English/numbers in Times New Roman, level 1–2 headings in Heiti (not bold), level-3 headings in bold Songti.
- A full font-size hierarchy — see [the table below](#font-size-hierarchy).
- Dominant-language punctuation normalization (Chinese vs. English), special marks (`……`, `——`, `《 》`), while protecting URLs, DOIs, code, formulas, decimals, and citation brackets.
- Body justified, first-line indent 2 字符, 1.5× line spacing.

**Document structure & page setup**
- Standard section order (cover → declarations → abstracts → TOC → body → references → appendix → acknowledgements).
- A4 page setup, margins/gutter, headers/footers, Roman→Arabic page numbering via section breaks.
- Auto-generated table of contents (with optional figure/table lists), multilevel heading numbering (`1 / 1.1 / 1.1.1`), widow/orphan control.
- Headings left-aligned (title centered), no trailing punctuation, depth ≤ 3–4 levels.

**Tables, figures, formulas**
- White three-line tables (top/bottom rules ~1.5 pt, header rule ~0.75 pt), content 五号/10.5 pt (one size below body) and centered, **caption above the table**, header row repeated across page breaks.
- Figures with **caption below**, by-chapter numbering (`图 1-1` / `表 2-3` / `(3-1)`), the "reference-before-appearance" rule.
- Formulas authored as LaTeX and rendered to **native Word OMML** — variables italic, units/operators/functions upright; equation numbers right-aligned per chapter, cited as "由式 (3-1) 可得".

**Citations & references**
- In-text citations as superscript, **field-backed cross-references to the whole bracketed `[1]`** (brackets included, not a bare number).
- Bibliography entries per **GB/T 7714—2015** (numeric sequential-coding by default; author-year supported), with the right type tags (`[J] [M] [D] [C] [S] [P] [EB/OL]` …).

**Numbers, units, appendix code**
- Half-width space between number and unit (`20 m³/s`) — except `%`, `°`, `℃`/`°C`, which attach directly (`50%`, `30°`, `25℃`); GB 3100/3101/3102 units, GB/T 15835 numeral usage.
- Appendix code after the main text with **red comments** (`C00000`).

## Font-size hierarchy

Defaults below; a school/journal template always takes precedence.

| Element | Size | Font / style |
| --- | --- | --- |
| Cover title | 二号 (22 pt) | Heiti, bold, centered |
| Section titles (abstract / TOC / references) | 小二 (18 pt) | Heiti, not bold |
| Heading 1 (chapter) | 三号 (16 pt) | Heiti, not bold; 0.5-line space before/after |
| Heading 2 (section) | 四号 (14 pt) | Heiti, not bold |
| Heading 3 (subsection) | 小四 (12 pt) | Songti, **bold** |
| Body | 小四 (12 pt) | Songti (CJK) / Times New Roman (Latin & digits) |
| Abstract / keywords | 小四 (12 pt) | Songti |
| Figure / table captions | 五号 (10.5 pt) | Songti (CJK) / Times New Roman (Latin), centered |
| Table content | 五号 (10.5 pt) | Songti, one size below body |
| Table notes | 小五 (9 pt) | Songti |
| Reference entries | 五号 (10.5 pt) | Songti, hanging indent |
| Header / footer | 小五 (9 pt) | Songti |
| Footnotes | 小五 (9 pt) | Songti |

## What the result looks like

A polished, navigable report: a clickable heading outline and an auto-updating TOC;
left-aligned Heiti headings over Songti body text at a consistent size; clean white
three-line tables with captions above and figures with captions below, all numbered
by chapter; real Word equation objects instead of pasted images or plain text;
superscript `[1]`-style citations that renumber themselves because they are Word
cross-reference fields; and a GB/T 7714—2015 reference list. Running the auditor on a
finished file prints `PASS: no machine-detected guardrail issues.`

See [`examples/`](examples/) for a deliberately *non-compliant* sample and the audit
report it produces — a concrete before/after of what the rules catch.

## Install

Install from GitHub:

```text
AllenWang2005/Word-typesetting
```

Or copy the folder into your skills directory:

```text
~/.codex/skills/word-report-formatting                 # Codex, macOS / Linux
%USERPROFILE%\.codex\skills\word-report-formatting     # Codex, Windows
~/.claude/skills/word-report-formatting                # Claude Code
```

## Use

Ask the assistant to use the skill when creating or polishing a Word report:

```text
Use $word-report-formatting to format this course design report.
```

It also triggers naturally for formal Chinese Word reports, three-line tables, OMML
formulas, figure/table captions, citation cross-references, or appendix code.

## The audit script

After formatting a DOCX, run the guardrail:

```text
python scripts/audit_docx_format.py path/to/report.docx
python scripts/audit_docx_format.py path/to/report.docx --json   # machine-readable
```

It reports `FAIL` (machine-certain violations) and `WARN` (needs a human look), ends
with a `FAIL/WARN` summary, and notes any per-code truncation. Checks include:

| Code | Severity | What it catches |
| --- | --- | --- |
| `ZH_PUNCT` / `EN_PUNCT` | FAIL | Wrong-language punctuation in prose |
| `ABSTRACT_INDENT` | FAIL | Abstract/keywords centered or indented |
| `H1_CENTER` | FAIL | Level-1 heading centered instead of left-aligned |
| `H1_GAP` | WARN | Level-1 heading missing a blank gap before it |
| `HEADING_PUNCT` | WARN | Heading ends with punctuation |
| `HEADING_FONT` / `BODY_FONT` | WARN | Heading in Songti / body in Heiti (direct fonts only) |
| `HEADING_NO_STYLE` | WARN | Heading-looking line that uses no Word heading style |
| `TABLE_SIZE` | FAIL | Table text not 五号/10.5 pt (小四 ok; never larger than body) |
| `TABLE_BORDERS` | WARN | Table has vertical/inner borders (not a three-line table) |
| `CAPTION_POSITION` | WARN | Table caption below table / figure caption above figure |
| `COLOR` | WARN | Stray non-black font color (hyperlinks/theme colors excluded) |
| `CITATION_BRACKETS` | FAIL | Full-width citation brackets `［1］` |
| `CITATION_FIELDS` | FAIL | Citations exist but no `REF ref_###` fields |
| `CITATION_NO_BRACKETS` | WARN | Bare superscript number citation missing its brackets |
| `VISIBLE_LATEX` | FAIL | Visible LaTeX source instead of OMML |
| `FORMULA_TEXT` | FAIL/WARN | Plain-text formula / quantity symbol that should be OMML |

It is intentionally **domain-neutral** — no hard-coded field-specific symbol list.

**Audit scope:** the script inspects the main document story in `word/document.xml`.
It does not audit headers, footers, footnotes, endnotes, comments, or embedded parts —
verify those visually or with a deeper OOXML pass when they matter. Table-cell text is
checked for font size and borders only, not punctuation/heading/formula (this avoids
false positives from numbers and short fragments in cells).

## The auto-fix script

`scripts/normalize_docx.py` mechanically fixes the two safest issues — full-width
citation brackets (`［1］` → `[1]`) and ASCII sentence punctuation between CJK
characters (`中文,中文` → `中文，中文`) — preserving every other byte of the package:

```text
python scripts/normalize_docx.py report.docx -o report.fixed.docx
python scripts/normalize_docx.py report.docx --in-place
```

It deliberately does **not** touch fonts, styles, formulas, or cross-references; those
need judgement and stay with the model + the main standard.

## Tests

Standard library only (no third-party dependencies):

```text
python -m unittest discover -s tests -v
```

CI runs `py_compile` plus the test suite on Python 3.9 and 3.12 on every push
(see `.github/workflows/ci.yml`).

## Repository structure

```text
.
├── SKILL.md
├── LICENSE
├── CHANGELOG.md
├── agents/
│   └── openai.yaml
├── references/
│   ├── formatting-standard.md                  # main checklist
│   ├── document-structure-and-page-setup.md    # structure, page setup, TOC, font sizes
│   ├── latex-omml-formula-workflow.md          # LaTeX → Word OMML workflow
│   ├── reference-style-gbt7714.md              # GB/T 7714—2015 bibliography format
│   └── citation-crossrefs-ooxml.md             # in-text REF cross-reference OOXML
├── scripts/
│   ├── audit_docx_format.py                    # read-only guardrail
│   └── normalize_docx.py                       # safe auto-fixer
├── examples/
│   ├── README.md
│   ├── make_sample.py                          # builds compliant / non-compliant samples
│   ├── sample-audit-output.txt
│   └── sample-compliant-audit-output.txt
└── tests/
    ├── test_audit_docx_format.py
    └── test_normalize_docx.py
```

See [`CHANGELOG.md`](CHANGELOG.md) for the version history.

## Maintenance notes

- Keep `SKILL.md` concise so it loads quickly when the skill triggers.
- Put detailed rules in `references/`.
- Keep the auditor conservative: `FAIL` only for machine-checkable violations, `WARN` for items needing visual review.
- Do not store private information, credentials, or one-off chat history in this repository.
- When the standard changes, update the relevant reference file and refresh the summary/paths in `SKILL.md`.

## License

[MIT](LICENSE).
