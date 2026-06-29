---
name: word-report-formatting
description: Apply Allen's formal Word report formatting standard. Use when creating, editing, reviewing, or finalizing Word/DOCX reports, course designs, internship reports, thesis-style documents, scientific/engineering reports, LaTeX-to-Word OMML formulas, inline variables, quantity symbols, three-line tables, figure/table captions, reference/citation cross-references, or appendices with code comments.
---

# Word Report Formatting

## Overview

Use this skill to format formal Chinese Word reports consistently with Allen's preferred standard. It complements the general DOCX/document skill: use the document tooling for implementation and rendering, and use this skill for the required typography, formula, table, figure, reference/citation, and appendix-code rules.

**Standard version: 2026-06-29.** Detailed rules live in `references/`:

- `references/formatting-standard.md` — main formatting checklist (typography, language/punctuation, headings, tables, formulas, figures, numbers/units, citations, lists/footnotes, appendix code).
- `references/document-structure-and-page-setup.md` — document structure and order, page setup, headers/footers, page numbering and sections, TOC, font-size hierarchy, multilevel numbering, widow/orphan control.
- `references/latex-omml-formula-workflow.md` — LaTeX→OMML formula discovery and conversion.
- `references/reference-style-gbt7714.md` — GB/T 7714 bibliography entry format.
- `references/citation-crossrefs-ooxml.md` — in-text superscript citations + Word REF cross-reference OOXML.
- `references/three-line-table-ooxml.md` — OOXML recipe for the three-line table (which borders to set, line widths, avoiding gridded styles).

## Required Workflow

1. Read `references/formatting-standard.md` before creating or substantially editing a Word report; also read `references/document-structure-and-page-setup.md` when the document needs structure, page setup, page numbering, a table of contents, or the font-size hierarchy.
2. Determine the dominant document language before punctuation cleanup. Chinese documents use Chinese punctuation in Chinese prose; English documents use ASCII punctuation in prose. Preserve English punctuation inside English phrases/quotes and protected tokens such as URLs, DOIs, code, formulas, file paths, decimals, and citation brackets.
3. Apply the standard unless the user, school template, or provided rubric explicitly overrides it.
4. Preserve existing document content and structure during edit tasks; make only the local formatting changes needed.
5. For any document containing formulas, variables, quantity symbols, subscripts/superscripts, units, or math-like expressions, read `references/latex-omml-formula-workflow.md` and run a formula discovery pass before styling. This includes bare one-letter quantity symbols in definition/explanation prose, such as `式中 Q 为流量，N 为出力`.
6. Write formulas, inline variables, math objects, and quantity symbols as LaTeX first, then render them into native Word OMML equations. Do not satisfy this standard by manually italicizing normal text.
7. For paper-style citations, read `references/citation-crossrefs-ooxml.md` and convert body citations such as `[1]`, `[1][2]`, `[1,2]`, or `[1-3]` into superscript Word cross-references to the matching bibliography entries; format the bibliography entries themselves per `references/reference-style-gbt7714.md` (GB/T 7714).
8. If editing an existing DOCX and Python is available, run `scripts/audit_docx_format.py <path-to-docx>` after formatting. Fix `FAIL` items and inspect `WARN` items; this script is a guardrail, not a substitute for visual QA. For the two safest mechanical fixes (full-width citation brackets and ASCII punctuation between CJK characters) you may run `scripts/normalize_docx.py <path-to-docx>`.
9. Before delivery, run a formatting audit against the checklist in `references/formatting-standard.md`.
10. Treat the Must-Fix Audit below as a completion gate: do not deliver or claim compliance while any must-fix item remains, and disclose any item that could not be verified in the current environment.
11. When possible, render/export the DOCX to PDF or page images and inspect pages for overlap, clipping, table overflow, font substitution, broken cross-references, formula rendering errors, and appendix code readability.

## Core Rules

- Body text: Chinese in Songti/SimSun, English and numbers in Times New Roman, default small-four/12 pt.
- Headings: real Word heading styles with outline levels — level 1–2 headings in Heiti (not bold), level-3 headings in bold Songti, body in Songti/SimSun; left aligned rather than centered. Use Word multilevel lists bound to heading styles for automatic `1 / 1.1 / 1.1.1` numbering, kept to about three or four levels. Headings carry no trailing punctuation; the document title is centered while section headings are left aligned. Level-1 Chinese headings such as `一、` / `二、` / `三、` need about half a line of space before them and must not sit alone at the bottom of a page.
- Body punctuation: match the dominant document language. Chinese prose uses Chinese punctuation; English prose uses ASCII punctuation. Keep citation brackets ASCII `[` and `]`. Use the Chinese ellipsis `……`, em dash `——`, and book-title marks `《 》` where applicable.
- Body color: keep body text, headings, captions, tables, formulas, abstract, keywords, and references black unless the user explicitly requests otherwise. Appendix code comments are the standard red exception.
- Abstract and keywords: left aligned with left indent 0 and first-line indent 0.
- Tables: white three-line tables — the top and bottom rules visible and **thicker** (~1.5 pt, `w:sz≈12`), a single **thinner** header rule in the middle (~0.75 pt, `w:sz≈6`), and every other border (verticals, inside horizontals) set to none. Do not use a gridded or header-shaded table style; set the borders explicitly (see `references/three-line-table-ooxml.md`). A "no top/bottom rule, thick middle line" result is wrong. Table text — including in-table formulas and symbols — is 五号/10.5 pt (one size below the body; 小四 also accepted, never larger than body) and centered by default; units go in the header. The table caption goes **above** the table, centered. Repeat the header row across page breaks; do not split a table mid-row. Normalize existing tables that do not meet the standard.
- Figures: the figure caption goes **below** the figure, centered; keep plot titles out of the image when a Word caption is present.
- Numbering: number figures, tables, and equations by chapter (`图 1-1`, `表 2-3`, `式(1-1)`); refer to them in text as “如图 1-1 所示” / “式(1-1)”. Figures and tables must appear **after** they are first referenced (proximity rule).
- Formulas and symbols: author every formula, inline variable, math object, and quantity symbol in LaTeX and convert it to native Word OMML, including bare symbols such as `Q`, `N`, `H`, or `V` when they define or denote physical quantities. Variables are italic; digits, operators, units, function names, constants, explanatory text, and explanatory subscripts are upright — do **not** blanket-italicize the whole equation (that slants the digits too). Multi-letter coefficients use one variable + upright subscript (`C_I`, not `CI`). Put a display equation on a left-aligned paragraph with a center tab (equation) and a right tab (number) so the number is right-aligned; do not center the whole line. Formula font size follows its context — body formulas at the body size (小四), in-table formulas at the table size (五号).
- Numbers and units: put a half-width space between a number and its unit (`20 m³/s`), except `%`, `°`, and `℃`/`°C`, which attach directly (`50%`, `30°`, `25℃`); follow GB 3100/3101/3102 for units and GB/T 15835 for Arabic-vs-Chinese numeral usage.
- Citations: make body references superscript and field-backed, not merely static superscript text. Reference the whole bracketed `[1]` as one cross-reference unit (brackets included), not a bare number. Use ASCII square brackets `[` and `]`, not Chinese/full-width brackets. Format bibliography entries per GB/T 7714—2015.
- Appendix code: put code after main text and tables; English/numbers in Times New Roman 小四 (12 pt), Chinese in Songti 小四; color code comments red, preferably `C00000`.

## Enforcement Protocol

Use this skill as a completion checklist, not background guidance. Work in passes: language/punctuation, styles and alignment, color/fonts, tables/captions, formula discovery plus LaTeX-to-OMML rendering, citations/cross-references, appendix code, then script/render audit. If a pass cannot be verified, say so instead of implying the document is fully compliant.

## Must-Fix Audit

Fail the formatting pass and revise if any of these remain:

- Section headings are centered when they should be left aligned (the document title is the centered exception), level 1–2 headings are not in Heiti or are bold (level-3 headings are bold Songti), or a heading ends with trailing punctuation.
- The dominant-language punctuation pass was skipped or left obvious mismatches in prose.
- Chinese documents still contain ASCII commas, periods, colons, semicolons, question marks, exclamation marks, or quote marks in Chinese prose outside protected English/code/formula contexts; English documents still contain Chinese punctuation outside quoted Chinese source text.
- Abstract or keywords are indented, centered, or have nonzero left/first-line indent.
- Body text, headings, captions, table text, or reference text contain unintended non-black color.
- A table caption is not above its table, or a figure caption is not below its figure, or figure/table names are not centered.
- Figures, tables, or equations are not numbered by chapter, or a figure/table appears before it is first referenced in the text.
- Level-1 Chinese headings do not have a blank-line visual gap before them when they follow body text, or a heading is stranded at the bottom of a page.
- Table text is larger than the body or not the expected 五号/10.5 pt (小四/12 pt also accepted), unless the user or official template explicitly requires a different size.
- Formulas, inline variables, quantity symbols, subscripts/superscripts, or math objects remain as ordinary styled text instead of LaTeX-rendered Word OMML, including bare one-letter quantity symbols in `式中` / `其中` / `表示` / `为` definition contexts.
- A document with formulas/symbols was edited without a formula discovery pass and a LaTeX source list/registry.
- A formula is blanket-italicized (digits/operators slanted), or a numbered display equation is centered instead of having a right-aligned number, or appendix code is not Times New Roman 小四.
- Existing tables, formulas, variables, units, or symbols still violate the standard.
- Body citations are only static superscript text instead of Word `REF` fields/bookmark cross-references, or a citation references only the bare number instead of the whole bracketed `[1]`, unless the user explicitly allows visual-only citations; or bibliography entries do not follow GB/T 7714—2015.
- The available DOCX audit script reports `FAIL` items that have not been fixed or explicitly explained.

## Conflict Handling

If older project memory says "formulas should not be italic", treat that as superseded for formal report formulas. The current rule is: variables italic, non-variables upright, following the provided teacher screenshot/specification.
