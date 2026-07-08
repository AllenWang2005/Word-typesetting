---
name: word-report-formatting
description: Apply Allen's formal Word report formatting standard. Use when creating, editing, reviewing, or finalizing Word/DOCX reports, course designs, internship reports, thesis-style documents, scientific/engineering reports, LaTeX-to-Word OMML formulas, inline variables, quantity symbols, three-line tables, figure/table captions, reference/citation cross-references, or appendices with code comments.
---

# Word Report Formatting

## Overview

Use this skill to format formal Chinese Word reports consistently with Allen's preferred standard. It complements the general DOCX/document skill: use the document tooling for implementation and rendering, and use this skill for the required typography, formula, table, figure, reference/citation, and appendix-code rules.

**Standard version: 2026-07-08 (v1.6.2).** Acceptance decisions: the standard follows GB norms (GB 3100/3101/3102, GB/T 15835, GB/T 7714) and applies to general formal documents; **MS Word is the canonical renderer** — when renderers disagree, what Word shows is what counts. The audit script is the primary completion gate; do **not** run PDF/page-image render passes unless the user explicitly asks (token cost) — a quick visual open in Word by the user replaces them.

Detailed rules live in `references/`:

- `references/formatting-standard.md` — main formatting checklist (typography, language/punctuation, headings, tables, formulas, figures, numbers/units, citations, lists/footnotes, appendix code).
- `references/document-structure-and-page-setup.md` — document structure and order, page setup, headers/footers, page numbering and sections, TOC, font-size hierarchy, multilevel numbering, widow/orphan control.
- `references/latex-omml-formula-workflow.md` — LaTeX→OMML formula discovery and conversion.
- `references/reference-style-gbt7714.md` — GB/T 7714 bibliography entry format.
- `references/citation-crossrefs-ooxml.md` — in-text superscript citations + Word REF cross-reference OOXML.
- `references/three-line-table-ooxml.md` — OOXML recipe for the three-line table (which borders to set, line widths, avoiding gridded styles).
- `references/generate-from-source.md` — source-first authoring for new documents (Markdown + LaTeX compiled with Pandoc; harvesting existing formulas back to LaTeX).

## Required Workflow

1. Read `references/formatting-standard.md` before creating or substantially editing a Word report; also read `references/document-structure-and-page-setup.md` when the document needs structure, page setup, page numbering, a table of contents, or the font-size hierarchy. For any document that contains or will contain tables, read `references/three-line-table-ooxml.md` before building or normalizing the tables — it specifies which borders to set and how to clear shading; do not build tables from memory of "三线表".
2. Determine the dominant document language before punctuation cleanup. Chinese documents use Chinese punctuation in Chinese prose; English documents use ASCII punctuation in prose. Preserve English punctuation inside English phrases/quotes and protected tokens such as URLs, DOIs, code, formulas, file paths, decimals, and citation brackets.
3. Apply the standard unless the user, school template, or provided rubric explicitly overrides it.
4. Preserve existing document content and structure during edit tasks; make only the local formatting changes needed.
5. When fixing headings, set the Word heading styles themselves in `word/styles.xml` **and** repair existing heading paragraph runs when needed. Do not rely on style names alone: localized Word/WPS files may use style IDs such as `1`, `21`, and `31`, and missing `w:rFonts` lets WPS inherit the wrong font. Level 1/2 heading styles must explicitly use Heiti and not bold; level 3 must explicitly use bold Songti.
6. For any document containing formulas, variables, quantity symbols, subscripts/superscripts, units, or math-like expressions, read `references/latex-omml-formula-workflow.md` and run a formula discovery pass before styling. This includes bare one-letter quantity symbols in definition/explanation prose, such as `式中 Q 为流量，N 为出力`.
7. Write formulas, inline variables, math objects, and quantity symbols as LaTeX first, then render them into native Word OMML equations — preferably by building a JSON registry and running `scripts/replace_math.py` (deterministic Pandoc conversion + exact-position splice; its summary's `not_found`/`still_plain_text` must be empty). Every conversion is an **in-place replacement**: the OMML object goes exactly where the original token was and the original plain text is removed; never append rendered math at the end of a paragraph or merge adjacent expressions. Do not satisfy this standard by manually italicizing normal text. For brand-new documents prefer source-first authoring per `references/generate-from-source.md`.
8. For paper-style citations, read `references/citation-crossrefs-ooxml.md` and convert body citations such as `[1]`, `[1][2]`, `[1,2]`, or `[1-3]` into superscript Word cross-references to the matching bibliography entries; format the bibliography entries themselves per `references/reference-style-gbt7714.md` (GB/T 7714).
9. **Mandatory delivery gate**: run `python scripts/finalize_docx.py <path-to-docx>` after formatting (it applies every safe mechanical fix — citation brackets, CJK punctuation, number-unit spacing, in-table shading, `w:tblLook`, header-row repeat, `w:updateFields` — then runs the full audit and prints a verdict). `DELIVERY GATE: FAIL` means the document must not be delivered; fix every `FAIL` and re-run until it passes, and copy the gate verdict plus the audit summary line into the delivery note. Every Must-Fix item is enforced at `FAIL` level — a `WARN`-only report passes the gate, so also read the WARN lines. If Python is unavailable, say so explicitly in the delivery note instead of implying the gate ran.
10. Before delivery, run a formatting audit against the checklist in `references/formatting-standard.md`.
11. Treat the Must-Fix Audit below as a completion gate: do not deliver or claim compliance while any must-fix item remains, and disclose any item that could not be verified in the current environment.
12. Do not run PDF/page-image render inspections unless the user explicitly asks for one. The audit script plus the Must-Fix gate are the completion criteria; recommend that the user opens the delivered file once in MS Word (the canonical renderer) to spot-check layout, fonts, and formulas. WPS/Word view aids such as formatting marks, object anchors, table gridlines, and selection handles are UI overlays, not printable document content; do not "fix" them by deleting content.

## Core Rules

- Body text: Chinese in Songti/SimSun, English and numbers in Times New Roman, default small-four/12 pt.
- Headings: real Word heading styles with outline levels — level 1–2 headings in Heiti (not bold), level-3 headings in bold Songti, body in Songti/SimSun; left aligned rather than centered. Set explicit `w:rFonts` on the heading styles in `styles.xml` (Chinese/eastAsia font plus Times New Roman for Latin/numbers) and, when repairing an existing document, direct-format existing heading runs if WPS is inheriting the wrong font. Use Word multilevel lists bound to heading styles for automatic `1 / 1.1 / 1.1.1` numbering, kept to about three or four levels. Headings carry no trailing punctuation; the document title is centered while section headings are left aligned. Level-1 Chinese headings such as `一、` / `二、` / `三、` need about half a line of space before them and must not sit alone at the bottom of a page.
- Body punctuation: match the dominant document language. Chinese prose uses Chinese punctuation; English prose uses ASCII punctuation. Keep citation brackets ASCII `[` and `]`. Use the Chinese ellipsis `……`, em dash `——`, and book-title marks `《 》` where applicable.
- Body color: keep body text, headings, captions, tables, formulas, abstract, keywords, and references black unless the user explicitly requests otherwise. Appendix code comments are the standard red exception.
- Abstract and keywords: left aligned with left indent 0 and first-line indent 0.
- Tables: white three-line tables — the top and bottom rules visible and **thicker** (~1.5 pt, `w:sz≈12`), a single **thinner** header rule in the middle (~0.75 pt, `w:sz≈6`), and every other border (verticals, inside horizontals) set to none. **Every cell must be shading-free (white)**: remove the table style reference (or use a plain style), zero the `w:tblLook` conditional-formatting flags, and clear every `w:shd` — header shading usually comes from the table style's `firstRow` conditional format, not from a manual fill. Do not use a gridded or header-shaded table style; set the borders explicitly (see `references/three-line-table-ooxml.md`). A "no top/bottom rule, thick middle line" result is wrong. Table text — including in-table formulas and symbols — is 五号/10.5 pt (one size below the body; 小四 also accepted, never larger than body) and centered by default; units go in the header. The table caption goes **above** the table, centered. Repeat the header row across page breaks; do not split a table mid-row. Normalize existing tables that do not meet the standard.
- Figures: the figure caption goes **below** the figure, centered; keep plot titles out of the image when a Word caption is present.
- Numbering: number figures, tables, and equations by chapter (`图 1-1`, `表 2-3`, `式(1-1)`); refer to them in text as “如图 1-1 所示” / “式(1-1)”. Figures and tables must appear **after** they are first referenced (proximity rule).
- Formulas and symbols: author every formula, inline variable, math object, and quantity symbol in LaTeX and convert it to native Word OMML, including bare symbols such as `Q`, `N`, `H`, or `V` when they define or denote physical quantities. Variables are italic; digits, operators, units, function names, constants, explanatory text, and explanatory subscripts are upright — do **not** blanket-italicize the whole equation (that slants the digits too). Beware the OMML default: *letters render math-italic unless explicitly marked upright* (`m:sty val="p"`, produced by LaTeX `\mathrm`/`\text`), so a unit like `km` left unmarked is italic even with no italic tag in the XML; and never fake a formula with an italicized text run. Multi-letter coefficients use one variable + upright subscript (`C_I`, not `CI`). Put a display equation on a left-aligned paragraph with a center tab (equation) and a right tab (number) so the number is right-aligned; do not center the whole line. If WPS/renderer tab-stop behavior is unreliable, use a borderless one-row 1x3 formula layout table (empty left cell, centered equation cell, right-aligned number cell) with equal left/right spacer widths; this is a formula layout device, not a three-line data table. Formula font size follows its context — body formulas at the body size (小四), in-table formulas at the table size (五号).
- Numbers and units: put a half-width space between a number and its unit (`20 m³/s`), except `%`, `°`, and `℃`/`°C`, which attach directly (`50%`, `30°`, `25℃`); follow GB 3100/3101/3102 for units and GB/T 15835 for Arabic-vs-Chinese numeral usage.
- Citations: make body references superscript and field-backed, not merely static superscript text. Reference the whole bracketed `[1]` as one cross-reference unit (brackets included), not a bare number. Use ASCII square brackets `[` and `]`, not Chinese/full-width brackets. Format bibliography entries per GB/T 7714—2015.
- Appendix code: put code after main text and tables; English/numbers in Times New Roman 小四 (12 pt), Chinese in Songti 小四; color code comments red, preferably `C00000`.

## Enforcement Protocol

Use this skill as a completion checklist, not background guidance. Work in passes: language/punctuation, styles and alignment, color/fonts, tables/captions, formula discovery plus LaTeX-to-OMML rendering, citations/cross-references, appendix code, then script/render audit. If a pass cannot be verified, say so instead of implying the document is fully compliant.

## Must-Fix Audit

Fail the formatting pass and revise if any of these remain:

- Section headings are centered when they should be left aligned (the document title is the centered exception), level 1–2 headings are not in Heiti or are bold (level-3 headings are bold Songti), the used heading styles in `styles.xml` lack explicit heading fonts, or a heading ends with trailing punctuation.
- The dominant-language punctuation pass was skipped or left obvious mismatches in prose.
- Chinese documents still contain ASCII commas, periods, colons, semicolons, question marks, exclamation marks, or quote marks in Chinese prose outside protected English/code/formula contexts; English documents still contain Chinese punctuation outside quoted Chinese source text.
- Abstract or keywords are indented, centered, or have nonzero left/first-line indent.
- Body text, headings, captions, table text, or reference text contain unintended non-black color.
- A table caption is not above its table, or a figure caption is not below its figure, or figure/table names are not centered.
- Figures, tables, or equations are not numbered by chapter, or a figure/table appears before it is first referenced in the text.
- A display equation has no chapter number `(3-1)`, or a numbered equation is never cited in prose (write “由式 (3-1) 可得” / “按式 (3-1) 计算”).
- Level-1 Chinese headings do not have a blank-line visual gap before them when they follow body text, or a heading is stranded at the bottom of a page.
- Table text is larger than the body or not the expected 五号/10.5 pt (小四/12 pt also accepted), unless the user or official template explicitly requires a different size.
- A table has any cell shading / non-white fill (including header shading inherited from a table style's `firstRow` conditional format), lacks visible thick top/bottom rules with a single thinner header rule, or retains body-row cell borders that make it render as a grid.
- Rendered math was appended instead of replacing the original text in place: a paragraph contains both an OMML object and the same expression as plain text, a trailing cluster of formulas repeats values from the prose, adjacent expressions were merged into one object, or stray fragments (`mm，Cv`) remain at a paragraph end.
- Formulas, inline variables, quantity symbols, subscripts/superscripts, or math objects remain as ordinary styled text instead of LaTeX-rendered Word OMML, including bare one-letter quantity symbols in `式中` / `其中` / `表示` / `为` definition contexts.
- A document with formulas/symbols was edited without a formula discovery pass and a LaTeX source list/registry.
- A formula is blanket-italicized (digits/operators slanted), or a numbered display equation is centered instead of having a right-aligned number, or appendix code is not Times New Roman 小四.
- Existing tables, formulas, variables, units, or symbols still violate the standard.
- Body citations are only static superscript text instead of Word `REF` fields/bookmark cross-references, or a citation references only the bare number instead of the whole bracketed `[1]`, unless the user explicitly allows visual-only citations; or bibliography entries do not follow GB/T 7714—2015.
- `scripts/finalize_docx.py` was not run, or its delivery gate reports `FAIL` items that have not been fixed or explicitly explained.

## Conflict Handling

If older project memory says "formulas should not be italic", treat that as superseded for formal report formulas. The current rule is: variables italic, non-variables upright, following the provided teacher screenshot/specification.
