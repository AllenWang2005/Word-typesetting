# Changelog

All notable changes to this skill are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/), and the project uses semantic versioning.

## [1.6.3] - 2026-07-09

### Fixed
- Auditor: in-table OMML formulas and quantity symbols now have their own
  `TABLE_FORMULA_SIZE` FAIL when they are missing an explicit size or still use
  the body formula size. Normal data-table formulas must be 五号/10.5 pt
  (`w:sz=21`, `w:szCs=21`), while body display formulas remain 小四/12 pt.

### Changed
- `replace_math.py` stamps formula size by replacement context when the registry
  does not override `sz`: body formulas use 小四, and formulas replaced inside
  table cells use 五号.
- `normalize_docx.py --tables` and `finalize_docx.py` now mechanically fix
  existing OMML formula runs inside normal data tables to 五号, while skipping
  borderless 1-row formula layout tables so centered display equations keep
  their body formula size.
- Skill/reference/README guidance now separates ordinary table text size from
  table formula size so the audit cannot pass a visually oversized table formula.

## [1.6.2] - 2026-07-08

### Fixed
- Auditor: heading font checks now resolve heading styles from `styles.xml`
  and localized Word/WPS built-in style IDs such as `1`, `21`, and `31`.
  Used heading styles without explicit Heiti/Songti fonts now fail the gate
  (`HEADING_STYLE_FONT`), and wrong heading boldness fails as `HEADING_BOLD`.
- Auditor: borderless one-row 1×3 formula layout tables are recognized as
  equation-layout helpers, not data tables, so they are skipped by three-line
  table, table-header-repeat, and table-formula-text checks.

### Changed
- Skill/reference docs now require heading fonts to be written into the actual
  Word heading styles, not just visually applied to some runs; this prevents WPS
  from inheriting the wrong title font.
- Formula layout guidance now explicitly recommends the 1×3 layout-table fallback
  when tab stops drift in WPS while preserving "equation centered, number right".
- The checklist now distinguishes WPS/Word UI overlays (formatting marks, object
  anchors, table gridlines, selection handles) from printable document content.

## [1.6.1] - 2026-07-08

### Fixed
- Auditor: `TABLE_BORDERS` / `TABLE_RULES` now inspect row-level table property
  exceptions (`w:tblPrEx/w:tblBorders`). This closes the real-world failure where
  a document rendered a full grid because every row carried exception borders,
  while the audit only inspected table-level `tblBorders`, style borders, and
  cell-level `tcBorders`.
- Normalizer: `normalize_docx --tables` now removes row-level exception borders
  so the delivery gate can safely clean hidden table grids before auditing.
- Tests: added regression coverage for row-exception table grids, body-cell
  bottom borders, and allowed last-row bottom rules.

## [1.6.0] - 2026-07-08

Driven by a third real-world failure (Codex delivered gridded tables, italic
formula digits, and unnumbered/unreferenced equations — and the audit "passed"):
prose rules are not executed, and **WARN does not stop delivery**. This release
makes the gate actually gate.

### Added
- `scripts/finalize_docx.py` — the mandatory one-command delivery gate: applies every safe mechanical fix in place (`normalize_docx --all` equivalent), runs the full audit, and prints `DELIVERY GATE: PASS/FAIL` with a non-zero exit on FAIL. `SKILL.md` now requires running it and pasting the verdict into the delivery note; `agents/openai.yaml`'s default prompt tells the agent not to deliver without a PASS.
- Auditor checks: `EQUATION_UNNUMBERED` (FAIL — a display equation with no chapter number `(3-1)`) and `EQUATION_NOT_REFERENCED` (WARN — a numbered equation never cited in prose as 由式 (3-1) 可得). Previously the auditor only checked the *layout* of numbers that existed; equations with no number at all were invisible.
- `TABLE_BORDERS` now resolves the referenced table style's `basedOn` chain, so grids drawn by a style (e.g. Table Grid, invisible in document.xml) are caught — this is how "所有表格都不是三线表" passed the audit.

### Changed
- **Severity promotions — every Must-Fix rule is now FAIL level** so it blocks the gate: `TABLE_BORDERS`, `TABLE_RULES`, `FORMULA_DIGIT_ITALIC`, `FORMULA_MULTILETTER_ITALIC`, `EQUATION_NUMBER_CENTER`, `EQUATION_NUMBER_TABS`, and `CAPTION_ALIGN` were WARN (the audit reported them but still said "pass" and exited 0).
- Must-Fix Audit gained the equation numbering/reference item; the audit-script item now names the delivery gate. Standard version → 2026-07-08. 12 new tests (136 total).



## [1.5.0] - 2026-07-07

Direction set by the blind-spot review: invest in executable tooling, not longer
prose rules. **MS Word is the canonical renderer**; the audit script is the
completion gate; PDF/page-render passes are not run unless explicitly requested.

### Added
- `scripts/replace_math.py` — the deterministic formula tool: converts a JSON registry of LaTeX formulas to native OMML (one Pandoc batch) and splices each equation **at the exact position of its plain-text original** (cross-run tokens handled, surrounding characters preserved, `w:sz`/`w:szCs` stamped, display equations laid out as left-aligned line + center tab + right-aligned number). Reports `not_found` / `still_plain_text`; namespace declarations (incl. `mc:Ignorable`) preserved verbatim. `--convert` prints the OMML of a single fragment.
- `references/generate-from-source.md` — source-first authoring for new documents: Markdown + LaTeX compiled with `pandoc --reference-doc`, correct by construction; harvesting existing OMML formulas back to LaTeX with `pandoc old.docx -t markdown` to seed the registry.
- Auditor checks: `FORMULA_TEXT_TABLE` (WARN — plain-text formulas/symbols inside table cells, previously a silent scope exclusion), `EQUATION_NUMBER_TABS` (WARN — numbered display equation with no right tab stop), `FIELDS_UPDATE` (WARN — TOC/REF fields present but `w:updateFields` unset, so field results may be stale), `FIRSTLINE_FIXED` (WARN — first-line indent in fixed twips instead of `firstLineChars`), and `PACKAGE_INTEGRITY` (FAIL — corrupt zip member or missing required package part, checked before everything else).
- `normalize_docx.py` opt-in auto-fixes: `--units` (insert the number-unit space, remove the space before `%`/`°`/`℃`, inside `w:t` text only), `--tables` (clear every in-table `w:shd` to `clear/auto`, zero the `w:tblLook` flags, add `w:tblHeader` to the first row of multi-row tables), `--update-fields` (set `w:updateFields` in settings.xml so MS Word refreshes fields on open), `--all`.
- CI installs Pandoc so the LaTeX→OMML round-trip tests run; 22 new tests (124 total).

### Changed
- The formula workflow now names `replace_math.py` as the preferred conversion mechanism; hand-editing OOXML for formulas is the fallback, not the norm.
- `SKILL.md` records the acceptance decisions (GB norms, general documents, MS Word canonical) and replaces the render-to-PDF step with "audit gate + user spot-check in Word; no render passes unless asked".
- The auditor's PASS line now states its scope explicitly so PASS is not misread as full compliance; scope includes `word/settings.xml`.

## [1.4.0] - 2026-07-07

Driven by two real-world failures: a delivered report whose three-line table kept the
table style's blue `firstRow` header shading, and formulas that were appended to
paragraph ends as (italic) duplicates instead of replacing the original text in place
(`……如图 1 所示。F = 44.5 km²L = 15.4 kmL/J^(1/3) = 75.1`).

### Added
- Formula workflow: an explicit **In-Place Replacement Contract** — every OMML object replaces the original token at its exact position; the plain-text original must be gone; never append math at paragraph ends, never merge adjacent expressions (`km²L`), never leave stray fragments (`mm，Cv`); self-check by comparing the math-stripped paragraph text against the original.
- Three-line table recipe: a **white-background section** — drop the `w:tblStyle` reference (header shading usually comes from the style's `firstRow` conditional format in styles.xml, invisible in document.xml), zero the `w:tblLook` flags, and clear every `w:shd` to `val=clear fill=auto`.
- Auditor checks: `TABLE_SHADING` (FAIL — direct or table-style-driven shading, resolved through the style's `basedOn` chain and `tblStylePr` conditional formats in `word/styles.xml`), `TABLE_RULES` (WARN — missing visible top/bottom rules, row-to-row `insideH` borders, or a header rule not thinner than the top/bottom rules), and `MATH_DUPLICATE` (FAIL — an OMML object's text still present as plain text in the same paragraph, i.e. append-instead-of-replace).
- Tests for all three checks; the non-compliant sample now demonstrates a shaded header cell, and the compliant sample builds a real explicit three-line table (borders + cleared shading).
- Italic/upright enforcement closing the "whole formula italicized" blind spots: the standard now documents that **OMML letters render math-italic by default** (units/functions need an explicit `m:sty val="p"`, i.e. LaTeX `\mathrm`/`\text` — "no italic tag" ≠ upright). New auditor checks: `MANUAL_ITALIC_MATH` (FAIL — an italicized plain-text pseudo-formula such as italic `F = 44.5 km²`, previously invisible because the italics check only looked inside OMML) and `FORMULA_MULTILETTER_ITALIC` (WARN — 2+ adjacent letters in a formula left at the italic default: a unit, function name, or `CI`-style coefficient). `FORMULA_DIGIT_ITALIC` now also catches mixed runs (an italic `F=44.5` used to be skipped because the run contained a letter).
- More layout guardrails: `CAPTION_ALIGN` (WARN — a 表/图 caption not centered), `FLOAT_ORDER` (WARN — a figure/table placed before its first in-text reference), `TABLE_HEADER_REPEAT` (WARN — a multi-row table whose header row lacks `w:tblHeader`), and `NUMBER_UNIT_SPACING` (WARN — `20km` glued together, or a space before `%`/`°`/`℃`).

### Fixed
- The auditor read the OMML run style with the wrong attribute namespace (`w:val` instead of `m:val` on `m:sty`), so OMML-style italics were never detected.

### Changed
- `SKILL.md`: reading `references/three-line-table-ooxml.md` is now a required workflow step for any document with tables; the in-place replacement rule joined the formula workflow step; the Must-Fix Audit gained "any cell shading / non-white fill" and "appended/duplicated rendered math" items.
- Auditor scope now includes `word/styles.xml` (table-style shading/borders only). Standard version → 2026-07-07.

## [1.3.0] - 2026-06-29

### Added
- Auditor checks: `FORMULA_DIGIT_ITALIC` (a number/operator italicized inside a formula — usually from blanket-italicizing the whole equation) and `EQUATION_NUMBER_CENTER` (a numbered display formula centered instead of having a right-aligned number).

### Changed
- Formula rules: do not blanket-italicize equations (digits/operators/parentheses stay upright); multi-letter coefficients use one variable + upright subscript (`C_I`, not `CI`); display equations use a left-aligned paragraph with a center tab (equation) and a right tab (number), not a centered line. Workflow doc gained an "Italic vs. upright" section.
- Appendix code: English/numbers in Times New Roman 小四 (12 pt), Chinese in Songti 小四 (was vaguely "monospace or Times New Roman"); added an appendix-code row to the font-size table.
- Standard version → 2026-06-29.

## [1.2.2] - 2026-06-28

### Added
- `references/three-line-table-ooxml.md`: a concrete OOXML recipe for the three-line table — exactly which borders to set (thick top/bottom `w:sz≈12`, a single thin header rule `w:sz≈6`, all other borders none), `w:sz` widths in eighths of a point, a python-docx note, and how to avoid gridded/header-shaded table styles.

### Changed
- Strengthened the three-line table rule in `SKILL.md` and `formatting-standard.md`: the top and bottom rules must be visible and thicker, the single middle (header) rule thinner; explicitly forbid gridded/header-shaded table styles. Fixes the common "no top/bottom rule, thick middle line" failure. Standard version → 2026-06-28.

## [1.2.1] - 2026-06-26

### Changed
- Standard: formulas, variables, and quantity symbols **inside a table** now follow the table text size (五号), not the body size (小四); formula size follows its context everywhere.
- Auditor: `TABLE_SIZE` now also inspects OMML math runs (`m:r`), so an in-table formula left at the wrong (larger) size is caught, not just ordinary text.

## [1.2.0] - 2026-06-26

### Added
- Auditor checks: `TABLE_BORDERS` (a table with direct vertical/inner borders instead of a three-line layout) and `HEADING_NO_STYLE` (a heading-looking line that uses no Word heading style, so it never enters the TOC).
- A **compliant** example: `examples/make_sample.py --compliant` builds a document the auditor reports as `PASS`, with its captured output in `examples/sample-compliant-audit-output.txt`.
- README font-size hierarchy table (English + 简体中文) and status badges (release, license, CI, stars).

### Changed
- Auditor: paragraph-based checks now **exclude table-cell text**, removing the `H1_CENTER`/`H1_GAP`/punctuation/formula false positives caused by numbers and short fragments in cells.
- Auditor: tightened the bare-symbol formula heuristic so prose like "为 A 方案" / "取 N 个" is no longer flagged as a formula; `BODY_FONT` no longer flags centered cover titles.
- Standard: default page margins are symmetric (~2.5 cm) with no binding gutter; standard version bumped to 2026-06-26.

## [1.1.0] - 2026-06-26

### Changed
- Headings: level 1–2 in Heiti (not bold), level-3 in bold Songti, cover title 二号; line spacing default 1.5×.
- Numbers & units: number–unit gap is a half-width (non-breaking) space; `%`, `°`, and `℃`/`°C` attach directly with no space.
- Tables: content font 五号 (one size below the body; 小四 also accepted), table notes 小五.
- Auditor: `TABLE_SIZE` accepts 五号/小四; `HEADING_FONT` limited to level 1–2 (level-3 is Songti by design); `COLOR` ignores hyperlink/theme colors.

### Added
- Bilingual README: `README.md` (English) and `README.zh-CN.md` (简体中文) with a language switcher.

## [1.0.0] - 2026-06-25

### Added
- Initial release: the written standard (`SKILL.md` + `references/`), the `audit_docx_format.py` guardrail (punctuation, headings, fonts, captions, tables, citations, formulas; `--json` output and a FAIL/WARN summary), the `normalize_docx.py` safe auto-fixer, a non-compliant example, standard-library unit tests, and GitHub Actions CI.

[1.6.3]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.6.3
[1.6.2]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.6.2
[1.6.1]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.6.1
[1.6.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.6.0
[1.5.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.5.0
[1.4.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.4.0
[1.3.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.3.0
[1.2.2]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.2
[1.2.1]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.1
[1.2.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.0
[1.1.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.1.0
[1.0.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.0.0
