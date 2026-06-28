# Changelog

All notable changes to this skill are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/), and the project uses semantic versioning.

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

[1.2.2]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.2
[1.2.1]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.1
[1.2.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.2.0
[1.1.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.1.0
[1.0.0]: https://github.com/AllenWang2005/Word-typesetting/releases/tag/v1.0.0
