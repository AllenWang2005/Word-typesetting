---
name: word-report-formatting
description: Apply Allen's formal Chinese Word report formatting standard. Use when creating, editing, reviewing, or finalizing Word/DOCX reports, course designs, internship reports, thesis-style documents, scientific/engineering reports, formulas, three-line tables, figure/table captions, reference/citation cross-references, or appendices with code comments.
---

# Word Report Formatting

## Overview

Use this skill to format formal Chinese Word reports consistently with Allen's preferred standard. It complements the general DOCX/document skill: use the document tooling for implementation and rendering, and use this skill for the required typography, formula, table, figure, reference/citation, and appendix-code rules.

## Required Workflow

1. Read `references/formatting-standard.md` before creating or substantially editing a Word report.
2. Apply the standard unless the user, school template, or provided rubric explicitly overrides it.
3. Preserve existing document content and structure during edit tasks; make only the local formatting changes needed.
4. For formulas, inspect the actual OMML/Word equation formatting rather than judging from plain text extraction.
5. For paper-style citations, read `references/citation-crossrefs-ooxml.md` and convert body citations such as `[1]`, `[1][2]`, `[1,2]`, or `[1-3]` into superscript Word cross-references to the matching bibliography entries.
6. Before delivery, run a formatting audit against the checklist in `references/formatting-standard.md`.
7. When possible, render/export the DOCX to PDF or page images and inspect pages for overlap, clipping, table overflow, font substitution, broken cross-references, and appendix code readability.

## Core Rules

- Body text: Chinese in Songti/SimSun, English and numbers in Times New Roman, default small-four/12 pt.
- Body color: keep body text, headings, captions, tables, formulas, abstract, keywords, and references black unless the user explicitly requests otherwise. Appendix code comments are the standard red exception.
- Abstract and keywords: left aligned with no first-line indent.
- Headings: real Word heading styles with outline levels, left aligned rather than centered. Level-1 Chinese headings such as `一、` / `二、` / `三、` must have one blank line of visual space before them when not at the top of a page.
- Tables: default to white three-line tables, with no fill, no vertical lines, and no full grid. Normalize existing tables that do not meet the standard.
- Formulas: use native Word OMML equations. Variables are italic; digits, operators, units, function names, constants, explanatory text, and explanatory subscripts are upright.
- Figures: keep plot titles out of the image when a Word caption is present; use consistent figure numbering and centered figure/table captions.
- Citations: make body references superscript and field-backed, not merely static superscript text. Use ASCII square brackets `[` and `]`, not Chinese/full-width brackets.
- Appendix code: put code after main text and tables; color code comments red, preferably `C00000`.

## Must-Fix Audit

Fail the formatting pass and revise if any of these remain:

- Headings are centered when the standard says they should be left aligned.
- Abstract or keywords are indented or centered.
- Body text, headings, captions, table text, or reference text contain unintended non-black color.
- Figure names or table names are not centered.
- Level-1 Chinese headings do not have a blank-line visual gap before them when they follow body text.
- Existing tables, formulas, variables, units, or symbols still violate the standard.
- Body citations are only static superscript text instead of Word `REF` fields/bookmark cross-references, unless the user explicitly allows visual-only citations.

## Conflict Handling

If older project memory says "formulas should not be italic", treat that as superseded for formal report formulas. The current rule is: variables italic, non-variables upright, following the provided teacher screenshot/specification.
