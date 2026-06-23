---
name: word-report-formatting
description: Apply Allen's formal Chinese Word report formatting standard. Use when creating, editing, reviewing, or finalizing Word/DOCX reports, course designs, internship reports, thesis-style documents, scientific/engineering reports, formulas, three-line tables, figure/table captions, or appendices with code comments.
---

# Word Report Formatting

## Overview

Use this skill to format formal Chinese Word reports consistently with Allen's preferred standard. It complements the general DOCX/document skill: use the document tooling for implementation and rendering, and use this skill for the required typography, formula, table, figure, and appendix-code rules.

## Required Workflow

1. Read `references/formatting-standard.md` before creating or substantially editing a Word report.
2. Apply the standard unless the user, school template, or provided rubric explicitly overrides it.
3. Preserve existing document content and structure during edit tasks; make only the local formatting changes needed.
4. For formulas, inspect the actual OMML/Word equation formatting rather than judging from plain text extraction.
5. Before delivery, run a formatting audit against the checklist in `references/formatting-standard.md`.
6. When possible, render/export the DOCX to PDF or page images and inspect pages for overlap, clipping, table overflow, font substitution, and appendix code readability.

## Core Rules

- Body text: Chinese in Songti/SimSun, English and numbers in Times New Roman, default small-four/12 pt.
- Headings: real Word heading styles with outline levels so navigation pane and automatic TOC work.
- Tables: default to white three-line tables, with no fill, no vertical lines, and no full grid.
- Formulas: use native Word OMML equations. Variables are italic; digits, operators, units, function names, constants, explanatory text, and explanatory subscripts are upright.
- Figures: keep plot titles out of the image when a Word caption is present; use consistent figure numbering and captions.
- Appendix code: put code after main text and tables; color code comments red, preferably `C00000`.

## Conflict Handling

If older project memory says "formulas should not be italic", treat that as superseded for formal report formulas. The current rule is: variables italic, non-variables upright, following the provided teacher screenshot/specification.
