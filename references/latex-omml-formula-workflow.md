# LaTeX to Word OMML Formula Workflow

Use this guide whenever a Word/DOCX report contains formulas, inline variables, quantity symbols, mathematical objects, subscripts/superscripts, units in expressions, or math-like plain text.

## Completion Contract

- Every display formula must be authored as LaTeX and rendered as native Word OMML, not pasted as an image, plain text, or visible LaTeX source.
- Inline variables and quantity symbols in prose, tables, and captions must also be authored as LaTeX and rendered as inline OMML when they carry mathematical meaning.
- Manual italic/upright styling of ordinary text is not enough for formula compliance.
- **Every conversion is an in-place replacement** — see the In-Place Replacement Contract below. Appending rendered math anywhere else is a hard failure.
- Keep a temporary formula registry during editing: location, original text, LaTeX source, inline/display, and conversion status.
- If a formula or symbol cannot be confidently interpreted, preserve the content and report it instead of guessing silently.

## In-Place Replacement Contract (hard rule)

The most damaging real-world failure is **append-instead-of-replace**: the model leaves the original plain-text token in the prose and dumps rendered (or worse, manually italicized) math at the end of the paragraph, producing garbage like `……如图 1 所示。F = 44.5 km²L = 15.4 kmL/J^(1/3) = 75.1`. Every rule below is mandatory:

- Each rendered OMML object must **replace the original token at its exact position** in the run sequence. The characters before and after the token must remain the same characters that surrounded it in the source text.
- After conversion, the original plain-text form of that token must **no longer exist anywhere** in the paragraph. Never keep the plain text "for safety" and add math next to it.
- **Never append** converted math at the end of a paragraph, the end of a section, or in a new paragraph, and never batch several conversions into one trailing cluster. One token, one in-place swap.
- Convert each occurrence **separately**. Do not merge adjacent expressions into one object (no `km²L` run-ons from gluing `F = 44.5\,\mathrm{km^2}` to `L = 15.4\,\mathrm{km}`); the prose characters between them (Chinese comma, `，`, spaces) must survive untouched.
- Do not leave stray fragments behind: no orphaned units, commas, or half-tokens such as `mm，Cv` hanging at a paragraph end.
- **Self-check after every paragraph**: read the paragraph text with math stripped, and compare it to the original text with the converted tokens removed. They must match exactly. If any converted value now appears twice (once as prose, once as math), the edit is wrong — revert and redo it in place.
- The auditor reports `MATH_DUPLICATE` (FAIL) when an OMML object's text also still appears as plain text in the same paragraph.

## Formula Discovery Pass

Scan the whole document before formatting. Treat all of these as candidates:

- Existing Word equations, equation fields, images of formulas, MathType/OLE objects, and visible LaTeX.
- Paragraphs or table cells containing operators such as `=`, `+`, `-`, `/`, `^`, `≤`, `≥`, `∑`, `Σ`, `∏`, `√`, fractions, integrals, or formula numbering.
- Inline quantity symbols in explanatory prose, such as `Q`, `N_p`, `x-y`, `P=0.01%`, `R_{SN}`, `T_N`, `T_D`, `I_i`, or any single letter that denotes a physical quantity.
- Symbol definitions after words such as `其中`, `式中`, `符号`, `变量`, `计算式`, `公式`, `取`, `为`, or `表示`.
- Bare one-letter symbols in definition/explanation contexts, e.g. `式中 Q 为流量，N 为出力`, `其中 H 表示水头`, and `V 为流速`. Do not skip them just because they have no subscript.
- Table headers/cells that contain variables, units coupled to variables, subscripts, superscripts, or engineering symbols.

Do not skip inline math just because it is short. A single variable such as `Q`, `N`, `H`, `V`, or `N_p` still needs inline OMML when it is a quantity symbol.

## LaTeX Authoring Rules

- Single-letter variables are italic by default in math mode: `Q`, `N`, `Z`, `V`, `T`.
- Variables with subscripts use LaTeX subscripts: `N_p`, `T_N`, `I_i`, `R_{\mathrm{SN}}`.
- Explanatory subscripts and fixed abbreviations are upright: `T_{\mathrm{N}}`, `T_{\mathrm{D}}`, `R_{\mathrm{SN}}`, `V_{\text{死}}`.
- A multi-letter abbreviation cannot directly be the variable. Use one variable plus upright explanatory subscript, e.g. signal-noise ratio as `R_{\mathrm{SN}}`, not italic `SNR`.
- Function names and operators are upright: `\sin`, `\exp`, `\lg`, `\max`, `\mathrm{erf}`.
- Units are upright and separated from values with a thin space: `43.50\,\mathrm{万kW}`, `20000\,\mathrm{m^3/s}`.
- Constants and special non-variable symbols are upright where the renderer supports it: `\mathrm{e}`, `\pi`, `\mathrm{i}`, `\mathrm{d}`, `\partial`, `\Delta`, `\Sigma`, `\Pi`.
- Use structured LaTeX for fractions, roots, sums, products, piecewise definitions, and aligned equations: `\frac{}`, `\sqrt{}`, `\sum`, `\prod`, `\begin{aligned}...\end{aligned}`.
- Use Chinese explanatory text outside equations when possible. If Chinese text must appear inside a formula, use `\text{...}` and verify the OMML rendering.

## Conversion Strategy

Prefer this pipeline for DOCX work:

1. Replace each formula/symbol occurrence with a unique placeholder such as `@@MATH_001@@` while preserving surrounding paragraph/table structure.
2. Store the LaTeX source in the formula registry. Use `$...$` for inline math and `$$...$$` or display blocks for display math.
3. Convert LaTeX fragments to Word OMML with a reliable converter, preferably Pandoc-generated DOCX, then extract the resulting `<m:oMath>` or `<m:oMathPara>` XML.
4. Replace placeholders in the target DOCX with the converted OMML XML **at the placeholder's exact position** (see the In-Place Replacement Contract). Inline symbols use inline `<m:oMath>`; display formulas use a centered equation paragraph and, when needed, a right-aligned equation number.
5. Set the OMML run font size to match the surrounding context: body formulas at the body size (小四, `w:sz=24`), formulas inside a table at the table size (五号, `w:sz=21`). Do not leave an in-table formula at the body's 小四 — it must match the 五号 table text.
6. Verify that no placeholder, visible LaTeX source, formula image, plain-text substitute, or duplicated plain-text original remains.

When Pandoc is unavailable, use another route that still produces native Word OMML, such as Word's equation conversion or an OMML-capable library. Do not fall back to styled normal text unless the user explicitly accepts a non-compliant visual-only draft.

## Display Formula Layout

- The equation is centered and its number is right-aligned, but do **not** set the paragraph alignment to centered — that centers the number too. Use a left-aligned paragraph with two tab stops: a **center tab** at the column midpoint (equation) and a **right tab** at the right margin (number). Insert `<tab>` + equation + `<tab>` + `(3-3)`. A borderless 1×3 table (empty / equation centered / number right-aligned) is an acceptable alternative.
- Number formulas by chapter in parentheses, e.g. `(3-1)`, `(3-2)`; reference them in text as "由式 (3-1) 可得".
- Multi-line formulas should be aligned in LaTeX before conversion, typically with `aligned`.
- Keep explanatory text such as `式中：` outside the equation when possible, with the following quantity symbols rendered as inline OMML.

## Italic vs. upright (do not blanket-italicize)

- Only variable letters are italic. Digits, operators, parentheses, commas, units, and function names are upright. In OMML this is the default — do **not** add `<w:i/>` to the whole equation, which forces the digits italic too (the most common failure).
- Multi-letter coefficients are not one italic variable: use one variable plus an upright subscript, e.g. recession coefficients `C_I` / `C_G` / `C_S` (italic `C`, upright `I`/`G`/`S`), never adjacent italic letters `CI` / `CG` (reads as `C × I`).
- The auditor flags `FORMULA_DIGIT_ITALIC` (an italic number/operator) and `EQUATION_NUMBER_CENTER` (a centered numbered equation).

## Verification

After editing the DOCX:

- Inspect `word/document.xml` and confirm that formulas/symbols are represented by `m:oMath` or `m:oMathPara`.
- Search for placeholders like `@@MATH_`, visible LaTeX commands such as `\frac`, and obvious plain-text equations such as `Q =` or `N_p =`.
- Check for append-instead-of-replace damage: no paragraph may contain both an OMML object and the same expression as plain text, and no paragraph may end with a trailing cluster of math objects or stray fragments (`mm，Cv`) that repeat values already in the prose (auditor: `MATH_DUPLICATE`).
- Check that variables are italic and non-variables are upright in the rendered Word output.
- Render/export to PDF or page images and inspect formula baseline, line height, equation numbering, and table-cell formula fit.
- If `scripts/audit_docx_format.py` reports formula warnings/failures, revise or explicitly explain why an item is intentionally not mathematical.
