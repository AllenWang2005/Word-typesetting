# LaTeX to Word OMML Formula Workflow

Use this guide whenever a Word/DOCX report contains formulas, inline variables, quantity symbols, mathematical objects, subscripts/superscripts, units in expressions, or math-like plain text.

## Completion Contract

- Every display formula must be authored as LaTeX and rendered as native Word OMML, not pasted as an image, plain text, or visible LaTeX source.
- Inline variables and quantity symbols in prose, tables, and captions must also be authored as LaTeX and rendered as inline OMML when they carry mathematical meaning.
- Manual italic/upright styling of ordinary text is not enough for formula compliance.
- Keep a temporary formula registry during editing: location, original text, LaTeX source, inline/display, and conversion status.
- If a formula or symbol cannot be confidently interpreted, preserve the content and report it instead of guessing silently.

## Formula Discovery Pass

Scan the whole document before formatting. Treat all of these as candidates:

- Existing Word equations, equation fields, images of formulas, MathType/OLE objects, and visible LaTeX.
- Paragraphs or table cells containing operators such as `=`, `+`, `-`, `/`, `^`, `≤`, `≥`, `∑`, `Σ`, `∏`, `√`, fractions, integrals, or formula numbering.
- Inline quantity symbols in explanatory prose, such as `Q`, `N_p`, `Z-V`, `P=0.01%`, `R_SN`, `Re`, `We`, `Ma`, `T_N`, `T_D`, `I_i`.
- Symbol definitions after words such as `其中`, `式中`, `符号`, `变量`, `计算式`, `公式`, `取`, `为`, or `表示`.
- Table headers/cells that contain variables, units coupled to variables, subscripts, superscripts, or engineering symbols.

Do not skip inline math just because it is short. A single variable such as `Q` or `N_p` still needs inline OMML when it is a quantity symbol.

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
4. Replace placeholders in the target DOCX with the converted OMML XML. Inline symbols use inline `<m:oMath>`; display formulas use a centered equation paragraph and, when needed, a right-aligned equation number.
5. Verify that no placeholder, visible LaTeX source, formula image, or plain-text substitute remains.

When Pandoc is unavailable, use another route that still produces native Word OMML, such as Word's equation conversion or an OMML-capable library. Do not fall back to styled normal text unless the user explicitly accepts a non-compliant visual-only draft.

## Display Formula Layout

- Display formulas are centered as equations. Formula numbers are right aligned, usually `（式x-x）` or the project's existing numbering style.
- Multi-line formulas should be aligned in LaTeX before conversion, typically with `aligned`.
- Keep explanatory text such as `式中：` outside the equation when possible, with the following quantity symbols rendered as inline OMML.

## Verification

After editing the DOCX:

- Inspect `word/document.xml` and confirm that formulas/symbols are represented by `m:oMath` or `m:oMathPara`.
- Search for placeholders like `@@MATH_`, visible LaTeX commands such as `\frac`, and obvious plain-text equations such as `Q =` or `N_p =`.
- Check that variables are italic and non-variables are upright in the rendered Word output.
- Render/export to PDF or page images and inspect formula baseline, line height, equation numbering, and table-cell formula fit.
- If `scripts/audit_docx_format.py` reports formula warnings/failures, revise or explicitly explain why an item is intentionally not mathematical.
