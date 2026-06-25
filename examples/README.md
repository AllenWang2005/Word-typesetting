# Examples

A minimal, reproducible demonstration of the audit script.

## Files

- `make_sample.py` — builds a deliberately **non-compliant** `sample.docx` with python-docx.
- `sample-audit-output.txt` — the captured output of running the auditor on that sample.

## Reproduce

```bash
pip install python-docx          # only needed to build the sample
python examples/make_sample.py sample.docx
python scripts/audit_docx_format.py sample.docx
```

The sample intentionally violates several rules, so the auditor reports a mix of
`FAIL` and `WARN` items (see `sample-audit-output.txt`):

| Code | Planted problem |
| --- | --- |
| `H1_CENTER` | Level-1 heading `一、绪论` is centered instead of left aligned |
| `ABSTRACT_INDENT` | Abstract paragraph is centered and first-line indented |
| `ZH_PUNCT` | ASCII comma / period used inside Chinese prose |
| `COLOR` | Red body run that is not the allowed appendix-comment red |
| `FORMULA_TEXT` | `式中 N_p ...` left as plain text instead of OMML |
| `VISIBLE_LATEX` | `\frac{a}{b}` left as visible LaTeX source |
| `CITATION_BRACKETS` | Full-width citation brackets `［1］` |
| `CITATION_FIELDS` | ASCII citation `[2]` with no backing `REF ref_###` field |
| `TABLE_SIZE` | Table text set to 14 pt instead of small-four / 12 pt |

A fully compliant document produces `PASS: no machine-detected guardrail issues.`

> Note: the auditor is a guardrail for machine-checkable issues only. Visual rules
> (caption above/below, line widths, page numbering, GB/T 7714 entry format) still
> require human or rendered-PDF review per the main standard.
