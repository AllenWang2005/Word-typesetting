# Citation Cross-References OOXML Guide

Use this guide whenever a DOCX contains paper-style body citations such as `[1]`, `[1][2]`, `[1,2]`, `[1-3]`, `［1］`, or `【1】`, and the user asks for Word report formatting or reference/citation cleanup.

## Completion Contract

- Body citations must be superscript.
- Citation brackets must be ASCII square brackets: `[` and `]`.
- The whole bracketed citation `[1]` — brackets included — is the cross-reference unit. Reference the entire `[1]`, not just the bare number `1`; the brackets must travel with the field result, not be left as detached static characters.
- Citations must be backed by Word fields and bibliography bookmarks. Static superscript text alone is not complete unless the user explicitly asks for visual-only cleanup.
- Missing references must be reported instead of silently renumbered or guessed.

## Required OOXML Strategy

1. Locate the bibliography section. Accept headings such as `参考文献`, `参考资料`, `References`, or a heading style whose visible text matches these labels.
2. Build a map from reference number to bibliography paragraph. Common entry starts include `[1]`, `［1］`, `1.`, `1、`, and `[1] 作者...`.
3. For each bibliography entry, create a stable bookmark named `ref_001`, `ref_002`, and so on.
4. The bookmark should wrap the **entire `[1]` token** — the ASCII brackets and the number — so a single `REF` field reproduces the whole `[1]`. If the bibliography entry has no brackets (starts as `1.` or `1、`), bookmark the bare number `1` and add the ASCII brackets in the body so the in-text result is still `[1]`.
5. Scan only the main body before the bibliography section for citations. Do not rewrite the bibliography entries as if they were citations.
6. Replace each citation with superscript, field-backed references that include the brackets:
   - `[1]` becomes a single superscript `REF ref_001 \h` field whose result is the whole `[1]` (bookmark wraps `[1]`).
   - `[1][2]` becomes two adjacent superscript fields, each reproducing its full `[n]`: `[1]` + `[2]`.
   - For combined forms the source keeps (`[1,2]`, `[1，2]`, `[1-3]`, `[1–3]`, `[1—3]`): keep one shared ASCII bracket pair and put a `REF` field around each number inside, e.g. superscript `[` + `REF ref_001 \h` + `,` + `REF ref_002 \h` + `]` or `[` + `REF ref_001 \h` + `-` + `REF ref_003 \h` + `]`. In this fallback the bookmarks wrap the bare numbers.
7. Apply `w:vertAlign w:val="superscript"` to every run in the citation group, including brackets, commas, hyphens, and field result runs.
8. Preserve the surrounding paragraph style and non-citation runs.

## Simple Field Pattern

Prefer a simple Word field when patching OOXML directly:

```xml
<w:fldSimple w:instr=" REF ref_001 \h ">
  <w:r>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
    <w:t>[1]</w:t>
  </w:r>
</w:fldSimple>
```

Here the bibliography bookmark wraps the whole `[1]`, so the single field result is `[1]`. (In the combined-range fallback, the bookmark wraps the bare number and you add shared `[` / `]` superscript runs around the field group.)

Complex fields using `w:fldChar` and `w:instrText` are also acceptable, but the final `document.xml` must contain a `REF ref_###` instruction for each body citation target.

## Bookmark Pattern

Use unique integer bookmark IDs that do not collide with existing bookmarks:

```xml
<w:bookmarkStart w:id="42" w:name="ref_001"/>
<w:r><w:t>[1]</w:t></w:r>
<w:bookmarkEnd w:id="42"/>
```

If the entry's `[1]` is split across runs, normalize it into one run (or wrap the brackets and number together) so the bookmark covers the whole `[1]`. Only when you must support combined ranges like `[1-3]` should you instead bookmark the bare number `1` and add shared brackets in the body.

## Verification

After patching, inspect `word/document.xml` and verify:

- Every referenced bibliography item has a `w:bookmarkStart` named `ref_###`.
- Every body citation target has either `w:fldSimple w:instr=" REF ref_### \h "` or complex field instruction text containing `REF ref_###`.
- The field result includes the ASCII brackets (`[1]`), or, in the combined-range fallback, shared `[ ]` runs surround the field group — no citation renders as a bare number without brackets.
- Citation runs use `w:vertAlign w:val="superscript"`.
- Citation brackets are ASCII `[` and `]`, not `［］`, `【】`, or other full-width punctuation.
- No baseline plain-text citation patterns remain in body paragraphs before the bibliography section.

Then update fields in Word when possible, or disclose if field updating could not be performed in the current environment.
