# Citation Cross-References OOXML Guide

Use this guide whenever a DOCX contains paper-style body citations such as `[1]`, `[1][2]`, `[1,2]`, `[1-3]`, `［1］`, or `【1】`, and the user asks for Word report formatting or reference/citation cleanup.

## Completion Contract

- Body citations must be superscript.
- Citation brackets must be ASCII square brackets: `[` and `]`.
- Citations must be backed by Word fields and bibliography bookmarks. Static superscript text alone is not complete unless the user explicitly asks for visual-only cleanup.
- Missing references must be reported instead of silently renumbered or guessed.

## Required OOXML Strategy

1. Locate the bibliography section. Accept headings such as `参考文献`, `References`, or a heading style whose visible text matches these labels.
2. Build a map from reference number to bibliography paragraph. Common entry starts include `[1]`, `［1］`, `1.`, `1、`, and `[1] 作者...`.
3. For each bibliography entry, create a stable bookmark named `ref_001`, `ref_002`, and so on.
4. The bookmark should wrap only the visible Arabic number token when possible, not the surrounding brackets or punctuation. If the entry starts as `[1]`, split the run so the bookmark wraps `1` only and leaves `[` and `]` outside.
5. Scan only the main body before the bibliography section for citations. Do not rewrite the bibliography entries as if they were citations.
6. Replace each citation group with superscript runs and field-backed references:
   - `[1]` becomes superscript `[` + `REF ref_001 \h` + `]`.
   - `[1][2]` becomes superscript `[` + `REF ref_001 \h` + `]` + `[` + `REF ref_002 \h` + `]`.
   - `[1,2]` or `[1，2]` becomes superscript `[` + `REF ref_001 \h` + `,` + `REF ref_002 \h` + `]`.
   - `[1-3]`, `[1–3]`, or `[1—3]` becomes superscript `[` + `REF ref_001 \h` + `-` + `REF ref_003 \h` + `]`.
7. Apply `w:vertAlign w:val="superscript"` to every run in the citation group, including brackets, commas, hyphens, and field result runs.
8. Preserve the surrounding paragraph style and non-citation runs.

## Simple Field Pattern

Prefer a simple Word field when patching OOXML directly:

```xml
<w:r>
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  <w:t>[</w:t>
</w:r>
<w:fldSimple w:instr=" REF ref_001 \h ">
  <w:r>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
    <w:t>1</w:t>
  </w:r>
</w:fldSimple>
<w:r>
  <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  <w:t>]</w:t>
</w:r>
```

Complex fields using `w:fldChar` and `w:instrText` are also acceptable, but the final `document.xml` must contain a `REF ref_###` instruction for each body citation target.

## Bookmark Pattern

Use unique integer bookmark IDs that do not collide with existing bookmarks:

```xml
<w:bookmarkStart w:id="42" w:name="ref_001"/>
<w:r><w:t>1</w:t></w:r>
<w:bookmarkEnd w:id="42"/>
```

If the number is inside an existing run such as `[1]`, split that run into separate runs for `[`, `1`, and `]`, then wrap only the `1` run.

## Verification

After patching, inspect `word/document.xml` and verify:

- Every referenced bibliography item has a `w:bookmarkStart` named `ref_###`.
- Every body citation target has either `w:fldSimple w:instr=" REF ref_### \h "` or complex field instruction text containing `REF ref_###`.
- Citation runs use `w:vertAlign w:val="superscript"`.
- Citation brackets are ASCII `[` and `]`, not `［］`, `【】`, or other full-width punctuation.
- No baseline plain-text citation patterns remain in body paragraphs before the bibliography section.

Then update fields in Word when possible, or disclose if field updating could not be performed in the current environment.
