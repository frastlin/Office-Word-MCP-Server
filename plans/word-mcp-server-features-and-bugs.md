# Word Document MCP Server — Issues

**Date:** 2026-02-01
**Context:** These issues were identified while using the `word-document-server` MCP to perform ~12 structural edits (paragraph replacements, section deletions, text condensing) on a 12,684-word academic paper (`tvm-paper.docx`).

---

## Issue 1 [Feature]: Add `delete_paragraph_range(filename, start_index, end_index)`

**Problem:** There is no way to delete multiple consecutive paragraphs in a single call. The only option is `delete_paragraph(filename, paragraph_index)`, which deletes one paragraph at a time and must be called repeatedly from the highest index downward to avoid index shifting.

**Example:** A document contains a section titled "Addressing Potential Counterarguments" spanning 8 paragraphs (a heading, 3 content paragraphs, and 4 empty spacing paragraphs):

```
Index 384: "### Addressing Potential Counterarguments"          (Heading 3)
Index 385: ""                                                    (empty spacer)
Index 386: "A potential objection is that interactive text..."   (Normal)
Index 387: ""                                                    (empty spacer)
Index 388: "Implementation costs are a valid concern..."         (Normal)
Index 389: ""                                                    (empty spacer)
Index 390: "Legacy maps can be made accessible provided..."      (Normal)
Index 391: ""                                                    (empty spacer)
```

To delete this entire section, 8 sequential `delete_paragraph` calls are required, working from the highest index down:

```
delete_paragraph(file, 391)
delete_paragraph(file, 390)
delete_paragraph(file, 389)
delete_paragraph(file, 388)
delete_paragraph(file, 387)
delete_paragraph(file, 386)
delete_paragraph(file, 385)
delete_paragraph(file, 384)
```

**Current Workaround:** Call `delete_paragraph` in a loop from the highest index down to the lowest. This requires 8 round trips instead of 1.

**Suggested API:**
```python
delete_paragraph_range(filename: str, start_index: int, end_index: int) -> str
# Deletes all paragraphs from start_index to end_index inclusive.
# Example: delete_paragraph_range("doc.docx", 384, 391)
# Deletes 8 paragraphs in one call.
```

**Alternative:** `replace_paragraph_range` with an empty `new_paragraphs: []` array could also serve this purpose, but it is unclear whether passing an empty array is supported. If it is, document this explicitly. If not, add support for it.

---

## Issue 2 [Feature]: Add `get_paragraph_range(filename, start_index, end_index)`

**Problem:** There is no way to read multiple paragraphs in a single call. `get_paragraph_text_from_document` reads exactly one paragraph per call. When mapping out the structure of a section (e.g., determining where a heading's content ends), you must make many sequential calls.

**Example:** A document has a "Policy Implications" section starting at paragraph 366. To understand the section's structure, you need to read each paragraph individually to find content paragraphs, empty spacing lines, figure references, and subheadings:

```
get_paragraph_text_from_document(file, 366)  → "## Policy Implications"     (Heading 2)
get_paragraph_text_from_document(file, 367)  → ""                            (empty spacer)
get_paragraph_text_from_document(file, 368)  → "Some digital accessibility..." (Normal, ~800 chars)
get_paragraph_text_from_document(file, 369)  → ""                            (empty spacer)
get_paragraph_text_from_document(file, 370)  → "\u00a0"                      (non-breaking space spacer)
get_paragraph_text_from_document(file, 371)  → ""                            (empty spacer)
get_paragraph_text_from_document(file, 372)  → "![Figure 3: A comparison...]" (image reference)
get_paragraph_text_from_document(file, 373)  → ""                            (empty spacer)
get_paragraph_text_from_document(file, 374)  → "[Note: Here is a link...]"   (hyperlink)
... (6+ more calls to find section end)
```

This required 15 separate calls to map out one section.

**Current Workaround:** Make many parallel `get_paragraph_text_from_document` calls in one message. This works but clutters the context and is inefficient.

**Suggested API:**
```python
get_paragraph_range(filename: str, start_index: int, end_index: int) -> list[dict]
# Returns a list of paragraph objects from start_index to end_index inclusive.
# Each object includes: index, text, style, is_heading
# Example: get_paragraph_range("doc.docx", 366, 392)
# Returns 27 paragraph objects in one call.
```

---

## Issue 3 [Feature]: Add `get_section_paragraphs(filename, heading_text)`

**Problem:** There is no way to get all paragraphs under a specific heading (up to the next same-or-higher-level heading) in a single call. This is the most common operation when editing a paper: "show me everything under the heading 'Preference and Cognitive Load'."

Currently this requires three steps:
1. `find_text_in_document` to locate the heading index
2. Multiple `get_paragraph_text_from_document` calls to walk forward through paragraphs
3. Manually checking each paragraph's `is_heading` and `style` to detect where the section ends

**Example:** To find the content under "## Preference and Cognitive Load":

Step 1: `find_text_in_document(file, "Preference and Cognitive Load")` → paragraph index 336

Step 2: Read paragraphs one by one:
```
Index 336: "## Preference and Cognitive Load"                    (Heading 2)
Index 337: ""                                                    (empty spacer)
Index 338: "When spatial information is present, participants..." (Normal, content)
Index 339: ""                                                    (empty spacer)
Index 340: "Although the preference question focused..."         (Normal, content)
Index 341: ""                                                    (empty spacer)
Index 342: "The apparent paradox of higher cognitive..."         (Normal, content)
Index 343: ""                                                    (empty spacer)
Index 344: "If visual maps serve as a benchmark..."             (Normal, content)
Index 345: ""                                                    (empty spacer)
Index 346: "## Theoretical Implications"                         (Heading 2 — STOP)
```

Step 3: Determine that section content spans indices 337–345.

This required 11 individual calls. `get_section_paragraphs` would return all of this in one call.

**Suggested API:**
```python
get_section_paragraphs(filename: str, heading_text: str) -> dict
# Returns:
# {
#   "heading_index": 336,
#   "heading_text": "## Preference and Cognitive Load",
#   "heading_style": "Heading 2",
#   "content_start_index": 337,
#   "content_end_index": 345,
#   "next_heading_index": 346,
#   "paragraphs": [
#     {"index": 337, "text": "", "style": "Normal", "is_heading": false},
#     {"index": 338, "text": "When spatial information is present, participants prefer map interfaces over tables. The preference order is visual map, ITM, and then table...", "style": "Normal", "is_heading": false},
#     ...
#   ]
# }
```

---

## Issue 4 [Feature]: `find_text_in_document` — add option to return full paragraph text

**Problem:** `find_text_in_document` returns a truncated `context` field (approximately 100 characters). After finding text, you almost always need to call `get_paragraph_text_from_document` to read the full paragraph before deciding what to do with it. This doubles the number of calls for every search.

**Example:** Searching for `"One reason the \"primary purpose\""` returns:

```json
{
  "paragraph_index": 380,
  "position": 0,
  "context": "One reason the \"primary purpose\" strategy may exist is that digital accessibility professionals may ..."
}
```

The actual paragraph at index 380 is 1,200 characters long:

```
"One reason the \"primary purpose\" strategy may exist is that digital accessibility
professionals may never have encountered a non-visual digital map that fully communicates
spatial information, primarily because such tools have not been widely available until
recently. This is no longer a valid excuse. International accessibility legislation, such
as the European Union's directive [@EU2017] and the United Kingdom's regulations
[@PublicSectorBodiesUKAccessibleWebsitesAndApps2018], have historically made explicit
exceptions for digital geographic thematic maps. Section 508 in the United States states
that \"If there are technically acceptable solutions available in the marketplace, you must
select one of those solutions...\" [...continues for ~600 more characters]"
```

The truncated context is insufficient to verify this is the correct paragraph or to understand its full content. A follow-up `get_paragraph_text_from_document(file, 380)` call is always needed.

**Suggested change:** Add an optional `include_paragraph_text: bool = false` parameter. When true, return the full paragraph text instead of the truncated context.

```python
find_text_in_document(
    filename: str,
    text_to_find: str,
    match_case: bool = True,
    whole_word: bool = False,
    include_paragraph_text: bool = False  # NEW
) -> dict
# When include_paragraph_text=True, each occurrence includes:
# { "paragraph_index": 380, "position": 0, "text": "<full paragraph text>", "style": "Normal" }
```

---

## Issue 5 [Feature]: `get_document_info` — add option to include document outline

**Problem:** `get_document_info` returns word count, paragraph count, and table count, but not a list of headings or section structure. When starting to edit a document, the first thing you need is the document's outline to plan your edits. Currently this requires a separate `get_document_outline` call.

**Example:** Current `get_document_info` response:
```json
{
  "title": "",
  "author": "",
  "word_count": 12544,
  "paragraph_count": 475,
  "table_count": 0
}
```

To also get the heading structure, a separate `get_document_outline` call is needed.

**Suggested change:** Add an argument to `get_document_info` called `include_outline` as a boolean that is false by default. When true, `get_document_info` will return a `headings` array containing heading text, style, and paragraph index for each heading in the document.

---

## Issue 6 [Feature]: Add batch `find_texts_in_document(filename, texts_to_find)`

**Problem:** When preparing for multiple edits, you need to locate many different text strings. Each requires a separate `find_text_in_document` call.

**Example:** Before editing a paper, I needed to locate 12 different sections by searching for unique text strings in each. This required 12 separate `find_text_in_document` calls:

```
find_text_in_document(file, "Addressing Potential Counterarguments")  → para 384
find_text_in_document(file, "Some digital accessibility practitioners argue") → para 368
find_text_in_document(file, "Policy Implications") → para 366
find_text_in_document(file, "Preference for each condition correlated") → para 338
find_text_in_document(file, "Within-subject correlations were examined") → para 302
find_text_in_document(file, "A parametric repeated measures ANOVA confirmed") → para 246
find_text_in_document(file, "parametric paired-samples t-test") → para 232
find_text_in_document(file, "Audio descriptions of maps") → para 68
find_text_in_document(file, "touchscreen") → para 64
find_text_in_document(file, "Screen readers used by BLVIs simply state") → para 52
find_text_in_document(file, "Map Equivalent Purpose (MEP) Framework") → para 56
find_text_in_document(file, "Audiom, a web-based") → not found (needed alternate search)
```

A batch function would reduce these 12 calls to 1.

**Suggested API:**
```python
find_texts_in_document(filename: str, texts_to_find: list[str], match_case: bool = True) -> dict
# Returns a dict keyed by search string, each containing the same results as find_text_in_document.
# Example: find_texts_in_document("doc.docx", [
#     "Addressing Potential Counterarguments",
#     "Some digital accessibility practitioners argue",
#     "Policy Implications"
# ])
# Returns all 3 results in one call instead of 3 separate calls.
```

---

## Issue 7 [Bug]: `get_document_info` word count may differ from Word's built-in count

**Problem:** `get_document_info` reports a word count that may not match Microsoft Word's built-in word count. Word's "Word Count" dialog has options to include or exclude footnotes, textboxes, and headers/footers. The MCP server's count method (likely using `python-docx`) may count differently. This causes confusion when targeting a specific word count for journal submissions where the journal uses Word's count.

**Example:** After applying all edits to a paper, `get_document_info` reported 11,184 words. There is no way to verify whether this matches what Word would show in the status bar or the Word Count dialog (Review → Word Count). If the journal requires "under 10,400 words" and uses Word's count, a discrepancy could mean the paper is still over the limit despite the MCP server reporting it as under.

**Suggested fix:**
1. Document in the tool description which counting method is used and how it compares to Word's count
2. Consider adding a `count_mode` parameter (e.g., `"body_only"`, `"include_footnotes"`) to match Word's behavior
3. At minimum, add a note to the `get_document_info` response indicating whether footnotes/textboxes are included

---

## Issue 8 [Bug]: Empty paragraphs required as spacers in `replace_paragraph_range` — undocumented

**Problem:** Word documents commonly use empty paragraphs (`""`) and non-breaking space paragraphs (`"\u00a0"`) as spacing between content paragraphs. When using `replace_paragraph_range`, you must explicitly include empty string entries in the `new_paragraphs` array to preserve this spacing. This is not documented in the tool description and is easy to miss.

**Example:** A section has 4 content paragraphs each separated by empty spacer paragraphs (8 paragraphs total):

```
Index 338: "When spatial information is present, participants prefer..."   (content)
Index 339: ""                                                              (spacer)
Index 340: "Although the preference question focused specifically..."      (content)
Index 341: ""                                                              (spacer)
Index 342: "The apparent paradox of higher cognitive workload..."          (content)
Index 343: ""                                                              (spacer)
Index 344: "If visual maps serve as a benchmark, it is unrealistic..."     (content)
Index 345: ""                                                              (spacer)
```

To replace these with 3 new paragraphs while preserving spacing, the `new_paragraphs` array must include explicit empty strings:

```json
[
  "Preference for each condition correlated with the amount of spatial information...",
  "",
  "NASA-TLX scores followed the pattern: ITM highest, visual intermediate...",
  "",
  "The apparent paradox of higher workload alongside better performance..."
]
```

If you omit the empty strings, the content paragraphs will run together without spacing.

**Suggested fix:** Add a note in the `replace_paragraph_range` tool description explaining that empty paragraphs used as spacers in the original document must be explicitly included as `""` entries in the `new_paragraphs` array.

---

## Issue 9 [Bug]: `replace_paragraph_text` clears inline formatting (bold, italic)

**Problem:** When using `replace_paragraph_text` to replace a paragraph's content, all inline formatting (bold, italic, underline) within the original paragraph is lost. The replacement text becomes a single unformatted run, even when `preserve_style: true` is set. `preserve_style` only preserves the paragraph-level style (e.g., "Normal", "Heading 2"), not character-level formatting within the paragraph.

**Example:** A paragraph contains statistical text with italic formatting on variable names:

```
Original paragraph (index 246):
"A parametric repeated measures ANOVA confirmed these findings (*F*(2, 38) = 108.37,
*p* < .001). Both parametric and non-parametric tests reached the same conclusion,
demonstrating that the results are robust regardless of statistical approach."
```

Where `*F*` and `*p*` are rendered as italic "F" and italic "p" in Word.

After calling:
```
replace_paragraph_text(file, 246, "A parametric repeated measures ANOVA confirmed the
same pattern (F(2, 38) = 108.37, p < .001). Both map-based representations (visual and
ITM) substantially outperformed tables for spatial question answering among sighted
participants.")
```

The replacement text is entirely plain — "F" and "p" are no longer italic.

**Current Workaround:** After replacing text, use `format_text(filename, paragraph_index, start_pos, end_pos, italic=True)` to re-apply formatting. This requires calculating the exact character positions of each formatted span, which is tedious and error-prone.

**Suggested fix:** Support basic markdown-style formatting in replacement text (e.g., `*italic*`, `**bold**`), or accept a list of formatting runs alongside the text:

```python
# Option A: Markdown-style
replace_paragraph_text(file, 246, "A parametric ANOVA confirmed (*F*(2, 38) = 108.37, *p* < .001).")

# Option B: Formatting runs
replace_paragraph_text(file, 246, "A parametric ANOVA confirmed (F(2, 38) = 108.37, p < .001).",
    formatting_runs=[
        {"start": 41, "end": 42, "italic": True},  # F
        {"start": 60, "end": 61, "italic": True}    # p
    ])
```

---

## Issue 10 [Workaround]: Working backward to preserve paragraph indices

**Problem:** When making multiple edits to different parts of a document using paragraph indices, edits that add or remove paragraphs cause all subsequent indices to shift. If you edit from top to bottom, every edit invalidates the indices you found for sections below it.

**Example:** You need to make two edits:
1. Replace paragraph 368 (a single paragraph → single paragraph, no shift)
2. Replace paragraphs 380–383 (4 paragraphs → 1 paragraph, shifts everything below by -3)

If done in forward order:
```
# Edit 1: Replace paragraph 368 with new text (no index shift — same paragraph count)
replace_paragraph_text(file, 368, "New policy text...")

# Edit 2: Replace paragraphs 380-383 with one paragraph
# These indices are STILL CORRECT because Edit 1 didn't change paragraph count
replace_paragraph_range(file, 380, 383, ["New legislation text..."])
# This works in this case, but if Edit 1 had used replace_paragraph_range and changed
# the paragraph count, 380-383 would be WRONG.
```

If Edit 1 had replaced 1 paragraph with 3 (adding 2 paragraphs), then index 380 would actually be at 382 by the time Edit 2 runs.

**Workaround rule:** Always edit from the highest paragraph indices first and work backward:
```
# CORRECT order (backward — higher indices first):
replace_paragraph_range(file, 380, 383, ["New legislation text..."])  # edit higher indices first
replace_paragraph_text(file, 368, "New policy text...")                # then lower indices
```

This guarantees that earlier edits never affect the indices of later edits.

**Suggested fix:** This workaround would be unnecessary if the server supported text-anchor-based editing (find-and-replace on paragraph content) rather than index-based editing. Alternatively, `replace_paragraph_range` could accept a `paragraph_text_anchor` parameter to locate the target by content instead of index.

---

## Issue 11 [Workaround]: Deleting multiple paragraphs requires bottom-up sequential calls

**Problem:** Since `delete_paragraph_range` does not exist (see Issue 1), deleting multiple paragraphs requires calling `delete_paragraph` once per paragraph. Each deletion shifts all subsequent paragraph indices down by 1, so deletions must proceed from the highest index downward.

**Example:** Delete paragraphs 384–391 (the "Addressing Potential Counterarguments" section):

```
# Content to delete:
# 384: "### Addressing Potential Counterarguments"
# 385: ""
# 386: "A potential objection is that interactive text maps..."
# 387: ""
# 388: "Implementation costs are a valid concern..."
# 389: ""
# 390: "Legacy maps can be made accessible provided..."
# 391: ""

# Must delete from bottom up:
delete_paragraph(file, 391)  # after this, doc has indices 0-390
delete_paragraph(file, 390)  # after this, doc has indices 0-389
delete_paragraph(file, 389)
delete_paragraph(file, 388)
delete_paragraph(file, 387)
delete_paragraph(file, 386)
delete_paragraph(file, 385)
delete_paragraph(file, 384)  # section fully removed
```

If you delete from top down (starting at 384), after the first deletion the paragraph that was at 385 shifts to 384, and you'd delete the wrong content on the next call.

**Suggested fix:** Implement `delete_paragraph_range(filename, start_index, end_index)` per Issue 1.

---

## Issue 12 [Workaround]: Finding section boundaries requires many sequential reads

**Problem:** To determine the exact paragraph range of a section (heading + content until next heading), there is no single-call method. You must manually walk through paragraphs one at a time.

**Example:** To find the boundaries of "## Within-Subject Correlations":

```
# Step 1: Find the heading
find_text_in_document(file, "Within-subject correlations were examined")
# → paragraph_index: 302 (but this is the content, not the heading)

# Need to also check the heading:
get_paragraph_text_from_document(file, 300)
# → "## Within-Subject Correlations" (Heading 2) — this is the heading

get_paragraph_text_from_document(file, 301)
# → "" (empty spacer)

get_paragraph_text_from_document(file, 302)
# → "Within-subject correlations were examined to assess consistency..." (Normal, content)

get_paragraph_text_from_document(file, 303)
# → "" (empty spacer)

get_paragraph_text_from_document(file, 304)
# → "For sighted participants, performance correlations between..." (Normal, content)

get_paragraph_text_from_document(file, 305)
# → "" (empty spacer)

get_paragraph_text_from_document(file, 306)
# → "## Summary of Key Findings" (Heading 2 — STOP, next section found)

# Result: Section content is paragraphs 301-305, next heading at 306
```

This required 7 calls to map one section. A long section (e.g., 20+ paragraphs) requires 20+ calls.

**Suggested fix:** Implement `get_section_paragraphs(filename, heading_text)` per Issue 3, or `get_paragraph_range(filename, start_index, end_index)` per Issue 2 to at least batch the reads.

---

## Issue 13 [Workaround]: `replace_paragraph_range` style behavior is unclear

**Problem:** `replace_paragraph_text` has a `preserve_style: bool = true` parameter that keeps the original paragraph's formatting. However, `replace_paragraph_range` only has a single `style` parameter that applies one style to all new paragraphs. When `style` is not provided, it is undocumented what style the new paragraphs receive. In practice, new paragraphs appear to get "Normal" style, but this should be explicitly documented.

**Example:** Replacing paragraphs 302–305 (which are all "Normal" style) with new content:

```python
replace_paragraph_range(file, 302, 305, [
    "Within-subject correlations were examined to assess consistency...",
    "",
    "For sighted participants, performance correlations between conditions..."
])
# No style parameter provided — new paragraphs appear to inherit "Normal" but this is not documented.
```

If the replaced paragraphs had been "Heading 2" style, it's unclear whether the new paragraphs would inherit that style or default to "Normal".

**Suggested fix:** Add `preserve_style: bool = true` to `replace_paragraph_range`. When true, new paragraphs inherit the style of the paragraph at `start_index`. The existing `style` parameter would override this when explicitly set. Document the default behavior.
