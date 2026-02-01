# Word Document MCP Server — Bugs, Shortcomings, and Workarounds

**Date:** 2026-02-01
**Context:** Encountered while editing `tvm-paper.docx`, an academic paper compiled from pandoc markdown to .docx format.

---

## Bug 1: `search_and_replace` Fails Silently on Cross-Run Text

### The Problem

`search_and_replace` cannot match text that spans multiple XML runs (`<w:r>` elements) inside a Word document. It returns "No occurrences found" even when the text provably exists in the document (confirmed by `find_text_in_document` which *does* match across run boundaries).

### Root Cause

Word's `.docx` format stores text as XML. Each paragraph (`<w:p>`) contains one or more **runs** (`<w:r>`), where each run holds a contiguous string with consistent formatting. When formatting changes mid-paragraph — bold, italic, hyperlink, citation bracket, font change, etc. — Word splits the text into separate runs.

For example, this paragraph text:
```
Evaluations have found high cognitive load [@zong2022rich; @hennig2017accessible]
```

May be stored internally as:
```xml
<w:p>
  <w:r><w:t>Evaluations have found high cognitive load </w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>[@zong2022rich</w:t></w:r>
  <w:r><w:t>; </w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>@hennig2017accessible]</w:t></w:r>
</w:p>
```

`find_text_in_document` joins all run text within a paragraph before searching, so it finds the full string. But `search_and_replace` appears to search within individual runs, so any find_text string that crosses a run boundary will fail.

### Examples of Failure

| find_text | Result | Why |
|-----------|--------|-----|
| `"Evaluations have found high cognitive load"` (40 chars, single run) | **Success** | Falls within one run |
| `"found high cognitive load [@zong2022rich"` (40 chars, crosses run boundary) | **Failure** | Spans the plain-text run and the hyperlink run |
| Any paragraph-length string (200+ chars) | **Almost always fails** | Virtually guaranteed to cross at least one run boundary |

### How to Detect

If `find_text_in_document` finds your string but `search_and_replace` says "No occurrences found," you've hit this bug.

### Recommended Fix for the MCP Server

The server's `search_and_replace` implementation should:

1. **Join all run text within each paragraph** into a single string (same as `find_text_in_document` does).
2. **Perform the find/replace on the joined string.**
3. **Rebuild the paragraph's runs** from the replaced string. The simplest approach: clear all existing runs and create a single new run with the paragraph's default formatting. A more sophisticated approach would preserve run-level formatting by mapping character offsets back to the original runs.

Alternatively, use `python-docx`'s lower-level lxml access:
```python
from docx.oxml.ns import qn

def search_and_replace_across_runs(paragraph, find_text, replace_text):
    # Join all run texts
    full_text = ''.join(run.text for run in paragraph.runs)
    if find_text not in full_text:
        return False

    new_text = full_text.replace(find_text, replace_text)

    # Clear all runs except first, put full text in first run
    for i, run in enumerate(paragraph.runs):
        if i == 0:
            run.text = new_text
        else:
            run.text = ''
    return True
```

**Trade-off:** This loses run-level formatting (bold/italic on specific words). A production fix should preserve formatting by tracking character offset ranges per run and re-splitting accordingly.

---

## Bug 2: `replace_block_between_manual_anchors` — Unreliable Anchor Matching

### The Problem

`replace_block_between_manual_anchors` was tested to replace content between two section headings. It sometimes fails to find anchors even when they are exact matches of heading text. The behavior is inconsistent — sometimes it works, sometimes it doesn't.

### Likely Cause

The anchor matching logic may be doing exact full-paragraph text comparison, but headings in .docx files often have invisible characters (non-breaking spaces, zero-width spaces) or numbering prefixes from auto-numbering styles that aren't visible to the user but affect string matching.

### Recommended Fix

- Normalize whitespace (strip leading/trailing, collapse internal whitespace) before comparing anchor text.
- Optionally support substring/contains matching in addition to exact matching.
- Log the actual paragraph text being compared so users can debug mismatches.

---

## Bug 3: `replace_paragraph_block_below_header` — Similar Anchor Issues

### The Problem

Same class of issue as Bug 2. The tool matches a heading by text, but the match can fail due to invisible formatting characters or paragraph-style differences.

### Recommended Fix

Same as Bug 2. Additionally, match on heading style (Heading 1, Heading 2, etc.) in addition to text content, so minor text discrepancies don't prevent matching.

---

## Shortcoming 1: No "Replace Paragraph by Index" Tool

### The Problem

There is no single tool to replace the text of an existing paragraph by index. The workaround requires three steps:

1. `get_paragraph_text_from_document(index=N)` — confirm content
2. `insert_line_or_paragraph_near_text(target_paragraph_index=N, position="before", line_text="new text")` — insert replacement
3. `delete_paragraph(paragraph_index=N+1)` — delete original (index shifted by +1 due to insert)

This is error-prone because:
- The insert shifts all subsequent paragraph indices by +1.
- If the insert succeeds but the delete fails (or is interrupted), you have a duplicate paragraph.
- The new paragraph gets a default style, not the style of the replaced paragraph.

### Recommended Fix

Add a `replace_paragraph_text(filename, paragraph_index, new_text, preserve_style=True)` tool that:
1. Reads the existing paragraph's style
2. Replaces all runs with the new text (in a single run, or preserving formatting structure)
3. Preserves the paragraph's style, alignment, spacing, etc.

---

## Shortcoming 2: `insert_line_or_paragraph_near_text` Doesn't Preserve Source Formatting

### The Problem

When inserting a new paragraph, the tool creates it with default formatting. If the original paragraph had a specific style (e.g., "Body Text", "Normal" with custom font/size), the inserted paragraph may look different.

### Recommended Fix

- Add an optional `copy_style_from_index` parameter that copies the paragraph style and run formatting from an existing paragraph.
- Or add `style` parameter support (this may already exist via `line_style`; documentation is unclear on which style names are accepted).

---

## Shortcoming 3: No Batch Operations

### The Problem

Replacing a section (e.g., 5 paragraphs under a heading) requires 5 separate insert + 5 separate delete calls = 10 API round trips, with paragraph index recalculation between each pair.

### Recommended Fix

Add a `replace_paragraph_range(filename, start_index, end_index, new_paragraphs: list[str], style="Normal")` tool that atomically replaces a contiguous range of paragraphs with new ones.

---

## General Workaround Strategy for Editing Academic .docx Files

Given these limitations, the most reliable approach is:

1. **Use `find_text_in_document`** to locate paragraphs by unique short text snippets.
2. **Use `get_paragraph_text_from_document`** to read full paragraph content and confirm indices.
3. **Use `insert_line_or_paragraph_near_text` + `delete_paragraph`** as the primary replacement method (insert before, then delete original at index+1).
4. **Work backward** (highest paragraph index first) so insertions/deletions don't shift indices for subsequent operations.
5. **Use `search_and_replace` only for short strings** (<80 characters) that are likely to fall within a single XML run — avoid strings containing citation brackets, italic markers, or other formatting transitions.
6. **Verify each operation** with `get_paragraph_text_from_document` immediately after.
7. **Always create a backup** with `copy_document` before starting.
