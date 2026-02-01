# Plan 6: Add `replace_paragraph_range` Batch Tool

**Branch:** `feat/replace-paragraph-range`
**Depends on:** Plan 0 (shared test infrastructure)

## Context

- **Repository:** `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
- **Fork origin:** `https://github.com/frastlin/Office-Word-MCP-Server.git`
- **Upstream:** `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Problem

Replacing a section (e.g., 5 paragraphs under a heading) requires 5 separate insert + 5 separate delete calls = 10 API round trips, with paragraph index recalculation between each pair. This is slow, error-prone, and can leave the document in an inconsistent state if interrupted.

## Files to modify
- `word_document_server/utils/document_utils.py` — add `replace_paragraph_range()` function
- `word_document_server/tools/content_tools.py` — add async wrapper
- `word_document_server/main.py` — register new MCP tool

## Files to create
- `tests/conftest.py` (see Plan 0)
- `tests/test_batch_replace.py`

## Architecture

- **Utility function** (synchronous) in `document_utils.py`
- **Async wrapper** in `content_tools.py` with file validation
- **MCP tool registration** in `main.py` inside `register_tools()` with `ToolAnnotations(title="Replace Paragraph Range", destructiveHint=True)`

## Tests (write first, verify they fail)

| Test | Input | Expected |
|------|-------|----------|
| Same count replacement | [A,B,C,D,E] replace indices [1,3] with [X,Y,Z] | [A,X,Y,Z,E] |
| Fewer replacements | [A,B,C,D,E] replace indices [1,3] with [X] | [A,X,E] |
| More replacements | [A,B,C,D,E] replace indices [1,2] with [X,Y,Z,W] | [A,X,Y,Z,W,D,E] |
| Invalid range | start=3, end=5 on 2-paragraph doc | Error message |
| Style parameter | Replace with style="Heading 1" | New paragraphs have Heading 1 style |
| Single paragraph range | start=1, end=1 with [X] | [A,X,C] |

## Implementation

### Utility function in `document_utils.py`
```python
def replace_paragraph_range(doc_path: str, start_index: int, end_index: int,
                            new_paragraphs: list, style: str = None) -> str:
    """Replace paragraphs from start_index to end_index (inclusive) with new_paragraphs.

    Args:
        doc_path: Path to the document
        start_index: First paragraph index to replace (inclusive)
        end_index: Last paragraph index to replace (inclusive)
        new_paragraphs: List of text strings for new paragraphs
        style: Optional style name for new paragraphs (defaults to Normal)
    """
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"

    try:
        doc = Document(doc_path)
        total = len(doc.paragraphs)

        if start_index < 0 or end_index >= total or start_index > end_index:
            return f"Invalid range [{start_index}, {end_index}]. Document has {total} paragraphs (0-{total-1})."

        # Get anchor element (paragraph before start_index)
        if start_index > 0:
            anchor_element = doc.paragraphs[start_index - 1]._element
        else:
            anchor_element = None

        # Remove paragraphs in range (reverse to preserve indices)
        for i in range(end_index, start_index - 1, -1):
            p = doc.paragraphs[i]._p
            p.getparent().remove(p)

        # Insert new paragraphs
        style_to_use = style or "Normal"
        body = doc.element.body

        prev_element = anchor_element
        for text in new_paragraphs:
            new_para = doc.add_paragraph(text, style=style_to_use)
            if prev_element is not None:
                prev_element.addnext(new_para._element)
            else:
                body.insert(0, new_para._element)
            prev_element = new_para._element

        doc.save(doc_path)
        removed = end_index - start_index + 1
        return f"Replaced {removed} paragraph(s) (indices {start_index}-{end_index}) with {len(new_paragraphs)} new paragraph(s)."
    except Exception as e:
        return f"Failed to replace paragraph range: {str(e)}"
```

### Async wrapper in `content_tools.py`
```python
async def replace_paragraph_range_tool(filename: str, start_index: int, end_index: int,
                                        new_paragraphs: list, style: str = None) -> str:
    """Replace a range of paragraphs in a single operation."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."
    return replace_paragraph_range(filename, start_index, end_index, new_paragraphs, style)
```

### MCP registration in `main.py`
```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Replace Paragraph Range",
        destructiveHint=True,
    ),
)
def replace_paragraph_range(filename: str, start_index: int, end_index: int,
                            new_paragraphs: list[str], style: str = None):
    """Replace a range of paragraphs (start to end index inclusive) with new paragraphs in a single operation."""
    return content_tools.replace_paragraph_range_tool(filename, start_index, end_index, new_paragraphs, style)
```

## Git workflow
```bash
git checkout -b feat/replace-paragraph-range
# 1. Create tests/conftest.py (Plan 0)
# 2. Create tests/test_batch_replace.py
# 3. Run tests, verify they fail: uv run pytest tests/test_batch_replace.py -v
# 4. Implement in document_utils.py, content_tools.py, main.py
# 5. Run tests, verify they pass: uv run pytest tests/test_batch_replace.py -v
# 6. Run full suite: uv run pytest tests/ -v
git add tests/conftest.py tests/test_batch_replace.py \
  word_document_server/utils/document_utils.py \
  word_document_server/tools/content_tools.py \
  word_document_server/main.py
git commit -m "feat: add replace_paragraph_range tool for batch paragraph replacement"
git push -u origin feat/replace-paragraph-range
gh pr create --repo GongRzhe/Office-Word-MCP-Server \
  --title "feat: add batch replace_paragraph_range tool" \
  --body "$(cat <<'EOF'
## Summary
- New tool `replace_paragraph_range` replaces a contiguous range of paragraphs in one atomic operation
- Eliminates N individual get/insert/delete calls with manual index recalculation
- Supports replacing N paragraphs with M paragraphs (shrink, expand, or same count)
- Optional style parameter applies to all new paragraphs

## Test plan
- [ ] Replace range with same number of paragraphs
- [ ] Replace range with fewer paragraphs (shrink)
- [ ] Replace range with more paragraphs (expand)
- [ ] Invalid range returns error
- [ ] Style parameter applies to new paragraphs
- [ ] Single-paragraph range works

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

## Verification
```bash
uv run pytest tests/test_batch_replace.py -v
uv run pytest tests/ -v
```
