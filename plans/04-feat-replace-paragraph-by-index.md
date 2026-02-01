# Plan 4: Add `replace_paragraph_text` Tool (Replace by Index)

**Branch:** `feat/replace-paragraph-by-index`
**Depends on:** Plan 0 (shared test infrastructure)

## Context

- **Repository:** `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
- **Fork origin:** `https://github.com/frastlin/Office-Word-MCP-Server.git`
- **Upstream:** `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Problem

There is no single tool to replace the text of an existing paragraph by index. The current workaround requires 3 error-prone steps:
1. `get_paragraph_text_from_document(index=N)` — confirm content
2. `insert_line_or_paragraph_near_text(target_paragraph_index=N, position="before", line_text="new text")` — insert replacement
3. `delete_paragraph(paragraph_index=N+1)` — delete original (index shifted +1)

This is fragile because the insert shifts indices, the new paragraph gets default style, and partial failures leave duplicates.

## Files to modify
- `word_document_server/utils/document_utils.py` — add `replace_paragraph_text()` function
- `word_document_server/tools/content_tools.py` — add async wrapper + update imports
- `word_document_server/main.py` — register new MCP tool in `register_tools()`

## Files to create
- `tests/conftest.py` (see Plan 0)
- `tests/test_replace_paragraph.py`

## Architecture

- **Utility function** (synchronous) in `document_utils.py`
- **Async wrapper** in `content_tools.py` with file validation
- **MCP tool registration** in `main.py` inside `register_tools()` with `ToolAnnotations(title="Replace Paragraph Text", destructiveHint=True)`

## Tests (write first, verify they fail)

| Test | Input | Expected |
|------|-------|----------|
| Basic replacement | 3 paragraphs, replace index 1 | Text changed, others intact |
| Preserves style | Heading 2 paragraph, preserve_style=True | Style stays "Heading 2" |
| Invalid index | Index 5 on 2-paragraph doc | Error message |
| Paragraph count unchanged | Replace middle paragraph | len(paragraphs) same before/after |
| preserve_style=False | Heading paragraph, flag off | Style becomes "Normal" |

## Implementation

### Utility function in `document_utils.py`
```python
def replace_paragraph_text(doc_path: str, paragraph_index: int, new_text: str, preserve_style: bool = True) -> str:
    """Replace the text of a paragraph at a given index, optionally preserving style."""
    import os
    if not os.path.exists(doc_path):
        return f"Document {doc_path} does not exist"

    try:
        doc = Document(doc_path)
        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return f"Invalid paragraph index: {paragraph_index}. Document has {len(doc.paragraphs)} paragraphs."

        para = doc.paragraphs[paragraph_index]
        old_style = para.style

        # Clear all runs
        for run in para.runs:
            run.text = ""

        # Set text on first run (preserves its formatting) or add new run
        if para.runs:
            para.runs[0].text = new_text
        else:
            para.add_run(new_text)

        if preserve_style and old_style:
            para.style = old_style
        elif not preserve_style:
            para.style = doc.styles["Normal"]

        doc.save(doc_path)
        return f"Paragraph at index {paragraph_index} replaced successfully."
    except Exception as e:
        return f"Failed to replace paragraph: {str(e)}"
```

### Async wrapper in `content_tools.py`
```python
async def replace_paragraph_text_tool(filename: str, paragraph_index: int, new_text: str, preserve_style: bool = True) -> str:
    """Replace text of a specific paragraph by index."""
    filename = ensure_docx_extension(filename)
    if not os.path.exists(filename):
        return f"Document {filename} does not exist"
    is_writeable, error_message = check_file_writeable(filename)
    if not is_writeable:
        return f"Cannot modify document: {error_message}."
    return replace_paragraph_text(filename, paragraph_index, new_text, preserve_style)
```

### MCP registration in `main.py`
Add inside `register_tools()`, after the `delete_paragraph` registration:
```python
@mcp.tool(
    annotations=ToolAnnotations(
        title="Replace Paragraph Text",
        destructiveHint=True,
    ),
)
def replace_paragraph_text(filename: str, paragraph_index: int, new_text: str, preserve_style: bool = True):
    """Replace the text of a specific paragraph by index, optionally preserving its style."""
    return content_tools.replace_paragraph_text_tool(filename, paragraph_index, new_text, preserve_style)
```

Update imports in `content_tools.py`:
```python
from word_document_server.utils.document_utils import ..., replace_paragraph_text
```

## Git workflow
```bash
git checkout -b feat/replace-paragraph-by-index
# 1. Create tests/conftest.py (Plan 0)
# 2. Create tests/test_replace_paragraph.py
# 3. Run tests, verify they fail: uv run pytest tests/test_replace_paragraph.py -v
# 4. Implement in document_utils.py, content_tools.py, main.py
# 5. Run tests, verify they pass: uv run pytest tests/test_replace_paragraph.py -v
# 6. Run full suite: uv run pytest tests/ -v
git add tests/conftest.py tests/test_replace_paragraph.py \
  word_document_server/utils/document_utils.py \
  word_document_server/tools/content_tools.py \
  word_document_server/main.py
git commit -m "feat: add replace_paragraph_text tool for atomic paragraph replacement by index"
git push -u origin feat/replace-paragraph-by-index
gh pr create --repo GongRzhe/Office-Word-MCP-Server \
  --title "feat: add replace-paragraph-by-index tool" \
  --body "$(cat <<'EOF'
## Summary
- New tool `replace_paragraph_text` replaces text at a given paragraph index in one atomic operation
- Eliminates the error-prone 3-step workflow (get, insert before, delete at shifted index)
- Preserves paragraph style by default (configurable with preserve_style=False)
- Preserves first run's formatting

## Test plan
- [ ] Basic text replacement works
- [ ] Style preserved when preserve_style=True
- [ ] Invalid index returns error message
- [ ] Paragraph count unchanged after replacement
- [ ] preserve_style=False resets to Normal

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

## Verification
```bash
uv run pytest tests/test_replace_paragraph.py -v
uv run pytest tests/ -v
```
