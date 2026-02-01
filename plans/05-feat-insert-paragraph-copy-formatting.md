# Plan 5: `insert_line_or_paragraph_near_text` Formatting Preservation

**Branch:** `feat/insert-paragraph-copy-formatting`
**Depends on:** Plan 0 (shared test infrastructure)

## Context

- **Repository:** `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
- **Fork origin:** `https://github.com/frastlin/Office-Word-MCP-Server.git`
- **Upstream:** `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Problem

When inserting a new paragraph, the tool creates it with default formatting. If the original paragraph had specific run-level formatting (bold, specific font, font size, color), the inserted paragraph looks different. The `line_style` parameter sets paragraph style but not run-level character formatting.

## Files to modify
- `word_document_server/utils/document_utils.py` — modify `insert_line_or_paragraph_near_text()` signature + logic (lines 243-295)
- `word_document_server/tools/content_tools.py` — update wrapper signature
- `word_document_server/main.py` — update tool registration signature

## Files to create
- `tests/conftest.py` (see Plan 0)
- `tests/test_insert_paragraph_formatting.py`

## Architecture

- Add helper `_copy_run_formatting(source_run, target_run)` in `document_utils.py`
- Add optional `copy_style_from_index: int = None` parameter to existing function
- Update async wrapper and MCP registration to expose the new parameter
- Fully backward compatible — omitting the parameter preserves existing behavior

## Tests (write first, verify they fail)

| Test | Input | Expected |
|------|-------|----------|
| Copies formatting from source | Bold Arial 14pt at index 0, copy_style_from_index=0, insert after | New paragraph has bold, Arial, 14pt |
| Without parameter, default behavior | No copy_style_from_index | Works as before, text inserted (regression) |

## Implementation

### Helper function in `document_utils.py`
```python
def _copy_run_formatting(source_run, target_run):
    """Copy character-level formatting from source run to target run."""
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    if source_run.font.name:
        target_run.font.name = source_run.font.name
    if source_run.font.size:
        target_run.font.size = source_run.font.size
    if source_run.font.color and source_run.font.color.rgb:
        target_run.font.color.rgb = source_run.font.color.rgb
```

### Modify `insert_line_or_paragraph_near_text` (line 243)

Add parameter to signature:
```python
def insert_line_or_paragraph_near_text(doc_path, target_text=None, line_text="",
    position='after', line_style=None, target_paragraph_index=None,
    copy_style_from_index=None):  # NEW PARAMETER
```

After creating the new paragraph (line 284), add:
```python
new_para = doc.add_paragraph(line_text, style=style)

# Copy run formatting if requested
if copy_style_from_index is not None:
    if 0 <= copy_style_from_index < len(doc.paragraphs):
        source_para = doc.paragraphs[copy_style_from_index]
        if source_para.runs and new_para.runs:
            _copy_run_formatting(source_para.runs[0], new_para.runs[0])
```

### Update wrapper in `content_tools.py`
```python
async def insert_line_or_paragraph_near_text_tool(filename, target_text=None,
    line_text="", position='after', line_style=None, target_paragraph_index=None,
    copy_style_from_index=None):
    ...
    return insert_line_or_paragraph_near_text(filename, target_text, line_text,
        position, line_style, target_paragraph_index, copy_style_from_index)
```

### Update MCP registration in `main.py`
Add `copy_style_from_index: int = None` to the registered function signature.

## Git workflow
```bash
git checkout -b feat/insert-paragraph-copy-formatting
# 1. Create tests/conftest.py (Plan 0) if not present
# 2. Create tests/test_insert_paragraph_formatting.py
# 3. Run tests, verify they fail: uv run pytest tests/test_insert_paragraph_formatting.py -v
# 4. Implement in document_utils.py, content_tools.py, main.py
# 5. Run tests, verify they pass: uv run pytest tests/test_insert_paragraph_formatting.py -v
# 6. Run full suite: uv run pytest tests/ -v
git add tests/conftest.py tests/test_insert_paragraph_formatting.py \
  word_document_server/utils/document_utils.py \
  word_document_server/tools/content_tools.py \
  word_document_server/main.py
git commit -m "feat: add copy_style_from_index parameter to insert_line_or_paragraph_near_text"
git push -u origin feat/insert-paragraph-copy-formatting
gh pr create --repo GongRzhe/Office-Word-MCP-Server \
  --title "feat: insert_line_or_paragraph_near_text can copy run formatting from source" \
  --body "$(cat <<'EOF'
## Summary
- Added optional `copy_style_from_index` parameter to insert_line_or_paragraph_near_text
- When provided, copies bold/italic/underline/font-name/font-size/color from first run of source paragraph
- Backward compatible: omitting the parameter preserves existing behavior

## Test plan
- [ ] Formatting copied from source paragraph when copy_style_from_index provided
- [ ] Default behavior unchanged when parameter omitted

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

## Verification
```bash
uv run pytest tests/test_insert_paragraph_formatting.py -v
uv run pytest tests/ -v
```
