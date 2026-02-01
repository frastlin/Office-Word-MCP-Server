# Plan 1: Fix `search_and_replace` Cross-Run Text Failure

**Branch:** `fix/search-and-replace-cross-run`
**Depends on:** Plan 0 (shared test infrastructure)

## Context

- **Repository:** `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
- **Fork origin:** `https://github.com/frastlin/Office-Word-MCP-Server.git`
- **Upstream:** `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Files to modify
- `word_document_server/utils/document_utils.py` — function `find_and_replace_text` (lines 138-177)

## Files to create
- `tests/conftest.py` (see Plan 0)
- `tests/test_search_and_replace.py`

## Root cause

`find_and_replace_text()` checks `old_text in para.text` (which joins runs via python-docx's `.text` property) but then iterates individual `run.text` values: `if old_text in run.text: run.text = run.text.replace(...)`. Text spanning multiple `<w:r>` XML elements is never found in any single run, so the replacement silently does nothing.

## Architecture

- Utility functions: `word_document_server/utils/document_utils.py` (synchronous)
- Tool wrappers: `word_document_server/tools/content_tools.py` (async)
- MCP registration: `word_document_server/main.py` inside `register_tools()`
- The tool wrapper `search_and_replace` in content_tools.py calls `find_and_replace_text` from document_utils.py — only the utility function needs changing.

## Tests (write first, verify they fail)

| Test | Input | Expected |
|------|-------|----------|
| Single-run replace | "Hello World" in one run, replace "Hello" → "Goodbye" | "Goodbye World", count=1 |
| Cross-run replace | "Hello " + "World" in two runs, replace "Hello World" → "Goodbye Earth" | Joined text = "Goodbye Earth", count>=1 |
| Preserves first-run formatting | Bold run + normal run, replace joined text | First run still bold |
| Table cell cross-run | Cross-run text in table cell | Replacement succeeds |
| No match returns 0 | Search for nonexistent text | count=0 |
| Multiple occurrences | Two paragraphs each with cross-run "Hello World" | Both replaced, count=2 |
| TOC paragraphs skipped | TOC-styled paragraph with target text | count=0, text unchanged |

## Implementation

Extract a helper `_replace_in_paragraph(para, old_text, new_text) -> int`:

1. **Build a character-to-run map:** For each char position in `para.text`, record `(run_index, char_offset_in_run)`
2. **Find first occurrence** of `old_text` in the joined text using `full_text.find(old_text)`
3. **Identify which runs** the match spans: `start_run_idx` through `end_run_idx`
4. **If single run:** Do simple in-run replacement (prefix + new_text + suffix within that run)
5. **If multi-run:**
   - First run: `prefix + new_text` (keeps text before match, appends replacement)
   - Intermediate runs: clear `.text = ""`
   - Last run: trim to keep only the suffix after the match
6. **Loop** while `old_text in para.text` for multiple occurrences in same paragraph

This preserves the first matched run's formatting (bold, italic, font) on the replacement text. Empty runs are harmless in .docx XML.

Replace the existing `find_and_replace_text` function body to use `_replace_in_paragraph` for both the paragraph loop and the table cell loop. Keep TOC skip logic unchanged.

## Git workflow
```bash
git checkout -b fix/search-and-replace-cross-run
# 1. Create tests/conftest.py (Plan 0)
# 2. Create tests/test_search_and_replace.py
# 3. Run tests, verify they fail: uv run pytest tests/test_search_and_replace.py -v
# 4. Implement fix in document_utils.py
# 5. Run tests, verify they pass: uv run pytest tests/test_search_and_replace.py -v
# 6. Run full suite: uv run pytest tests/ -v
git add tests/conftest.py tests/test_search_and_replace.py word_document_server/utils/document_utils.py
git commit -m "fix: handle cross-run text matching in search_and_replace"
git push -u origin fix/search-and-replace-cross-run
gh pr create --repo GongRzhe/Office-Word-MCP-Server \
  --title "fix: search_and_replace now works when target text spans multiple runs" \
  --body "$(cat <<'EOF'
## Summary
- Fixed find_and_replace_text() to handle text spanning multiple XML runs within a paragraph
- Old implementation checked para.text (joined) but searched individual runs, missing cross-run matches
- New implementation builds a character-to-run map, finds matches in joined text, modifies only affected runs
- First matched run's formatting is preserved on replacement text
- Works for both body paragraphs and table cells
- TOC paragraph skipping preserved

## Test plan
- [ ] Single-run replacement (regression)
- [ ] Cross-run replacement
- [ ] First-run formatting preservation
- [ ] Table cell cross-run replacement
- [ ] No-match returns zero
- [ ] Multiple occurrences across paragraphs
- [ ] TOC paragraphs skipped

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

## Verification
```bash
uv run pytest tests/test_search_and_replace.py -v
uv run pytest tests/ -v
```
