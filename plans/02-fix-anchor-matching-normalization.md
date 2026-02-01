# Plan 2: Fix `replace_block_between_manual_anchors` Anchor Matching

**Branch:** `fix/anchor-matching-normalization`
**Depends on:** Plan 0 (shared test infrastructure)

## Context

- **Repository:** `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
- **Fork origin:** `https://github.com/frastlin/Office-Word-MCP-Server.git`
- **Upstream:** `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Files to modify
- `word_document_server/utils/document_utils.py` — function `replace_block_between_manual_anchors` (lines 531-618)

## Files to create
- `tests/conftest.py` (see Plan 0)
- `tests/test_anchor_replace.py`

## Root cause

Anchor matching uses exact string comparison: `p_text == start_anchor_text.strip()`. This fails when:
- Document contains NBSP (U+00A0) instead of regular spaces
- Document contains ZWSP (U+200B) or other invisible Unicode
- Extra leading/trailing whitespace exists
- Auto-numbering adds prefixes like "1. " to the paragraph text

## Architecture

- The utility function in `document_utils.py` is synchronous
- The tool wrapper in `content_tools.py` is async and delegates to it
- No changes needed to the wrapper or MCP registration — only the matching logic changes

## Tests (write first, verify they fail)

| Test | Input | Expected |
|------|-------|----------|
| Exact match works | Clean anchors with content between | Content replaced, old content gone |
| NBSP in document | Anchor has U+00A0 instead of space | Still matches |
| Extra whitespace | Leading/trailing spaces on anchor paragraph | Still matches |
| Contains/substring fallback | "1. --- START ANCHOR ---" with auto-number prefix | Still matches via fallback |
| Anchor not found | Non-existent anchor text | Error message with "not found" |

## Implementation

### Add module-level helper (reused by Plan 3 / Bug 3)
```python
import unicodedata
import re
import logging

logger = logging.getLogger(__name__)

def _normalize_text(s: str) -> str:
    """Normalize text for reliable matching: NFKC normalize, collapse whitespace, strip."""
    s = unicodedata.normalize("NFKC", s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()
```

### Modify start anchor search (lines 554-563)

**Two-pass matching:**

1. **Pass 1 — exact normalized:** `_normalize_text(p_text) == _normalize_text(start_anchor_text)`
2. **Pass 2 — contains fallback (if pass 1 fails):** `_normalize_text(start_anchor_text) in _normalize_text(p_text)`
3. Log which pass succeeded via `logger.info()`

### Modify end anchor search (lines 567-578)

Same two-pass approach for `end_anchor_text`.

### Modify post-deletion re-find (line 606)

Currently: `para.text.strip() == start_anchor_text.strip()`
Change to: `_normalize_text(para.text) == _normalize_text(start_anchor_text)` with contains fallback.

## Git workflow
```bash
git checkout -b fix/anchor-matching-normalization
# 1. Create tests/conftest.py (Plan 0)
# 2. Create tests/test_anchor_replace.py
# 3. Run tests, verify they fail: uv run pytest tests/test_anchor_replace.py -v
# 4. Implement fix in document_utils.py
# 5. Run tests, verify they pass: uv run pytest tests/test_anchor_replace.py -v
# 6. Run full suite: uv run pytest tests/ -v
git add tests/conftest.py tests/test_anchor_replace.py word_document_server/utils/document_utils.py
git commit -m "fix: normalize whitespace and support substring matching for anchor detection"
git push -u origin fix/anchor-matching-normalization
gh pr create --repo GongRzhe/Office-Word-MCP-Server \
  --title "fix: replace_block_between_manual_anchors handles NBSP and auto-numbering" \
  --body "$(cat <<'EOF'
## Summary
- Anchor text matching now normalizes Unicode (NFKC) and collapses whitespace before comparison
- Falls back to substring/contains matching when exact normalized match fails
- Handles NBSP (U+00A0), ZWSP (U+200B), auto-numbering prefixes, and extra whitespace
- Added logging for anchor search diagnostics

## Test plan
- [ ] Exact match still works (regression)
- [ ] NBSP in document matches regular-space anchor
- [ ] Extra leading/trailing whitespace handled
- [ ] Auto-numbering prefix handled via contains fallback
- [ ] Non-existent anchor returns clear error

🤖 Generated with [Claude Code](https://claude.com/claude-code)
EOF
)"
```

## Verification
```bash
uv run pytest tests/test_anchor_replace.py -v
uv run pytest tests/ -v
```
