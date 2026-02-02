# Changes — 2026-02-01

## Summary

Six bug fixes and new features were implemented across the Office-Word-MCP-Server codebase, merged into the `integration/all-fixes` branch. All changes include corresponding test coverage.

---

## Bug Fixes

### 1. Cross-run text matching in `search_and_replace`
**Commit:** `34ac48d`
**Files:** `word_document_server/utils/document_utils.py`

The `find_and_replace_text` function previously only matched text within individual XML runs. When Word splits text across multiple `<w:r>` elements (e.g., due to spell-check, formatting changes, or editing history), the old per-run search would silently miss matches.

Added `_replace_in_paragraph()` helper that builds a character-to-run map, locates the match across run boundaries, and surgically updates only the affected runs (prefix of first run, clear intermediates, suffix of last run).

### 2. Anchor matching normalization in `replace_block_between_manual_anchors`
**Commit:** `48f3a57`
**Files:** `word_document_server/utils/document_utils.py`

Anchor text comparisons were using simple `.strip()` which failed when documents contained non-breaking spaces, multiple spaces, or Unicode whitespace variants.

- Added `_normalize_text()` utility (NFKC normalization + whitespace collapse + strip).
- Changed all anchor matching to use two-pass strategy: exact normalized match first, then substring/contains fallback.
- Applied the same normalization to both start and end anchor detection.
- Pre-computed `qn('w:p')` and `qn('w:tbl')` tag constants (`_W_P`, `_W_TBL`) for cleaner element comparisons.

### 3. Header matching normalization in `replace_paragraph_block_below_header` and `delete_block_under_header`
**Commit:** `26a3428`
**Files:** `word_document_server/utils/document_utils.py`

Both header-matching functions suffered from the same whitespace sensitivity as anchors.

- Applied `_normalize_text()` to header matching in both `replace_paragraph_block_below_header` and `delete_block_under_header`.
- Added two-pass matching: exact normalized match, then contains-match restricted to heading-styled paragraphs.
- Added logging for contains-fallback matches to aid debugging.

---

## New Features

### 4. `replace_paragraph_text` tool — atomic paragraph replacement by index
**Commit:** `386bbf5`
**Files:** `word_document_server/utils/document_utils.py`, `word_document_server/tools/content_tools.py`, `word_document_server/main.py`

New MCP tool that replaces the text content of a single paragraph identified by its index. Clears all existing runs and writes new text into the first run to preserve character-level formatting. Optionally preserves or resets the paragraph style.

**Parameters:** `filename`, `paragraph_index`, `new_text`, `preserve_style` (default `True`)

### 5. `copy_style_from_index` parameter on `insert_line_or_paragraph_near_text`
**Commit:** `978d102`
**Files:** `word_document_server/utils/document_utils.py`, `word_document_server/tools/content_tools.py`, `word_document_server/main.py`

Added an optional `copy_style_from_index` parameter that copies character-level formatting (bold, italic, underline, font name, size, color) from a source paragraph's first run to the newly inserted paragraph's first run. This enables inserting paragraphs that visually match existing content without needing to specify individual formatting properties.

Added `_copy_run_formatting()` helper for the run-level property copy.

### 6. `replace_paragraph_range` tool — batch paragraph replacement
**Commit:** `62ce7da`
**Files:** `word_document_server/utils/document_utils.py`, `word_document_server/tools/content_tools.py`, `word_document_server/main.py`

New MCP tool that replaces a contiguous range of paragraphs (by start/end index, inclusive) with a new list of paragraphs in a single document save. Removes old paragraph XML elements in reverse order to preserve indices, then inserts new paragraphs at the correct position using XML sibling insertion.

**Parameters:** `filename`, `start_index`, `end_index`, `new_paragraphs` (list of strings), `style` (optional)

---

## Tests Added

All new functionality includes pytest-based tests under `tests/`:

| Test File | Covers |
|---|---|
| `tests/conftest.py` | Shared fixtures (temp document creation, paragraph seeding) |
| `tests/test_search_and_replace.py` | Cross-run search and replace |
| `tests/test_anchor_replace.py` | Normalized anchor matching |
| `tests/test_header_replace.py` | Normalized header matching |
| `tests/test_replace_paragraph.py` | `replace_paragraph_text` tool |
| `tests/test_insert_paragraph_formatting.py` | `copy_style_from_index` parameter |
| `tests/test_batch_replace.py` | `replace_paragraph_range` tool |

### 7. `delete_paragraph_range` tool — batch paragraph deletion
**Branch:** `feat/delete-paragraph-range`
**Files:** `document_utils.py`, `content_tools.py`, `main.py`
**Issues resolved:** 1, 10, 11

New MCP tool that deletes a contiguous range of paragraphs (by start/end index, inclusive) in a single operation. Removes XML elements in reverse order to preserve indices internally. Tool docstring includes guidance on working backward when making multiple range operations, eliminating the workarounds documented in Issues 10 and 11.

---

## Other Changes

- Added `pytest` as a dev dependency in `pyproject.toml`
- Updated `uv.lock` with new dependency resolution
- Created implementation plan documents under `plans/` (later cleaned up from the branch)
