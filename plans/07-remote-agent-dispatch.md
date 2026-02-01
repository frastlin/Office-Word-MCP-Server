# Plan 7: Remote Claude Coding Agent Dispatch

Instructions for dispatching Plans 1-6 to remote Claude coding agents on claude.ai.

## Dependency graph
```
conftest.py (Plan 0 — included as first step in each branch)
    |
    +-- Plan 1: Bug 1 (cross-run replace)       [independent]
    +-- Plan 2: Bug 2 (anchor normalization)     [independent]
    +-- Plan 3: Bug 3 (header normalization)     [depends on Plan 2 for _normalize_text]
    +-- Plan 4: Shortcoming 1 (replace by index) [independent]
    +-- Plan 5: Shortcoming 2 (copy formatting)  [independent]
    +-- Plan 6: Shortcoming 3 (batch replace)    [independent]
```

## Recommended phasing

### Phase 1 — Launch 4 agents in parallel

| Agent | Task | Branch | Plan file |
|-------|------|--------|-----------|
| A | conftest.py + Bug 1 | `fix/search-and-replace-cross-run` | `01-fix-search-and-replace-cross-run.md` |
| B | conftest.py + Bug 2 | `fix/anchor-matching-normalization` | `02-fix-anchor-matching-normalization.md` |
| C | conftest.py + Shortcoming 1 | `feat/replace-paragraph-by-index` | `04-feat-replace-paragraph-by-index.md` |
| D | conftest.py + Shortcoming 3 | `feat/replace-paragraph-range` | `06-feat-replace-paragraph-range.md` |

### Phase 2 — After Agent B completes

| Agent | Task | Branch | Plan file |
|-------|------|--------|-----------|
| E | Bug 3 (branch from Bug 2's branch) | `fix/header-matching-normalization` | `03-fix-header-matching-normalization.md` |
| F | Shortcoming 2 | `feat/insert-paragraph-copy-formatting` | `05-feat-insert-paragraph-copy-formatting.md` |

### Phase 3 — Integration agent

- Create integration branch from `main`
- Merge all 6 feature branches
- Resolve conflicts (most likely in `document_utils.py` imports and `conftest.py`)
- Run full test suite: `uv run pytest tests/ -v`
- Verify no regressions in existing `tests/test_convert_to_pdf.py`

## Context to provide each agent

Each agent prompt should include:

1. **Repository info:**
   - Path: `C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server`
   - Fork origin: `https://github.com/frastlin/Office-Word-MCP-Server.git`
   - Upstream: `https://github.com/GongRzhe/Office-Word-MCP-Server`

2. **The specific plan file** for the bug/feature (copy-paste the full plan)

3. **The conftest.py spec** from `00-shared-test-infrastructure.md`

4. **Architectural pattern:**
   - Utility functions: `word_document_server/utils/document_utils.py` (synchronous)
   - Tool wrappers: `word_document_server/tools/content_tools.py` (async)
   - MCP registration: `word_document_server/main.py` inside `register_tools()`

5. **Hard rules:**
   - Use `uv run` for all Python execution
   - TDD: write tests first, verify they fail, implement, verify they pass
   - One branch per fix, named as specified
   - One commit per fix with specified message
   - PR against upstream `https://github.com/GongRzhe/Office-Word-MCP-Server`

## Agent prompt template

```
You are working on the Office-Word-MCP-Server project.

REPOSITORY: C:\Users\brandon\.claude\mcp\Office-Word-MCP-Server
FORK ORIGIN: https://github.com/frastlin/Office-Word-MCP-Server.git
UPSTREAM: https://github.com/GongRzhe/Office-Word-MCP-Server

TASK: [Plan title]
BRANCH: [branch name]

This is a TDD task. Follow this exact sequence:
1. git checkout -b [branch-name]
2. If tests/conftest.py does not exist, create it (spec below)
3. Add pytest config to pyproject.toml if missing
4. Create the test file with tests from the plan
5. Run tests to verify they FAIL: uv run pytest [test-file] -v
6. Implement the fix/feature per the plan
7. Run tests to verify they PASS: uv run pytest [test-file] -v
8. Run full test suite: uv run pytest tests/ -v
9. Stage and commit: git commit -m "[commit message]"
10. Push: git push -u origin [branch-name]
11. Create PR: gh pr create --repo GongRzhe/Office-Word-MCP-Server ...

--- PLAN ---
[Paste the full plan file content here]

--- CONFTEST.PY SPEC ---
[Paste 00-shared-test-infrastructure.md content here]
```

## Open questions / risks

1. **Merge conflicts in `document_utils.py`**: Plans 1-6 all modify this file. The integration agent must resolve import statements and function ordering. Mitigation: each change is additive (new helpers, modified functions) with minimal overlap in line ranges.

2. **Plan 2 → Plan 3 dependency**: Plan 3 reuses `_normalize_text` from Plan 2. If dispatching in parallel, Plan 3's agent should define `_normalize_text` independently, and the integration agent deduplicates. Recommended: sequence Plan 3 after Plan 2.

3. **`uv` availability**: All agents assume `uv` is installed. If not, fall back to `python -m pytest`.

4. **Upstream PR acceptance**: The upstream repo (GongRzhe) may have different coding standards or require changes. Each PR is self-contained to maximize acceptance.

5. **`conftest.py` duplication**: Each parallel agent creates its own `conftest.py`. The integration agent must merge into one unified file. Mitigation: all agents use the same fixture spec from Plan 0.

6. **Upstream may have been updated**: Before branching, agents should `git fetch` and ensure `main` is current.
