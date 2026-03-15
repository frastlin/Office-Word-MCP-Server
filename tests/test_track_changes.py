"""Tests for track changes tools."""
import asyncio
import json
import pytest

from word_document_server.tools.track_changes_tools import (
    replace_with_track_changes,
    delete_with_track_changes,
    insert_after_with_track_changes,
    insert_before_with_track_changes,
    list_revisions,
    accept_revision,
    reject_revision,
    accept_all_revisions,
    reject_all_revisions,
    get_visible_text,
    count_tracked_matches,
)


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


# ── Replace ──────────────────────────────────────────────────────────────

class TestReplaceWithTrackChanges:
    def test_replace_single_occurrence(self, make_docx):
        path = make_docx(paragraphs=["Hello World", "Hello World"])
        data = _parse(_run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester", occurrence=0)))
        assert data["success"] is True
        assert data["replacements"] == 1

        vis = _parse(_run(get_visible_text(path)))
        assert vis["text"].count("Goodbye") == 1
        assert vis["text"].count("Hello") == 1  # second occurrence untouched

    def test_replace_all_occurrences(self, make_docx):
        path = make_docx(paragraphs=["Hello World", "Hello World"])
        data = _parse(_run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester")))
        assert data["success"] is True
        assert data["replacements"] == 2

        vis = _parse(_run(get_visible_text(path)))
        assert vis["text"].count("Goodbye") == 2
        assert "Hello" not in vis["text"]

    def test_replace_no_match_returns_message(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(replace_with_track_changes(path, "Nonexistent", "X", author="Tester")))
        assert data["success"] is True
        assert data["replacements"] == 0
        assert "No matches" in data["message"]

    def test_replace_empty_find_text_returns_error(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(replace_with_track_changes(path, "", "X")))
        assert data["success"] is False
        assert "empty" in data["error"]

    def test_replace_nonexistent_file_returns_error(self):
        data = _parse(_run(replace_with_track_changes("/nonexistent/file.docx", "a", "b")))
        assert data["success"] is False
        assert "does not exist" in data["error"]

    def test_replace_with_author(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(replace_with_track_changes(path, "Hello", "Goodbye", author="Alice")))
        assert data["success"] is True
        assert data["author"] == "Alice"

        revs = _parse(_run(list_revisions(path)))
        assert revs["success"] is True
        for r in revs["revisions"]:
            assert r["author"] == "Alice"


# ── Delete ───────────────────────────────────────────────────────────────

class TestDeleteWithTrackChanges:
    def test_delete_single_occurrence(self, make_docx):
        path = make_docx(paragraphs=["Hello World", "Hello World"])
        data = _parse(_run(delete_with_track_changes(path, "Hello", author="Tester", occurrence=0)))
        assert data["success"] is True
        assert data["deletions"] == 1

        vis = _parse(_run(get_visible_text(path)))
        assert vis["text"].count("Hello") == 1  # one still visible

    def test_delete_all_occurrences(self, make_docx):
        path = make_docx(paragraphs=["Hello World", "Hello World"])
        data = _parse(_run(delete_with_track_changes(path, "Hello", author="Tester")))
        assert data["success"] is True
        assert data["deletions"] == 2

        vis = _parse(_run(get_visible_text(path)))
        assert "Hello" not in vis["text"]

    def test_delete_no_match_returns_message(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(delete_with_track_changes(path, "Nonexistent", author="Tester")))
        assert data["success"] is True
        assert data["deletions"] == 0


# ── Insert After ─────────────────────────────────────────────────────────

class TestInsertAfterWithTrackChanges:
    def test_insert_after_basic(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(insert_after_with_track_changes(path, "World", " Extra", author="Tester")))
        assert data["success"] is True

        vis = _parse(_run(get_visible_text(path)))
        assert "World Extra" in vis["text"]

    def test_insert_after_specific_occurrence(self, make_docx):
        path = make_docx(paragraphs=["Hello World", "Another World"])
        data = _parse(_run(insert_after_with_track_changes(path, "World", "[1]", author="Tester", occurrence=1)))
        assert data["success"] is True

        vis = _parse(_run(get_visible_text(path)))
        # Second "World" (in paragraph 2) should have the insertion
        assert "World[1]" in vis["text"]

    def test_insert_after_empty_anchor_returns_error(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(insert_after_with_track_changes(path, "", "X", author="Tester")))
        assert data["success"] is False
        assert "empty" in data["error"]


# ── Insert Before ────────────────────────────────────────────────────────

class TestInsertBeforeWithTrackChanges:
    def test_insert_before_basic(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(insert_before_with_track_changes(path, "Hello", "NEW ", author="Tester")))
        assert data["success"] is True

        vis = _parse(_run(get_visible_text(path)))
        assert "NEW Hello" in vis["text"]


# ── List Revisions ───────────────────────────────────────────────────────

class TestListRevisions:
    def test_list_no_revisions(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(list_revisions(path)))
        assert data["success"] is True
        assert data["total"] == 0
        assert data["revisions"] == []

    def test_list_after_replace(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Alice"))

        data = _parse(_run(list_revisions(path)))
        assert data["success"] is True
        assert data["total"] >= 2  # at least a deletion and an insertion
        types = {r["type"] for r in data["revisions"]}
        assert "deletion" in types
        assert "insertion" in types
        for r in data["revisions"]:
            assert r["author"] == "Alice"
            assert r["id"] is not None

    def test_list_filter_by_author(self, make_docx):
        path = make_docx(paragraphs=["Hello World again"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Alice"))
        _run(insert_after_with_track_changes(path, "again", " END", author="Bob"))

        alice_data = _parse(_run(list_revisions(path, author="Alice")))
        bob_data = _parse(_run(list_revisions(path, author="Bob")))

        assert alice_data["success"] is True
        assert bob_data["success"] is True
        for r in alice_data["revisions"]:
            assert r["author"] == "Alice"
        for r in bob_data["revisions"]:
            assert r["author"] == "Bob"


# ── Accept / Reject Revision ────────────────────────────────────────────

class TestAcceptRejectRevision:
    def test_accept_revision(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(insert_after_with_track_changes(path, "World", " Extra", author="Tester"))

        revs = _parse(_run(list_revisions(path)))
        rev_id = revs["revisions"][0]["id"]

        data = _parse(_run(accept_revision(path, rev_id)))
        assert data["success"] is True

        # After accepting, that revision should be gone
        revs2 = _parse(_run(list_revisions(path)))
        remaining_ids = {r["id"] for r in revs2["revisions"]}
        assert rev_id not in remaining_ids

        # Text should still include the insertion (now permanent)
        vis = _parse(_run(get_visible_text(path)))
        assert "Extra" in vis["text"]

    def test_reject_revision(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(insert_after_with_track_changes(path, "World", " Extra", author="Tester"))

        revs = _parse(_run(list_revisions(path)))
        rev_id = revs["revisions"][0]["id"]

        data = _parse(_run(reject_revision(path, rev_id)))
        assert data["success"] is True

        # After rejecting insertion, the text should be gone
        vis = _parse(_run(get_visible_text(path)))
        assert "Extra" not in vis["text"]


# ── Accept / Reject All ─────────────────────────────────────────────────

class TestAcceptRejectAll:
    def test_accept_all(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester"))

        data = _parse(_run(accept_all_revisions(path)))
        assert data["success"] is True
        assert data["accepted"] >= 2

        revs = _parse(_run(list_revisions(path)))
        assert revs["total"] == 0

        vis = _parse(_run(get_visible_text(path)))
        assert "Goodbye" in vis["text"]
        assert "Hello" not in vis["text"]

    def test_reject_all(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester"))

        data = _parse(_run(reject_all_revisions(path)))
        assert data["success"] is True
        assert data["rejected"] >= 2

        revs = _parse(_run(list_revisions(path)))
        assert revs["total"] == 0

        vis = _parse(_run(get_visible_text(path)))
        assert "Hello" in vis["text"]
        assert "Goodbye" not in vis["text"]

    def test_accept_all_by_author(self, make_docx):
        path = make_docx(paragraphs=["Hello World again"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Alice"))
        _run(insert_after_with_track_changes(path, "again", " END", author="Bob"))

        data = _parse(_run(accept_all_revisions(path, author="Alice")))
        assert data["success"] is True

        # Bob's revision should still be pending
        revs = _parse(_run(list_revisions(path, author="Bob")))
        assert revs["total"] >= 1


# ── Visible Text ─────────────────────────────────────────────────────────

class TestGetVisibleText:
    def test_visible_text_shows_insertions_hides_deletions(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester"))

        vis = _parse(_run(get_visible_text(path)))
        assert vis["success"] is True
        assert "Goodbye" in vis["text"]
        assert "Hello" not in vis["text"]


# ── Count Tracked Matches ────────────────────────────────────────────────

class TestCountTrackedMatches:
    def test_count_in_visible_text(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        _run(replace_with_track_changes(path, "Hello", "Goodbye", author="Tester"))

        data = _parse(_run(count_tracked_matches(path, "Goodbye")))
        assert data["success"] is True
        assert data["count"] == 1

        data2 = _parse(_run(count_tracked_matches(path, "Hello")))
        assert data2["count"] == 0  # deleted text not counted

    def test_count_empty_text_returns_error(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(count_tracked_matches(path, "")))
        assert data["success"] is False
        assert "empty" in data["error"]
