"""Tests for get_all_comments round-tripping with add_comment.

Reproduces Bug 3: after adding N comments via add_comment, get_all_comments
should return N — not 0.
"""
import asyncio
import json
import zipfile

import pytest

from word_document_server.tools.comment_management_tools import add_comment
from word_document_server.tools.comment_tools import get_all_comments


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


class TestGetAllCommentsRoundTrip:
    def test_returns_comments_added_via_add_comment(self, make_docx):
        path = make_docx(paragraphs=[
            "First paragraph text",
            "Second paragraph text",
            "Third paragraph text",
        ])
        for anchor, body in [
            ("First", "alpha"),
            ("Second", "beta"),
            ("Third", "gamma"),
        ]:
            r = _parse(_run(add_comment(path, anchor, body, author="Tester")))
            assert r["success"], r

        data = _parse(_run(get_all_comments(path)))
        assert data["success"] is True
        assert data["total_comments"] == 3, data
        texts = sorted(c["text"] for c in data["comments"])
        assert texts == ["alpha", "beta", "gamma"]
        for c in data["comments"]:
            assert c["author"] == "Tester"

    def test_returns_empty_when_no_comments(self, make_docx):
        path = make_docx(paragraphs=["Untouched paragraph"])
        data = _parse(_run(get_all_comments(path)))
        assert data["success"] is True
        assert data["total_comments"] == 0
        assert data["comments"] == []

    def test_still_reads_comments_when_extra_zip_entry_present(self, make_docx):
        """Regression safeguard: even if the .docx has an undeclared extra
        zip part (e.g. the historical Bug 1 meta.json), comment extraction
        must still work because it should read word/comments.xml directly."""
        path = make_docx(paragraphs=["Some text here"])
        _run(add_comment(path, "Some", "note", author="Tester"))
        # Inject an extra zip entry that is NOT declared in [Content_Types].xml
        with zipfile.ZipFile(path, "a") as z:
            z.writestr("unused-side-data.json", '{"x":1}')

        data = _parse(_run(get_all_comments(path)))
        assert data["success"] is True, data
        assert data["total_comments"] == 1
        assert data["comments"][0]["text"] == "note"
