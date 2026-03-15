"""Tests for comment management tools."""
import asyncio
import json
import pytest

from word_document_server.tools.comment_management_tools import (
    add_comment,
    reply_to_comment,
    resolve_comment,
    delete_comment,
)


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


# ── Add Comment ──────────────────────────────────────────────────────────

class TestAddComment:
    def test_add_comment_basic(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(add_comment(path, "Hello", "Please review", author="Tester")))
        assert data["success"] is True
        assert data["comment_id"] == 0
        assert data["anchor_text"] == "Hello"
        assert data["author"] == "Tester"

    def test_add_comment_with_author(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(add_comment(path, "Hello", "Review this", author="Alice")))
        assert data["success"] is True
        assert data["author"] == "Alice"

    def test_add_comment_empty_text_returns_error(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(add_comment(path, "Hello", "", author="Tester")))
        assert data["success"] is False
        assert "empty" in data["error"]

    def test_add_comment_anchor_not_found(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(add_comment(path, "Nonexistent", "Comment", author="Tester")))
        assert data["success"] is False
        assert "not found" in data["error"]


# ── Reply To Comment ────────────────────────────────────────────────────

class TestReplyToComment:
    def test_reply_to_comment(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        # First add a comment
        add_data = _parse(_run(add_comment(path, "Hello", "Please review", author="Alice")))
        comment_id = add_data["comment_id"]

        # Then reply to it
        data = _parse(_run(reply_to_comment(path, comment_id, "Looks good", author="Bob")))
        assert data["success"] is True
        assert data["parent_comment_id"] == comment_id
        assert data["author"] == "Bob"

    def test_reply_to_nonexistent_comment(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(reply_to_comment(path, 999, "Reply", author="Tester")))
        assert data["success"] is False


# ── Resolve Comment ──────────────────────────────────────────────────────

class TestResolveComment:
    def test_resolve_comment(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        add_data = _parse(_run(add_comment(path, "Hello", "Please review", author="Alice")))
        comment_id = add_data["comment_id"]

        data = _parse(_run(resolve_comment(path, comment_id)))
        assert data["success"] is True
        assert data["resolved"] == comment_id


# ── Delete Comment ───────────────────────────────────────────────────────

class TestDeleteComment:
    def test_delete_comment(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        add_data = _parse(_run(add_comment(path, "Hello", "Please review", author="Alice")))
        comment_id = add_data["comment_id"]

        data = _parse(_run(delete_comment(path, comment_id)))
        assert data["success"] is True
        assert data["deleted"] == comment_id

    def test_delete_nonexistent_comment(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(delete_comment(path, 999)))
        assert data["success"] is False
