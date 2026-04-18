"""Integrity tests for .docx zip output produced by mutating tools.

These tests verify that mutating MCP tools (add_comment, delete_comment, etc.)
produce valid OOXML packages — i.e. no stray parts at the zip root and no
orphan XML entries that would cause Word to flag "unreadable content".
"""
import asyncio
import json
import re
import zipfile

import pytest

from word_document_server.tools.comment_management_tools import (
    add_comment,
    delete_comment,
)


def _run(coro):
    return asyncio.run(coro)


def _parse(result):
    return json.loads(result)


# ── Bug 1: meta.json must not appear at the zip root ─────────────────────

class TestMetaJsonStripped:
    def test_add_comment_does_not_leave_meta_json(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        data = _parse(_run(add_comment(path, "Hello", "please review", author="Tester")))
        assert data["success"] is True, data
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
        assert "meta.json" not in names, (
            f"meta.json leaked into docx zip root; full name list: {names}"
        )

    def test_delete_comment_does_not_leave_meta_json(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        add_data = _parse(_run(add_comment(path, "Hello", "x", author="Tester")))
        cid = add_data["comment_id"]
        _parse(_run(delete_comment(path, cid)))
        with zipfile.ZipFile(path) as z:
            assert "meta.json" not in z.namelist()
