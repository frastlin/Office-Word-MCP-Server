"""Tests for get_document_info with include_outline enhancement."""
import pytest
from docx import Document
from word_document_server.utils.document_utils import get_document_properties


class TestGetDocumentInfoDefault:
    """Backward-compatible default behavior."""

    def test_default_no_headings_key(self, make_docx):
        """Default response does not include headings key."""
        path = make_docx(paragraphs=["Content"])
        result = get_document_properties(path)
        assert "headings" not in result

    def test_default_still_has_word_count(self, make_docx):
        """Default response still includes standard fields."""
        path = make_docx(paragraphs=["Hello world"])
        result = get_document_properties(path)
        assert "word_count" in result
        assert "paragraph_count" in result


class TestGetDocumentInfoWithOutline:
    """include_outline=True behavior."""

    def test_include_outline_adds_headings(self, make_docx):
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter 1"}]},
            "Content",
            {"style": "Heading 2", "runs": [{"text": "Section 1.1"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        assert "headings" in result
        assert len(result["headings"]) == 2

    def test_headings_have_correct_fields(self, make_docx):
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        h = result["headings"][0]
        assert "index" in h
        assert "text" in h
        assert "style" in h
        assert "level" in h

    def test_heading_levels_correct(self, make_docx):
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "H1"}]},
            {"style": "Heading 2", "runs": [{"text": "H2"}]},
            {"style": "Heading 3", "runs": [{"text": "H3"}]},
        ])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"][0]["level"] == 1
        assert result["headings"][1]["level"] == 2
        assert result["headings"][2]["level"] == 3

    def test_heading_index_matches_position(self, make_docx):
        path = make_docx(paragraphs=[
            "Content before",
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content after",
        ])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"][0]["index"] == 1
        assert result["headings"][0]["text"] == "Title"

    def test_no_headings_returns_empty_array(self, make_docx):
        path = make_docx(paragraphs=["Just normal text", "More text"])
        result = get_document_properties(path, include_outline=True)
        assert result["headings"] == []

    def test_standard_fields_still_present(self, make_docx):
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_document_properties(path, include_outline=True)
        assert "word_count" in result
        assert "paragraph_count" in result
        assert "table_count" in result

    def test_include_outline_false_same_as_default(self, make_docx):
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
        ])
        result = get_document_properties(path, include_outline=False)
        assert "headings" not in result
