"""Tests for get_paragraph_range utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import get_paragraph_range


class TestGetParagraphRange:
    """Core read behavior."""

    def test_read_full_range(self, make_docx):
        """Read paragraphs 1-3 from 5-paragraph doc, verify 3 results."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 1, 3)
        assert "error" not in result
        assert len(result["paragraphs"]) == 3
        assert result["paragraphs"][0]["text"] == "B"
        assert result["paragraphs"][1]["text"] == "C"
        assert result["paragraphs"][2]["text"] == "D"

    def test_single_paragraph_range(self, make_docx):
        """start == end returns exactly one paragraph."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 1, 1)
        assert len(result["paragraphs"]) == 1
        assert result["paragraphs"][0]["text"] == "B"

    def test_includes_correct_fields(self, make_docx):
        """Each result has index, text, style, is_heading fields."""
        path = make_docx(paragraphs=["Hello world"])
        result = get_paragraph_range(path, 0, 0)
        para = result["paragraphs"][0]
        assert "index" in para
        assert "text" in para
        assert "style" in para
        assert "is_heading" in para

    def test_heading_detected(self, make_docx):
        """Heading paragraphs have is_heading=True and correct style."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
            {"style": "Heading 2", "runs": [{"text": "Subtitle"}]},
        ])
        result = get_paragraph_range(path, 0, 2)
        assert result["paragraphs"][0]["is_heading"] is True
        assert result["paragraphs"][0]["style"] == "Heading 1"
        assert result["paragraphs"][1]["is_heading"] is False
        assert result["paragraphs"][2]["is_heading"] is True
        assert result["paragraphs"][2]["style"] == "Heading 2"

    def test_empty_paragraphs_included(self, make_docx):
        """Empty paragraphs are returned with empty text."""
        path = make_docx(paragraphs=["Content", "", "More content"])
        result = get_paragraph_range(path, 0, 2)
        assert result["paragraphs"][1]["text"] == ""

    def test_indices_are_correct(self, make_docx):
        """Returned index values match the actual document indices."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 2, 4)
        assert result["paragraphs"][0]["index"] == 2
        assert result["paragraphs"][1]["index"] == 3
        assert result["paragraphs"][2]["index"] == 4

    def test_count_field(self, make_docx):
        """Result includes count field matching paragraphs length."""
        path = make_docx(paragraphs=["A", "B", "C", "D", "E"])
        result = get_paragraph_range(path, 1, 3)
        assert result["count"] == 3


class TestGetParagraphRangeValidation:
    """Input validation."""

    def test_start_greater_than_end(self, make_docx):
        """start > end returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 2, 1)
        assert "error" in result

    def test_end_out_of_bounds(self, make_docx):
        """end beyond doc length returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, 0, 10)
        assert "error" in result

    def test_negative_start(self, make_docx):
        """Negative start returns error."""
        path = make_docx(paragraphs=["A", "B", "C"])
        result = get_paragraph_range(path, -1, 2)
        assert "error" in result

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error dict."""
        result = get_paragraph_range(str(tmp_path / "missing.docx"), 0, 1)
        assert "error" in result
