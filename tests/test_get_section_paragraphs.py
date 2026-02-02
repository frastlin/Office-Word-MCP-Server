"""Tests for get_section_paragraphs utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import get_section_paragraphs


class TestGetSectionParagraphs:
    """Core section extraction behavior."""

    def test_basic_h1_section(self, make_docx):
        """H1 section returns content until next H1."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Section A"}]},
            "Content 1",
            "Content 2",
            {"style": "Heading 1", "runs": [{"text": "Section B"}]},
            "Content 3",
        ])
        result = get_section_paragraphs(path, "Section A")
        assert "error" not in result
        assert result["heading_index"] == 0
        assert result["heading_text"] == "Section A"
        assert result["heading_style"] == "Heading 1"
        assert result["heading_level"] == 1
        assert result["next_heading_index"] == 3
        # Content paragraphs (excluding heading itself)
        content = [p for p in result["paragraphs"] if p["index"] != 0]
        content_texts = [p["text"] for p in content]
        assert "Content 1" in content_texts
        assert "Content 2" in content_texts
        assert "Content 3" not in content_texts

    def test_h2_section_stops_at_same_level(self, make_docx):
        """H2 section stops at next H2."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter"}]},
            {"style": "Heading 2", "runs": [{"text": "Part A"}]},
            "Content A",
            {"style": "Heading 2", "runs": [{"text": "Part B"}]},
            "Content B",
        ])
        result = get_section_paragraphs(path, "Part A")
        assert result["heading_level"] == 2
        assert result["next_heading_index"] == 3
        content = [p for p in result["paragraphs"] if p["index"] > result["heading_index"]]
        assert len(content) == 1
        assert content[0]["text"] == "Content A"

    def test_h2_section_stops_at_higher_level(self, make_docx):
        """H2 section stops at H1 (higher level = lower number)."""
        path = make_docx(paragraphs=[
            {"style": "Heading 2", "runs": [{"text": "Subsection"}]},
            "Content",
            {"style": "Heading 1", "runs": [{"text": "Next Chapter"}]},
        ])
        result = get_section_paragraphs(path, "Subsection")
        assert result["next_heading_index"] == 2

    def test_last_section_no_next_heading(self, make_docx):
        """Last section returns content to end, next_heading_index is null."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Only Section"}]},
            "Content 1",
            "Content 2",
        ])
        result = get_section_paragraphs(path, "Only Section")
        assert result["next_heading_index"] is None
        content = [p for p in result["paragraphs"] if p["index"] > 0]
        assert len(content) == 2

    def test_empty_section(self, make_docx):
        """Section with no content between headings returns only heading in paragraphs."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Empty"}]},
            {"style": "Heading 1", "runs": [{"text": "Next"}]},
        ])
        result = get_section_paragraphs(path, "Empty")
        assert result["next_heading_index"] == 1
        content = [p for p in result["paragraphs"] if p["index"] > 0]
        assert len(content) == 0

    def test_include_heading_true(self, make_docx):
        """Default includes heading paragraph in results."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Title")
        assert any(p["text"] == "Title" for p in result["paragraphs"])

    def test_include_heading_false(self, make_docx):
        """include_heading=False omits heading from paragraphs list."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Title"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Title", include_heading=False)
        assert not any(p["text"] == "Title" for p in result["paragraphs"])

    def test_with_empty_spacer_paragraphs(self, make_docx):
        """Empty spacer paragraphs between content are included."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Section"}]},
            "",
            "Content 1",
            "",
            "Content 2",
            "",
            {"style": "Heading 1", "runs": [{"text": "Next"}]},
        ])
        result = get_section_paragraphs(path, "Section", include_heading=False)
        assert len(result["paragraphs"]) == 5  # 3 spacers + 2 content


class TestSectionHeadingMatching:
    """Heading text matching behavior."""

    def test_heading_not_found(self, make_docx):
        """Non-existent heading returns error."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Actual Title"}]},
        ])
        result = get_section_paragraphs(path, "Nonexistent")
        assert "error" in result

    def test_partial_text_match(self, make_docx):
        """Substring of heading text finds the heading."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "Chapter 1: Introduction"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "Introduction")
        assert result["heading_index"] == 0

    def test_normalized_whitespace_match(self, make_docx):
        """Extra whitespace in search text still matches."""
        path = make_docx(paragraphs=[
            {"style": "Heading 1", "runs": [{"text": "My  Section"}]},
            "Content",
        ])
        result = get_section_paragraphs(path, "My Section")
        assert result["heading_index"] == 0

    def test_file_not_found(self, tmp_path):
        """Non-existent file returns error dict."""
        result = get_section_paragraphs(str(tmp_path / "missing.docx"), "Heading")
        assert "error" in result
