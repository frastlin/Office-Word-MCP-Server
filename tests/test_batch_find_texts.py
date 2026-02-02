"""Tests for batch find_texts utility."""
import pytest
from docx import Document
from word_document_server.utils.extended_document_utils import find_texts


class TestBatchFindTexts:
    def test_find_multiple_strings(self, make_docx):
        path = make_docx(paragraphs=["Alpha content here", "Beta content here", "Gamma content here"])
        result = find_texts(path, ["Alpha", "Beta", "Gamma"])
        assert "Alpha" in result
        assert "Beta" in result
        assert "Gamma" in result
        assert result["Alpha"]["occurrences"][0]["paragraph_index"] == 0
        assert result["Beta"]["occurrences"][0]["paragraph_index"] == 1
        assert result["Gamma"]["occurrences"][0]["paragraph_index"] == 2

    def test_missing_string_returns_empty(self, make_docx):
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Nonexistent"])
        assert result["Nonexistent"]["total_count"] == 0
        assert result["Nonexistent"]["occurrences"] == []

    def test_single_string_works(self, make_docx):
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Hello"])
        assert result["Hello"]["total_count"] == 1

    def test_case_insensitive(self, make_docx):
        path = make_docx(paragraphs=["Hello World"])
        result = find_texts(path, ["hello", "WORLD"], match_case=False)
        assert result["hello"]["total_count"] == 1
        assert result["WORLD"]["total_count"] == 1

    def test_include_paragraph_text_propagates(self, make_docx):
        path = make_docx(paragraphs=["Alpha content", "Beta content"])
        result = find_texts(path, ["Alpha", "Beta"], include_paragraph_text=True)
        assert "text" in result["Alpha"]["occurrences"][0]
        assert result["Alpha"]["occurrences"][0]["text"] == "Alpha content"

    def test_empty_list_returns_empty_dict(self, make_docx):
        path = make_docx(paragraphs=["Content"])
        result = find_texts(path, [])
        assert result == {} or result.get("results") == {}

    def test_duplicate_search_terms(self, make_docx):
        path = make_docx(paragraphs=["Hello world"])
        result = find_texts(path, ["Hello", "Hello"])
        assert "Hello" in result
        assert result["Hello"]["total_count"] == 1

    def test_each_result_has_standard_fields(self, make_docx):
        path = make_docx(paragraphs=["Target text"])
        result = find_texts(path, ["Target"])
        entry = result["Target"]
        assert "occurrences" in entry
        assert "total_count" in entry

    def test_file_not_found(self, tmp_path):
        result = find_texts(str(tmp_path / "missing.docx"), ["anything"])
        assert "error" in result

    def test_multiple_occurrences_same_string(self, make_docx):
        path = make_docx(paragraphs=["Match here", "No match", "Match again"])
        result = find_texts(path, ["Match"])
        assert result["Match"]["total_count"] == 2
        indices = [o["paragraph_index"] for o in result["Match"]["occurrences"]]
        assert 0 in indices
        assert 2 in indices
