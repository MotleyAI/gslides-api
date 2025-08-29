"""
Tests for the markdown_to_text_elements() function in gslides_api.markdown.from_markdown.

This file tests the core functionality of converting markdown text to TextElement objects.
"""

import pytest

from gslides_api import InsertTextRequest
from gslides_api.markdown.from_markdown import markdown_to_text_elements
from gslides_api.text import TextElement, TextRun, TextStyle


class TestMarkdownToTextElements:
    """Test the markdown_to_text_elements() function."""

    def test_simple_string_returns_single_text_element(self):
        """Test that markdown_to_text_elements() on a simple string returns a list of length 1 with correct content."""
        # Test with a simple string
        result = markdown_to_text_elements("Test")

        # Should return a list of length 1
        assert len(result) == 2, f"Expected list of length 2, got {len(result)}"

        # The single element should be a TextElement
        element = result[0]
        assert isinstance(
            element, InsertTextRequest
        ), f"InsertTextRequest, got {type(element)}"

        assert element.text == "Test", f"Expected 'Test', got '{element.text}'"

        # Verify the TextElement has proper indices
        assert (
            element.insertionIndex == 0
        ), f"Expected startIndex 0, got {element.startIndex}"
