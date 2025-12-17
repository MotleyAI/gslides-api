"""Tests for platform-agnostic presentation classes."""

import pytest

from gslides_api.agnostic.element import MarkdownTextElement
from gslides_api.agnostic.presentation import MarkdownSlide


class TestMarkdownSlideUniqueElementNames:
    """Tests for MarkdownSlide element name uniqueness validation."""

    def test_unique_element_names_allowed(self):
        """Should allow slides with unique element names."""
        slide = MarkdownSlide(
            name="Test",
            elements=[
                MarkdownTextElement.placeholder(name="Title"),
                MarkdownTextElement.placeholder(name="Subtitle"),
                MarkdownTextElement.placeholder(name="Body"),
            ],
        )
        assert len(slide.elements) == 3

    def test_duplicate_element_names_raises_on_construction(self):
        """Should raise ValueError when creating slide with duplicate element names."""
        with pytest.raises(ValueError, match="Duplicate element names found"):
            MarkdownSlide(
                name="Test",
                elements=[
                    MarkdownTextElement.placeholder(name="Title"),
                    MarkdownTextElement.placeholder(name="Title"),
                ],
            )

    def test_duplicate_element_names_raises_on_assignment(self):
        """Should raise ValueError when assigning elements with duplicate names."""
        slide = MarkdownSlide(
            name="Test",
            elements=[
                MarkdownTextElement.placeholder(name="Title"),
            ],
        )
        with pytest.raises(ValueError, match="Duplicate element names found"):
            slide.elements = [
                MarkdownTextElement.placeholder(name="Body"),
                MarkdownTextElement.placeholder(name="Body"),
            ]

    def test_error_message_contains_duplicate_names(self):
        """Error message should list the duplicate names."""
        with pytest.raises(ValueError, match="Title"):
            MarkdownSlide(
                name="Test",
                elements=[
                    MarkdownTextElement.placeholder(name="Title"),
                    MarkdownTextElement.placeholder(name="Title"),
                ],
            )

    def test_multiple_duplicates_reported(self):
        """All duplicate names should be reported."""
        with pytest.raises(ValueError) as exc_info:
            MarkdownSlide(
                name="Test",
                elements=[
                    MarkdownTextElement.placeholder(name="A"),
                    MarkdownTextElement.placeholder(name="A"),
                    MarkdownTextElement.placeholder(name="B"),
                    MarkdownTextElement.placeholder(name="B"),
                ],
            )
        error_msg = str(exc_info.value)
        assert "A" in error_msg
        assert "B" in error_msg

    def test_empty_elements_list_allowed(self):
        """Should allow slides with no elements."""
        slide = MarkdownSlide(name="Empty", elements=[])
        assert len(slide.elements) == 0

    def test_single_element_allowed(self):
        """Should allow slides with a single element."""
        slide = MarkdownSlide(
            name="Single",
            elements=[MarkdownTextElement.placeholder(name="Title")],
        )
        assert len(slide.elements) == 1
