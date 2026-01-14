"""Tests for platform-agnostic presentation classes."""

import pytest

from gslides_api.agnostic.element import (
    ContentType,
    MarkdownTextElement,
    MarkdownImageElement,
    MarkdownTableElement,
    MarkdownChartElement,
)
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


class TestMarkdownSlideEmptyElements:
    """Tests for parsing and handling empty elements in MarkdownSlide."""

    def test_from_markdown_consecutive_comments(self):
        """Two comments in a row should create first element with content=None."""
        markdown = """<!-- slide: Test -->
<!-- text: First -->
<!-- text: Second -->
Some content for second"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert slide.name == "Test"
        assert len(slide.elements) == 2
        assert slide.elements[0].name == "First"
        assert slide.elements[0].content is None
        assert slide.elements[1].name == "Second"
        assert slide.elements[1].content == "Some content for second"

    def test_from_markdown_trailing_comment(self):
        """Comment at end of string should create element with content=None."""
        markdown = """<!-- slide: Test -->
<!-- text: First -->
Some content
<!-- text: Second -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert slide.name == "Test"
        assert len(slide.elements) == 2
        assert slide.elements[0].name == "First"
        assert slide.elements[0].content == "Some content"
        assert slide.elements[1].name == "Second"
        assert slide.elements[1].content is None

    def test_from_markdown_all_empty_elements(self):
        """All elements with no content should have content=None."""
        markdown = """<!-- slide: Test -->
<!-- text: A -->
<!-- text: B -->
<!-- text: C -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 3
        for el in slide.elements:
            assert el.content is None

    def test_from_markdown_empty_text_element(self):
        """Empty text element should have content=None and content_type=TEXT."""
        markdown = """<!-- text: Title -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "Title"
        assert slide.elements[0].content is None
        assert slide.elements[0].content_type == ContentType.TEXT

    def test_from_markdown_empty_image_element(self):
        """Empty image element should have content=None."""
        markdown = """<!-- image: MyImage -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "MyImage"
        assert slide.elements[0].content is None
        assert slide.elements[0].content_type == ContentType.IMAGE

    def test_from_markdown_empty_table_element(self):
        """Empty table element should have content=None."""
        markdown = """<!-- table: MyTable -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "MyTable"
        assert slide.elements[0].content is None
        assert slide.elements[0].content_type == ContentType.TABLE

    def test_from_markdown_empty_chart_element(self):
        """Empty chart element should have content=None."""
        markdown = """<!-- chart: MyChart -->"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "MyChart"
        assert slide.elements[0].content is None
        assert slide.elements[0].content_type == ContentType.CHART

    def test_to_markdown_empty_text_element(self):
        """Empty text element should output just the comment."""
        element = MarkdownTextElement(name="Title", content=None)
        assert element.to_markdown() == "<!-- text: Title -->"

    def test_to_markdown_empty_image_element(self):
        """Empty image element should output just the comment."""
        element = MarkdownImageElement(name="MyImage", content=None)
        assert element.to_markdown() == "<!-- image: MyImage -->"

    def test_to_markdown_empty_table_element(self):
        """Empty table element should output just the comment."""
        element = MarkdownTableElement(name="MyTable", content=None)
        assert element.to_markdown() == "<!-- table: MyTable -->"

    def test_to_markdown_empty_chart_element(self):
        """Empty chart element should output just the comment."""
        element = MarkdownChartElement(name="MyChart", content=None)
        assert element.to_markdown() == "<!-- chart: MyChart -->"

    def test_roundtrip_empty_elements(self):
        """Empty elements should survive to_markdown -> from_markdown roundtrip."""
        # Create slide with empty elements
        slide = MarkdownSlide(
            name="Test",
            elements=[
                MarkdownTextElement(name="Text1", content=None),
                MarkdownTextElement(name="Text2", content="Has content"),
                MarkdownChartElement(name="Chart", content=None),
            ],
        )
        # Convert to markdown and back
        markdown = slide.to_markdown()
        parsed = MarkdownSlide.from_markdown(markdown)

        assert parsed.name == "Test"
        assert len(parsed.elements) == 3
        assert parsed.elements[0].name == "Text1"
        assert parsed.elements[0].content is None
        assert parsed.elements[1].name == "Text2"
        assert parsed.elements[1].content == "Has content"
        assert parsed.elements[2].name == "Chart"
        assert parsed.elements[2].content is None

    def test_mixed_empty_and_filled_elements(self):
        """Parsing should correctly handle mix of empty and filled elements."""
        markdown = """<!-- slide: Mixed -->
<!-- text: Title -->
My Title

<!-- image: EmptyImage -->
<!-- text: Body -->
Some body text

<!-- table: EmptyTable -->
<!-- chart: Chart -->
Chart description"""
        slide = MarkdownSlide.from_markdown(markdown)
        assert len(slide.elements) == 5

        # Title has content
        assert slide.elements[0].name == "Title"
        assert slide.elements[0].content == "My Title"

        # EmptyImage is empty
        assert slide.elements[1].name == "EmptyImage"
        assert slide.elements[1].content is None

        # Body has content
        assert slide.elements[2].name == "Body"
        assert slide.elements[2].content == "Some body text"

        # EmptyTable is empty
        assert slide.elements[3].name == "EmptyTable"
        assert slide.elements[3].content is None

        # Chart has content
        assert slide.elements[4].name == "Chart"
        assert slide.elements[4].content == "Chart description"
