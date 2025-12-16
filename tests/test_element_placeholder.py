"""Tests for placeholder() classmethods on MarkdownSlideElement subclasses."""

import pytest

from gslides_api.agnostic.element import (
    ContentType,
    MarkdownChartElement,
    MarkdownImageElement,
    MarkdownTableElement,
    MarkdownTextElement,
)
from gslides_api.agnostic.presentation import MarkdownSlide


class TestMarkdownTextElementPlaceholder:
    """Tests for MarkdownTextElement.placeholder()."""

    def test_creates_valid_instance(self):
        """Test that placeholder creates a valid MarkdownTextElement instance."""
        elem = MarkdownTextElement.placeholder("MyText")
        assert isinstance(elem, MarkdownTextElement)

    def test_name_is_set(self):
        """Test that the name parameter is correctly set."""
        elem = MarkdownTextElement.placeholder("TestName")
        assert elem.name == "TestName"

    def test_content_type_is_text(self):
        """Test that content_type is TEXT."""
        elem = MarkdownTextElement.placeholder("MyText")
        assert elem.content_type == ContentType.TEXT

    def test_has_placeholder_content(self):
        """Test that the element has appropriate placeholder content."""
        elem = MarkdownTextElement.placeholder("MyText")
        assert elem.content == "Placeholder text"

    def test_to_markdown_works(self):
        """Test that to_markdown() produces valid output."""
        elem = MarkdownTextElement.placeholder("MyText")
        md = elem.to_markdown()
        assert "<!-- text: MyText -->" in md
        assert "Placeholder text" in md


class TestMarkdownImageElementPlaceholder:
    """Tests for MarkdownImageElement.placeholder()."""

    def test_creates_valid_instance(self):
        """Test that placeholder creates a valid MarkdownImageElement instance."""
        elem = MarkdownImageElement.placeholder("MyImage")
        assert isinstance(elem, MarkdownImageElement)

    def test_name_is_set(self):
        """Test that the name parameter is correctly set."""
        elem = MarkdownImageElement.placeholder("TestImage")
        assert elem.name == "TestImage"

    def test_content_type_is_image(self):
        """Test that content_type is IMAGE."""
        elem = MarkdownImageElement.placeholder("MyImage")
        assert elem.content_type == ContentType.IMAGE

    def test_has_placeholder_url(self):
        """Test that the element has a placeholder URL."""
        elem = MarkdownImageElement.placeholder("MyImage")
        assert "placeholder.com" in elem.content

    def test_to_markdown_works(self):
        """Test that to_markdown() produces valid output."""
        elem = MarkdownImageElement.placeholder("MyImage")
        md = elem.to_markdown()
        assert "<!-- image: MyImage -->" in md
        assert "![" in md
        assert "placeholder.com" in md


class TestMarkdownTableElementPlaceholder:
    """Tests for MarkdownTableElement.placeholder()."""

    def test_creates_valid_instance(self):
        """Test that placeholder creates a valid MarkdownTableElement instance."""
        elem = MarkdownTableElement.placeholder("MyTable")
        assert isinstance(elem, MarkdownTableElement)

    def test_name_is_set(self):
        """Test that the name parameter is correctly set."""
        elem = MarkdownTableElement.placeholder("TestTable")
        assert elem.name == "TestTable"

    def test_content_type_is_table(self):
        """Test that content_type is TABLE."""
        elem = MarkdownTableElement.placeholder("MyTable")
        assert elem.content_type == ContentType.TABLE

    def test_has_headers_and_rows(self):
        """Test that the table has headers and rows."""
        elem = MarkdownTableElement.placeholder("MyTable")
        assert len(elem.content.headers) == 3
        assert len(elem.content.rows) == 2

    def test_to_markdown_works(self):
        """Test that to_markdown() produces valid output."""
        elem = MarkdownTableElement.placeholder("MyTable")
        md = elem.to_markdown()
        assert "<!-- table: MyTable -->" in md
        assert "Column A" in md
        assert "|" in md


class TestMarkdownChartElementPlaceholder:
    """Tests for MarkdownChartElement.placeholder()."""

    def test_creates_valid_instance(self):
        """Test that placeholder creates a valid MarkdownChartElement instance."""
        elem = MarkdownChartElement.placeholder("MyChart")
        assert isinstance(elem, MarkdownChartElement)

    def test_name_is_set(self):
        """Test that the name parameter is correctly set."""
        elem = MarkdownChartElement.placeholder("TestChart")
        assert elem.name == "TestChart"

    def test_content_type_is_chart(self):
        """Test that content_type is CHART."""
        elem = MarkdownChartElement.placeholder("MyChart")
        assert elem.content_type == ContentType.CHART

    def test_has_json_content(self):
        """Test that the chart has JSON content."""
        elem = MarkdownChartElement.placeholder("MyChart")
        assert "```json" in elem.content
        assert "```" in elem.content

    def test_metadata_has_chart_data(self):
        """Test that metadata contains parsed chart data."""
        elem = MarkdownChartElement.placeholder("MyChart")
        assert "chart_data" in elem.metadata
        assert elem.metadata["chart_data"]["type"] == "bar"

    def test_to_markdown_works(self):
        """Test that to_markdown() produces valid output."""
        elem = MarkdownChartElement.placeholder("MyChart")
        md = elem.to_markdown()
        assert "<!-- chart: MyChart -->" in md
        assert "```json" in md


class TestMarkdownSlideRoundtrip:
    """Tests for MarkdownSlide roundtrip with placeholder elements."""

    def test_markdown_slide_roundtrip_with_placeholders(self):
        """Test MarkdownSlide with placeholder elements survives to_markdown/from_markdown roundtrip."""
        # Create slide with all placeholder element types
        slide = MarkdownSlide(
            name="Test Slide",
            elements=[
                MarkdownTextElement.placeholder("Text_1"),
                MarkdownImageElement.placeholder("Image_1"),
                MarkdownTableElement.placeholder("Table_1"),
                MarkdownChartElement.placeholder("Chart_1"),
            ],
        )

        # Convert to markdown
        markdown = slide.to_markdown()

        # Parse back from markdown
        parsed_slide = MarkdownSlide.from_markdown(markdown)

        # Convert back to markdown again
        markdown_roundtrip = parsed_slide.to_markdown()

        # Verify perfect reproduction
        assert markdown == markdown_roundtrip

        # Also verify structural equality
        assert parsed_slide.name == slide.name
        assert len(parsed_slide.elements) == len(slide.elements)
        for orig, parsed in zip(slide.elements, parsed_slide.elements):
            assert orig.name == parsed.name
            assert orig.content_type == parsed.content_type

    def test_markdown_slide_roundtrip_text_only(self):
        """Test roundtrip with only text elements."""
        slide = MarkdownSlide(
            name="Text Only",
            elements=[
                MarkdownTextElement.placeholder("Text_A"),
                MarkdownTextElement.placeholder("Text_B"),
            ],
        )
        markdown = slide.to_markdown()
        parsed_slide = MarkdownSlide.from_markdown(markdown)
        markdown_roundtrip = parsed_slide.to_markdown()
        assert markdown == markdown_roundtrip

    def test_markdown_slide_roundtrip_no_slide_name(self):
        """Test roundtrip without a slide name."""
        slide = MarkdownSlide(
            elements=[
                MarkdownTextElement.placeholder("Text_1"),
                MarkdownImageElement.placeholder("Image_1"),
            ],
        )
        markdown = slide.to_markdown()
        parsed_slide = MarkdownSlide.from_markdown(markdown)
        markdown_roundtrip = parsed_slide.to_markdown()
        assert markdown == markdown_roundtrip
