"""Tests for presentation_to_library function."""

from unittest.mock import MagicMock, PropertyMock

import pytest

from gslides_api.element.base import ElementKind
from gslides_api.agnostic.element import (
    ContentType,
    MarkdownTextElement,
    MarkdownChartElement,
    MarkdownTableElement,
    MarkdownImageElement,
    TableData,
)
from gslides_api.presentation_to_library import presentation_to_library, _convert_element


def _create_mock_element(element_type, title=None, text_content="", image_url=None, table_data=None):
    """Create a mock page element."""
    element = MagicMock()
    element.type = element_type
    element.title = title

    if element_type == ElementKind.SHAPE:
        element.read_text = MagicMock(return_value=text_content)

    elif element_type == ElementKind.IMAGE:
        element.image = MagicMock()
        element.image.contentUrl = image_url

    elif element_type == ElementKind.TABLE:
        if table_data:
            element.extract_table_data = MagicMock(return_value=table_data)
        else:
            element.extract_table_data = MagicMock(side_effect=ValueError("No data"))

    return element


def _create_mock_slide(notes_text, elements):
    """Create a mock slide with speaker notes and elements."""
    slide = MagicMock()
    slide.speaker_notes = MagicMock()
    slide.speaker_notes.read_text = MagicMock(return_value=notes_text)
    slide.page_elements_flat = elements
    return slide


def _create_mock_presentation(slides):
    """Create a mock presentation with slides."""
    presentation = MagicMock()
    presentation.slides = slides
    return presentation


class TestPresentationToLibrary:
    """Tests for the presentation_to_library function."""

    def test_empty_presentation_returns_empty_library(self):
        """An empty presentation should return an empty library."""
        presentation = _create_mock_presentation(slides=[])
        library = presentation_to_library(presentation)
        assert len(library.slides) == 0

    def test_presentation_with_none_slides_returns_empty_library(self):
        """A presentation with None slides should return an empty library."""
        presentation = _create_mock_presentation(slides=None)
        library = presentation_to_library(presentation)
        assert len(library.slides) == 0

    def test_slide_without_speaker_notes_is_skipped(self):
        """Slides without speaker notes should be skipped."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="Some text"
        )
        slide = _create_mock_slide(notes_text="", elements=[element])
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert len(library.slides) == 0

    def test_slide_with_whitespace_only_notes_is_skipped(self):
        """Slides with whitespace-only speaker notes should be skipped."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="Some text"
        )
        slide = _create_mock_slide(notes_text="   \n\t  ", elements=[element])
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert len(library.slides) == 0

    def test_element_without_alt_title_is_skipped(self):
        """Elements without alt-title should be skipped."""
        element_with_title = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="HasTitle",
            text_content="With title"
        )
        element_without_title = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title=None,
            text_content="No title"
        )
        slide = _create_mock_slide(
            notes_text="Test Slide",
            elements=[element_with_title, element_without_title]
        )
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert len(library.slides) == 1
        assert len(library.slides[0].elements) == 1
        assert library.slides[0].elements[0].name == "HasTitle"

    def test_element_with_empty_string_alt_title_is_skipped(self):
        """Elements with empty string alt-title should be skipped."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="",
            text_content="Empty title"
        )
        slide = _create_mock_slide(notes_text="Test Slide", elements=[element])
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert len(library.slides) == 0  # No elements with titles, so slide is not added

    def test_slide_name_comes_from_speaker_notes(self):
        """Slide name should be the speaker notes text."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="Text"
        )
        slide = _create_mock_slide(notes_text="My Custom Slide Name", elements=[element])
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert library.slides[0].name == "My Custom Slide Name"

    def test_slide_name_is_stripped(self):
        """Speaker notes text should be stripped of whitespace."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="Text"
        )
        slide = _create_mock_slide(notes_text="  Padded Name  \n", elements=[element])
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)
        assert library.slides[0].name == "Padded Name"


class TestConvertElement:
    """Tests for the _convert_element helper function."""

    def test_shape_converts_to_text_element(self):
        """Shape elements should convert to MarkdownTextElement."""
        element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="# Hello World"
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownTextElement)
        assert result.name == "Title"
        assert result.content == "# Hello World"
        assert result.content_type == ContentType.TEXT

    def test_table_converts_to_table_element(self):
        """Table elements should convert to MarkdownTableElement."""
        table_data = TableData(headers=["A", "B"], rows=[["1", "2"]])
        element = _create_mock_element(
            element_type=ElementKind.TABLE,
            title="Data Table",
            table_data=table_data
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownTableElement)
        assert result.name == "Data Table"
        assert result.content == table_data
        assert result.content_type == ContentType.TABLE

    def test_table_with_extraction_error_gets_default_data(self):
        """Table with extraction error should get default data."""
        element = _create_mock_element(
            element_type=ElementKind.TABLE,
            title="Empty Table",
            table_data=None  # Will cause ValueError
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownTableElement)
        assert result.name == "Empty Table"
        assert result.content.headers == ["Column"]
        assert result.content.rows == [["Data"]]

    def test_image_converts_to_image_element(self):
        """Image elements should convert to MarkdownImageElement."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="My Image",
            image_url="https://example.com/image.png"
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownImageElement)
        assert result.name == "My Image"
        assert result.content == "https://example.com/image.png"
        assert result.content_type == ContentType.IMAGE
        assert result.metadata["alt_text"] == "My Image"

    def test_image_without_url_gets_placeholder(self):
        """Image without contentUrl should get placeholder URL."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="No URL Image",
            image_url=None
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownImageElement)
        assert result.content == "https://via.placeholder.com/400x300"

    def test_image_with_chart_prefix_converts_to_chart_element(self):
        """Image with 'chart' prefix in alt-title should convert to MarkdownChartElement."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="chart_sales",
            image_url="https://example.com/chart.png"
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownChartElement)
        assert result.name == "chart_sales"
        assert result.content == "Chart placeholder"
        assert result.content_type == ContentType.CHART

    def test_chart_prefix_is_case_insensitive_uppercase(self):
        """'CHART' prefix should be detected case-insensitively."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="CHART Revenue",
            image_url="https://example.com/chart.png"
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownChartElement)

    def test_chart_prefix_is_case_insensitive_mixed_case(self):
        """'Chart' prefix should be detected case-insensitively."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Chart - Quarterly Results",
            image_url="https://example.com/chart.png"
        )

        result = _convert_element(element)

        assert isinstance(result, MarkdownChartElement)

    def test_chart_prefix_must_be_at_start(self):
        """'chart' must be at the start of alt-title to be detected."""
        element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Sales chart image",  # 'chart' not at start
            image_url="https://example.com/chart.png"
        )

        result = _convert_element(element)

        # Should be image, not chart
        assert isinstance(result, MarkdownImageElement)
        assert result.content_type == ContentType.IMAGE

    def test_unsupported_element_type_returns_none(self):
        """Unsupported element types should return None."""
        element = _create_mock_element(
            element_type=ElementKind.VIDEO,
            title="Some Video"
        )

        result = _convert_element(element)

        assert result is None

    def test_line_element_returns_none(self):
        """Line elements should return None (unsupported)."""
        element = _create_mock_element(
            element_type=ElementKind.LINE,
            title="Some Line"
        )

        result = _convert_element(element)

        assert result is None


class TestFullConversion:
    """Integration tests for complete presentation conversion."""

    def test_multiple_slides_with_multiple_elements(self):
        """Test conversion of presentation with multiple slides and elements."""
        # Slide 1: Text and Image
        text_element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Title",
            text_content="# Welcome"
        )
        image_element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Logo",
            image_url="https://example.com/logo.png"
        )
        slide1 = _create_mock_slide(
            notes_text="Title Slide",
            elements=[text_element, image_element]
        )

        # Slide 2: Chart and Table
        chart_element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Chart - Sales",
            image_url="https://example.com/chart.png"
        )
        table_element = _create_mock_element(
            element_type=ElementKind.TABLE,
            title="Data",
            table_data=TableData(headers=["Month", "Sales"], rows=[["Jan", "100"]])
        )
        slide2 = _create_mock_slide(
            notes_text="Data Slide",
            elements=[chart_element, table_element]
        )

        presentation = _create_mock_presentation(slides=[slide1, slide2])
        library = presentation_to_library(presentation)

        assert len(library.slides) == 2

        # Check first slide
        assert library.slides[0].name == "Title Slide"
        assert len(library.slides[0].elements) == 2
        assert library.slides[0].elements[0].content_type == ContentType.TEXT
        assert library.slides[0].elements[1].content_type == ContentType.IMAGE

        # Check second slide
        assert library.slides[1].name == "Data Slide"
        assert len(library.slides[1].elements) == 2
        assert library.slides[1].elements[0].content_type == ContentType.CHART
        assert library.slides[1].elements[1].content_type == ContentType.TABLE

    def test_no_any_content_types_in_output(self):
        """Ensure no elements have ContentType.ANY in the output."""
        text_element = _create_mock_element(
            element_type=ElementKind.SHAPE,
            title="Text",
            text_content="Content"
        )
        image_element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Image",
            image_url="https://example.com/img.png"
        )
        chart_element = _create_mock_element(
            element_type=ElementKind.IMAGE,
            title="Chart Data",
            image_url="https://example.com/chart.png"
        )
        table_element = _create_mock_element(
            element_type=ElementKind.TABLE,
            title="Table",
            table_data=TableData(headers=["A"], rows=[["1"]])
        )

        slide = _create_mock_slide(
            notes_text="Test",
            elements=[text_element, image_element, chart_element, table_element]
        )
        presentation = _create_mock_presentation(slides=[slide])

        library = presentation_to_library(presentation)

        for slide in library.slides:
            for element in slide.elements:
                assert element.content_type != ContentType.ANY, \
                    f"Element {element.name} has ANY content type"
