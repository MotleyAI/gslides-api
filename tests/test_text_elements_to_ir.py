"""Tests for text_elements_to_ir() converter.

This file tests the conversion from Google Slides TextElements to
platform-agnostic IR (FormattedDocument).
"""

import pytest

from gslides_api.agnostic.converters import text_elements_to_ir
from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedParagraph,
)
from gslides_api.domain.text import (
    Bullet,
    ParagraphMarker,
    TextElement,
    TextRun,
    TextStyle,
)


class TestTextElementsToIrBasic:
    """Basic conversion tests."""

    def test_empty_elements(self):
        """Empty list should return empty document."""
        result = text_elements_to_ir([])
        assert isinstance(result, FormattedDocument)
        assert len(result.elements) == 0

    def test_single_plain_text(self):
        """Single plain text element."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=5,
                textRun=TextRun(content="Hello", style=TextStyle()),
            )
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedParagraph)
        assert len(result.elements[0].runs) == 1
        assert result.elements[0].runs[0].content == "Hello"

    def test_text_with_newline_creates_paragraph(self):
        """Text ending with newline should complete a paragraph."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=6,
                textRun=TextRun(content="Hello\n", style=TextStyle()),
            )
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedParagraph)
        # Content should include the newline
        assert result.elements[0].runs[0].content == "Hello\n"


class TestTextElementsToIrFormatting:
    """Tests for text formatting conversion."""

    def test_bold_text(self):
        """Bold text should preserve style."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=5,
                textRun=TextRun(content="Bold", style=TextStyle(bold=True)),
            )
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert result.elements[0].runs[0].style.markdown.bold is True

    def test_italic_text(self):
        """Italic text should preserve style."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=6,
                textRun=TextRun(content="Italic", style=TextStyle(italic=True)),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].style.markdown.italic is True

    def test_strikethrough_text(self):
        """Strikethrough text should preserve style."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=7,
                textRun=TextRun(content="Striked", style=TextStyle(strikethrough=True)),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].style.markdown.strikethrough is True

    def test_bold_and_italic(self):
        """Combined bold and italic should preserve both."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=4,
                textRun=TextRun(
                    content="Both", style=TextStyle(bold=True, italic=True)
                ),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].style.markdown.bold is True
        assert result.elements[0].runs[0].style.markdown.italic is True


class TestTextElementsToIrHyperlinks:
    """Tests for hyperlink conversion."""

    def test_hyperlink(self):
        """Hyperlink should be converted to IR."""
        from gslides_api.domain.text import Link

        elements = [
            TextElement(
                startIndex=0,
                endIndex=4,
                textRun=TextRun(
                    content="Link",
                    style=TextStyle(link=Link(url="https://example.com")),
                ),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].style.markdown.hyperlink == "https://example.com"


class TestTextElementsToIrCodeSpans:
    """Tests for code span (monospace font) conversion."""

    def test_monospace_font_detected_as_code(self):
        """Monospace font should be detected as code."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=4,
                textRun=TextRun(
                    content="code", style=TextStyle(fontFamily="Courier New")
                ),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].style.markdown.is_code is True


class TestTextElementsToIrLists:
    """Tests for list conversion."""

    def test_unordered_list(self):
        """Bullet list should be converted to unordered list."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=0,
                endIndex=5,
                textRun=TextRun(content="Item\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedList)
        assert result.elements[0].ordered is False
        assert len(result.elements[0].items) == 1
        assert result.elements[0].items[0].paragraphs[0].runs[0].content == "Item"

    def test_ordered_list(self):
        """Numbered list should be converted to ordered list."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="1.")
                ),
            ),
            TextElement(
                startIndex=0,
                endIndex=6,
                textRun=TextRun(content="First\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedList)
        assert result.elements[0].ordered is True

    def test_nested_list(self):
        """Nested list should preserve nesting level."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=0,
                endIndex=7,
                textRun=TextRun(content="Parent\n", style=TextStyle()),
            ),
            TextElement(
                startIndex=7,
                endIndex=7,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=1, glyph="●")
                ),
            ),
            TextElement(
                startIndex=7,
                endIndex=13,
                textRun=TextRun(content="Child\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedList)
        assert len(result.elements[0].items) == 2
        assert result.elements[0].items[0].nesting_level == 0
        assert result.elements[0].items[1].nesting_level == 1

    def test_multiple_list_items(self):
        """Multiple list items should be captured."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=0,
                endIndex=5,
                textRun=TextRun(content="One\n", style=TextStyle()),
            ),
            TextElement(
                startIndex=5,
                endIndex=5,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=5,
                endIndex=9,
                textRun=TextRun(content="Two\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert isinstance(result.elements[0], FormattedList)
        assert len(result.elements[0].items) == 2


class TestTextElementsToIrMultipleRuns:
    """Tests for multiple text runs in a paragraph."""

    def test_adjacent_runs_same_style(self):
        """Adjacent runs with same style should be preserved as separate runs."""
        style = TextStyle(bold=True)
        elements = [
            TextElement(
                startIndex=0,
                endIndex=6,
                textRun=TextRun(content="Hello ", style=style),
            ),
            TextElement(
                startIndex=6,
                endIndex=12,
                textRun=TextRun(content="World\n", style=style),
            ),
        ]
        result = text_elements_to_ir(elements)
        # Both runs should be in the same paragraph
        assert len(result.elements) == 1
        assert len(result.elements[0].runs) == 2
        # Consolidation happens in ir_to_markdown, not here

    def test_mixed_formatting_in_paragraph(self):
        """Mixed formatting should create multiple runs."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=7,
                textRun=TextRun(content="Normal ", style=TextStyle()),
            ),
            TextElement(
                startIndex=7,
                endIndex=11,
                textRun=TextRun(content="bold", style=TextStyle(bold=True)),
            ),
            TextElement(
                startIndex=11,
                endIndex=19,
                textRun=TextRun(content=" normal\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 1
        assert len(result.elements[0].runs) == 3
        assert result.elements[0].runs[0].style.markdown.bold is False
        assert result.elements[0].runs[1].style.markdown.bold is True
        assert result.elements[0].runs[2].style.markdown.bold is False


class TestTextElementsToIrTrailingSpaces:
    """Tests for trailing space handling."""

    def test_bold_with_trailing_spaces(self):
        """Bold text with trailing spaces should preserve content."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=10,
                textRun=TextRun(content="Bold text  ", style=TextStyle(bold=True)),
            )
        ]
        result = text_elements_to_ir(elements)
        assert result.elements[0].runs[0].content == "Bold text  "
        assert result.elements[0].runs[0].style.markdown.bold is True

    def test_split_bold_and_spaces(self):
        """Bold text split from trailing spaces (API behavior) should be handled."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=9,
                textRun=TextRun(content="Bold text", style=TextStyle(bold=True)),
            ),
            TextElement(
                startIndex=9,
                endIndex=11,
                textRun=TextRun(content="  ", style=TextStyle(bold=True)),
            ),
        ]
        result = text_elements_to_ir(elements)
        # Both runs should be preserved
        assert len(result.elements[0].runs) == 2
        assert result.elements[0].runs[0].content == "Bold text"
        assert result.elements[0].runs[1].content == "  "


class TestTextElementsToIrMixedContent:
    """Tests for mixed paragraphs and lists."""

    def test_paragraph_then_list(self):
        """Paragraph followed by list."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(),  # No bullet
            ),
            TextElement(
                startIndex=0,
                endIndex=6,
                textRun=TextRun(content="Intro\n", style=TextStyle()),
            ),
            TextElement(
                startIndex=6,
                endIndex=6,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=6,
                endIndex=12,
                textRun=TextRun(content="Item\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 2
        assert isinstance(result.elements[0], FormattedParagraph)
        assert isinstance(result.elements[1], FormattedList)

    def test_list_then_paragraph(self):
        """List followed by paragraph."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=0,
                paragraphMarker=ParagraphMarker(
                    bullet=Bullet(listId="list1", nestingLevel=0, glyph="●")
                ),
            ),
            TextElement(
                startIndex=0,
                endIndex=5,
                textRun=TextRun(content="Item\n", style=TextStyle()),
            ),
            TextElement(
                startIndex=5,
                endIndex=5,
                paragraphMarker=ParagraphMarker(),  # No bullet
            ),
            TextElement(
                startIndex=5,
                endIndex=12,
                textRun=TextRun(content="Outro\n", style=TextStyle()),
            ),
        ]
        result = text_elements_to_ir(elements)
        assert len(result.elements) == 2
        assert isinstance(result.elements[0], FormattedList)
        assert isinstance(result.elements[1], FormattedParagraph)
