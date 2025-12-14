"""Tests for the bold text with trailing spaces bug."""

import pytest

from gslides_api.agnostic.converters import gslides_style_to_full, text_elements_to_ir
from gslides_api.agnostic.ir_to_markdown import ir_to_markdown, _format_run_to_markdown
from gslides_api.markdown.from_markdown import markdown_to_text_elements
from gslides_api.domain.text import TextElement, TextRun, TextStyle
from gslides_api.request.request import InsertTextRequest


def _apply_markdown_formatting(content: str, style: TextStyle) -> str:
    """Wrapper to convert old-style calls to new IR-based formatting.

    This provides backward compatibility for tests.
    """
    full_style = gslides_style_to_full(style)
    return _format_run_to_markdown(content, full_style)


def text_elements_to_markdown(elements):
    """Wrapper to convert old-style calls to new IR-based conversion.

    This provides backward compatibility for tests.
    """
    ir_doc = text_elements_to_ir(elements)
    return ir_to_markdown(ir_doc)


class TestBoldTrailingSpaceBug:
    """Test cases for the bold text with trailing spaces bug."""

    def test_apply_markdown_formatting_bold_with_trailing_spaces(self):
        """Scenario C: Bold text with trailing spaces should produce valid markdown."""
        style = TextStyle(bold=True)
        content = "This is my title  "  # Two trailing spaces

        result = _apply_markdown_formatting(content, style)

        # Should produce: "**This is my title**  " (spaces outside markers)
        assert result == "**This is my title**  ", f"Got: {result!r}"
        assert not result.endswith("** "), "Spaces should be OUTSIDE the closing **"

    def test_whitespace_only_with_bold_style_not_doubled(self):
        """Whitespace-only content with bold should not be doubled or formatted."""
        style = TextStyle(bold=True)

        # Test pure spaces
        result = _apply_markdown_formatting("  ", style)
        assert result == "  ", f"Pure spaces should return as-is, got: {result!r}"
        assert "**" not in result

        # Test mixed whitespace (this is where the bug manifests)
        result = _apply_markdown_formatting("\t ", style)
        assert result == "\t ", f"Mixed whitespace should return as-is, got: {result!r}"
        assert len(result) == 2, f"Whitespace should not be doubled, got length {len(result)}"

    def test_text_elements_to_markdown_bold_trailing_spaces(self):
        """Test conversion of TextElements with bold and trailing spaces."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=18,
                textRun=TextRun(
                    content="This is my title  ",
                    style=TextStyle(bold=True)
                )
            )
        ]

        result = text_elements_to_markdown(elements)

        assert "**This is my title**" in result
        assert "title **" not in result  # Spaces should NOT be inside markers

    def test_split_text_elements_bold_then_spaces(self):
        """Test when Slides API splits bold text and trailing spaces into separate runs."""
        elements = [
            TextElement(
                startIndex=0,
                endIndex=16,
                textRun=TextRun(
                    content="This is my title",
                    style=TextStyle(bold=True)
                )
            ),
            TextElement(
                startIndex=16,
                endIndex=18,
                textRun=TextRun(
                    content="  ",
                    style=TextStyle(bold=True)
                )
            )
        ]

        result = text_elements_to_markdown(elements)

        assert "**This is my title**" in result
        assert "****" not in result, f"Adjacent asterisks found: {result!r}"
        assert "** **" not in result, f"Empty bold found: {result!r}"

    def test_roundtrip_bold_with_trailing_spaces(self):
        """Scenario B: Read bold text with trailing spaces, write it back."""
        original_elements = [
            TextElement(
                startIndex=0,
                endIndex=18,
                textRun=TextRun(
                    content="This is my title  ",
                    style=TextStyle(bold=True)
                )
            )
        ]

        # Convert to markdown (simulating read)
        markdown = text_elements_to_markdown(original_elements)

        # Parse markdown back to elements (simulating write)
        requests = markdown_to_text_elements(markdown)

        # Extract text from InsertTextRequest objects
        insert_requests = [r for r in requests if isinstance(r, InsertTextRequest)]
        full_text = "".join(r.text for r in insert_requests)

        # Text should NOT contain literal asterisks
        assert "**" not in full_text, f"Literal asterisks in output: {full_text!r}"
        assert "This is my title" in full_text


class TestWhitespaceEdgeCases:
    """Additional edge case tests for whitespace handling."""

    def test_tabs_only_with_bold(self):
        """Tabs-only content with bold should not be formatted."""
        style = TextStyle(bold=True)
        result = _apply_markdown_formatting("\t\t", style)
        assert result == "\t\t", f"Got: {result!r}"
        assert "**" not in result

    def test_leading_and_trailing_spaces_with_text(self):
        """Text with both leading and trailing spaces."""
        style = TextStyle(bold=True)
        content = "  bold text  "

        result = _apply_markdown_formatting(content, style)
        assert result == "  **bold text**  ", f"Got: {result!r}"

    def test_single_space_with_bold(self):
        """Single space with bold style."""
        style = TextStyle(bold=True)
        result = _apply_markdown_formatting(" ", style)
        assert result == " ", f"Got: {result!r}"
        assert "**" not in result

    def test_newline_handling_with_bold(self):
        """Test bold text ending with newline."""
        style = TextStyle(bold=True)
        result = _apply_markdown_formatting("Bold text\n", style)
        assert result == "**Bold text**\n", f"Got: {result!r}"
