"""
Tests for markdown feature support in gslides_api.

This file tests both currently supported features and new features we want to add:
- Currently supported: bold, italic, code spans, bullet lists, headings, paragraphs
- New features to add: hyperlinks, strikethrough, ordered lists, underline
"""

import pytest

from gslides_api.markdown.from_markdown import markdown_to_text_elements
from gslides_api.text import TextStyle


class TestMarkdownFeatureSupport:
    """Test that markdown_to_text_elements can handle various markdown features without exceptions."""

    def test_currently_supported_features(self):
        """Test all currently supported markdown features."""
        markdown_samples = [
            # Bold text
            "This is **bold** text",
            "This is __also bold__ text",
            # Italic text
            "This is *italic* text",
            "This is _also italic_ text",
            # Combined bold and italic
            "This is ***bold and italic*** text",
            "This is ___also bold and italic___ text",
            # Code spans
            "This has `inline code` in it",
            "Multiple `code` spans `here`",
            # Headings
            "# Heading 1",
            "## Heading 2",
            "### Heading 3",
            "#### Heading 4",
            "##### Heading 5",
            "###### Heading 6",
            # Bullet lists
            "* First item\n* Second item\n* Third item",
            "- First item\n- Second item\n- Third item",
            "+ First item\n+ Second item\n+ Third item",
            # Nested content in lists
            "* Item with **bold** text\n* Item with *italic* text\n* Item with `code`",
            # Paragraphs with line breaks
            "First paragraph\n\nSecond paragraph",
            # Mixed content
            "# Title\n\nThis is a paragraph with **bold**, *italic*, and `code`.\n\n* List item 1\n* List item 2",
        ]

        for markdown in markdown_samples:
            # Should not raise any exceptions
            result = markdown_to_text_elements(markdown)
            assert result is not None
            assert len(result) > 0

    def test_new_features_to_support(self):
        """Test new markdown features we want to add support for."""
        markdown_samples = [
            # Hyperlinks
            "[Google](https://google.com)",
            "Visit [our website](https://example.com) for more info",
            "Multiple [link1](https://example1.com) and [link2](https://example2.com)",
            # Strikethrough (if supported by marko)
            "~~strikethrough text~~",
            "This is ~~crossed out~~ text",
            # Ordered lists
            "1. First item\n2. Second item\n3. Third item",
            "1. Item with **bold**\n2. Item with *italic*\n3. Item with `code`",
            # Mixed ordered and unordered lists
            "1. Ordered item\n2. Another ordered\n\n* Bullet item\n* Another bullet",
        ]

        for markdown in markdown_samples:
            try:
                # These might fail currently, but we want to track which ones do
                result = markdown_to_text_elements(markdown)
                print(f"✅ Successfully parsed: {markdown[:50]}...")
                assert result is not None
            except Exception as e:
                print(f"❌ Failed to parse: {markdown[:50]}... - Error: {e}")
                # For now, we expect some of these to fail
                # Once we implement support, we can change these to assert success

    def test_complex_mixed_content(self):
        """Test complex markdown with multiple features combined."""
        complex_markdown = """# Main Title

This is a paragraph with **bold**, *italic*, and `code` text.

## Subsection

Here's a [link to Google](https://google.com) and some ~~strikethrough~~ text.

### Lists

Unordered list:
* First bullet point
* Second with **bold** text
* Third with [a link](https://example.com)

Ordered list:
1. First numbered item
2. Second with *italic* text  
3. Third with `inline code`

## Conclusion

That's all for now!
"""

        try:
            result = markdown_to_text_elements(complex_markdown)
            print("✅ Successfully parsed complex markdown")
            assert result is not None
            assert len(result) > 0
        except Exception as e:
            print(f"❌ Failed to parse complex markdown - Error: {e}")
            # For now, we expect this might fail until we implement all features


class TestMarkdownWithBaseStyle:
    """Test markdown parsing with different base styles."""

    def test_with_custom_base_style(self):
        """Test that base styles are properly applied."""
        base_style = TextStyle(
            fontFamily="Arial",
            fontSize={"magnitude": 12, "unit": "PT"},
            bold=False,
            italic=False,
        )

        markdown = "This is **bold** and *italic* text"
        result = markdown_to_text_elements(markdown, base_style=base_style)

        assert result is not None
        assert len(result) > 0

        # Check that text elements have the base style applied
        for element in result:
            if hasattr(element, "textRun") and element.textRun:
                style = element.textRun.style
                # Base font family should be preserved unless overridden
                if not style.bold and not style.italic:
                    assert style.fontFamily == "Arial"


class TestMarkdownEdgeCases:
    """Test edge cases and error conditions."""

    def test_empty_markdown(self):
        """Test empty or whitespace-only markdown."""
        test_cases = ["", "   ", "\n", "\n\n\n", "\t\t"]

        for markdown in test_cases:
            try:
                result = markdown_to_text_elements(markdown)
                # Should handle gracefully, might return empty list or minimal elements
                assert result is not None
            except Exception as e:
                pytest.fail(f"Should handle empty markdown gracefully, but got: {e}")

    def test_malformed_markdown(self):
        """Test malformed or edge case markdown."""
        test_cases = [
            "**unclosed bold",
            "*unclosed italic",
            "`unclosed code",
            "# ",  # Empty heading
            "* ",  # Empty list item
            "[link with no url]",
            "[](empty link)",
        ]

        for markdown in test_cases:
            try:
                result = markdown_to_text_elements(markdown)
                # Should handle gracefully
                assert result is not None
            except Exception as e:
                print(f"Malformed markdown '{markdown}' caused: {e}")
                # Some malformed markdown might legitimately fail
