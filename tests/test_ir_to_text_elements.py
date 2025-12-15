"""
Tests for the _ir_to_text_elements function in gslides_api.markdown.from_markdown.

This file tests the conversion from platform-agnostic IR to Google Slides TextElements.
"""

import pytest

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.text import FullTextStyle, MarkdownRenderableStyle, RichStyle
from gslides_api.domain.text import TextStyle
from gslides_api.markdown.from_markdown import (
    BulletPointGroup,
    LineBreakAfterParagraph,
    ListItemTab,
    NumberedListGroup,
    _ir_to_text_elements,
)


class TestIrToTextElementsBasic:
    """Test basic IR to TextElement conversion."""

    def test_empty_document(self):
        """Test that an empty IR document produces no elements."""
        doc = FormattedDocument(elements=[])
        result = _ir_to_text_elements(doc)
        assert result == []

    def test_single_paragraph_single_run(self):
        """Test a simple paragraph with a single run."""
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Hello", style=FullTextStyle())]
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        # Should produce: TextElement + LineBreakAfterParagraph
        assert len(result) == 2
        assert result[0].textRun.content == "Hello"
        assert isinstance(result[1], LineBreakAfterParagraph)
        assert result[1].textRun.content == "\n"

    def test_paragraph_with_multiple_runs(self):
        """Test a paragraph with multiple runs."""
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[
                        FormattedTextRun(content="Hello ", style=FullTextStyle()),
                        FormattedTextRun(content="World", style=FullTextStyle()),
                    ]
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        # Should produce: TextElement, TextElement, LineBreakAfterParagraph
        assert len(result) == 3
        assert result[0].textRun.content == "Hello "
        assert result[1].textRun.content == "World"
        assert isinstance(result[2], LineBreakAfterParagraph)


class TestIrToTextElementsFormatting:
    """Test IR to TextElement conversion with formatting."""

    def test_bold_text(self):
        """Test that bold formatting is preserved."""
        bold_style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Bold", style=bold_style)]
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        assert result[0].textRun.style.bold is True

    def test_italic_text(self):
        """Test that italic formatting is preserved."""
        italic_style = FullTextStyle(markdown=MarkdownRenderableStyle(italic=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Italic", style=italic_style)]
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        assert result[0].textRun.style.italic is True

    def test_code_span(self):
        """Test that code span formatting is preserved."""
        code_style = FullTextStyle(
            markdown=MarkdownRenderableStyle(is_code=True),
            rich=RichStyle(font_family="Courier New"),
        )
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="code", style=code_style)]
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        assert result[0].textRun.style.fontFamily == "Courier New"


class TestIrToTextElementsLists:
    """Test IR to TextElement conversion with lists."""

    def test_unordered_list_single_item(self):
        """Test an unordered list with a single item."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Item 1", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=0,
                        )
                    ],
                    ordered=False,
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        # Should produce: ListItemTab, TextElement, LineBreakAfterParagraph, BulletPointGroup
        tabs = [e for e in result if isinstance(e, ListItemTab)]
        assert len(tabs) == 1  # One tab for nesting level 0

        text_elements = [
            e
            for e in result
            if not isinstance(e, (ListItemTab, LineBreakAfterParagraph, BulletPointGroup))
        ]
        assert len(text_elements) == 1
        assert text_elements[0].textRun.content == "Item 1"

        bullet_groups = [e for e in result if isinstance(e, BulletPointGroup)]
        assert len(bullet_groups) == 1

    def test_ordered_list_single_item(self):
        """Test an ordered list with a single item."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Item 1", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=0,
                        )
                    ],
                    ordered=True,
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        numbered_groups = [e for e in result if isinstance(e, NumberedListGroup)]
        assert len(numbered_groups) == 1

    def test_list_with_multiple_items(self):
        """Test a list with multiple items."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Item 1", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=0,
                        ),
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Item 2", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=0,
                        ),
                    ],
                    ordered=False,
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        # Should have 2 tabs (one per item)
        tabs = [e for e in result if isinstance(e, ListItemTab)]
        assert len(tabs) == 2

        # Should have 2 text elements
        text_elements = [
            e
            for e in result
            if not isinstance(e, (ListItemTab, LineBreakAfterParagraph, BulletPointGroup))
        ]
        assert len(text_elements) == 2

    def test_nested_list(self):
        """Test a list with nested items."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Level 0", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=0,
                        ),
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[
                                        FormattedTextRun(
                                            content="Level 1", style=FullTextStyle()
                                        )
                                    ]
                                )
                            ],
                            nesting_level=1,
                        ),
                    ],
                    ordered=False,
                )
            ]
        )
        result = _ir_to_text_elements(doc)

        # First item: 1 tab (nesting_level 0 + 1)
        # Second item: 2 tabs (nesting_level 1 + 1)
        tabs = [e for e in result if isinstance(e, ListItemTab)]
        assert len(tabs) == 3  # 1 + 2


class TestIrToTextElementsEmptyParagraph:
    """Test IR to TextElement conversion with empty paragraphs."""

    def test_empty_paragraph(self):
        """Test that empty paragraphs are handled correctly."""
        doc = FormattedDocument(
            elements=[FormattedParagraph(runs=[])]
        )
        result = _ir_to_text_elements(doc)

        # Should produce just a LineBreakAfterParagraph
        assert len(result) == 1
        assert isinstance(result[0], LineBreakAfterParagraph)


class TestIrToTextElementsWithBaseStyle:
    """Test IR to TextElement conversion with custom base style."""

    def test_base_style_applied_to_line_breaks(self):
        """Test that base style is applied to line breaks."""
        base_style = TextStyle(fontFamily="Arial")
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Text", style=FullTextStyle())]
                )
            ]
        )
        result = _ir_to_text_elements(doc, base_style=base_style)

        # The LineBreakAfterParagraph should have the base style
        line_breaks = [e for e in result if isinstance(e, LineBreakAfterParagraph)]
        assert len(line_breaks) == 1
        assert line_breaks[0].textRun.style.fontFamily == "Arial"
