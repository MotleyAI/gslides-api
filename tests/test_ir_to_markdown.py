"""Tests for IR to markdown conversion with run consolidation."""

import pytest

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.ir_to_markdown import (
    ir_to_markdown,
    _consolidate_runs,
    _format_run_to_markdown,
    _same_markdown_style,
)
from gslides_api.agnostic.text import (
    FullTextStyle,
    MarkdownRenderableStyle,
    RichStyle,
)


class TestConsolidateRuns:
    """Tests for run consolidation logic."""

    def test_empty_runs(self):
        """Empty list should return empty list."""
        result = _consolidate_runs([])
        assert result == []

    def test_single_run(self):
        """Single run should return as-is."""
        run = FormattedTextRun(content="Hello")
        result = _consolidate_runs([run])
        assert len(result) == 1
        assert result[0].content == "Hello"

    def test_adjacent_bold_runs_consolidated(self):
        """Adjacent bold runs should be merged."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        runs = [
            FormattedTextRun(content="Hello ", style=style),
            FormattedTextRun(content="World", style=style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 1
        assert result[0].content == "Hello World"
        assert result[0].style.markdown.bold is True

    def test_different_formatting_not_consolidated(self):
        """Bold followed by italic should remain separate."""
        bold_style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        italic_style = FullTextStyle(markdown=MarkdownRenderableStyle(italic=True))
        runs = [
            FormattedTextRun(content="Hello ", style=bold_style),
            FormattedTextRun(content="World", style=italic_style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 2
        assert result[0].content == "Hello "
        assert result[0].style.markdown.bold is True
        assert result[1].content == "World"
        assert result[1].style.markdown.italic is True

    def test_plain_bold_plain_sequence(self):
        """Plain-bold-plain sequence should stay as 3 runs."""
        plain_style = FullTextStyle()
        bold_style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        runs = [
            FormattedTextRun(content="Normal ", style=plain_style),
            FormattedTextRun(content="bold", style=bold_style),
            FormattedTextRun(content=" normal", style=plain_style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 3

    def test_three_adjacent_bold_runs_all_consolidated(self):
        """Three consecutive bold runs should merge into one."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        runs = [
            FormattedTextRun(content="One ", style=style),
            FormattedTextRun(content="two ", style=style),
            FormattedTextRun(content="three", style=style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 1
        assert result[0].content == "One two three"

    def test_hyperlink_prevents_consolidation(self):
        """Different hyperlinks should not be consolidated."""
        link1_style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        link2_style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://other.com")
        )
        runs = [
            FormattedTextRun(content="Link1 ", style=link1_style),
            FormattedTextRun(content="Link2", style=link2_style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 2

    def test_same_hyperlink_consolidated(self):
        """Same hyperlink should be consolidated."""
        link_style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        runs = [
            FormattedTextRun(content="Click ", style=link_style),
            FormattedTextRun(content="here", style=link_style),
        ]
        result = _consolidate_runs(runs)
        assert len(result) == 1
        assert result[0].content == "Click here"

    def test_rich_style_differences_do_not_affect_consolidation(self):
        """Rich style differences should NOT affect consolidation.

        Consolidation is based on markdown-renderable style only.
        """
        style1 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=RichStyle(font_size_pt=14.0),
        )
        style2 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=RichStyle(font_size_pt=16.0),  # Different font size
        )
        runs = [
            FormattedTextRun(content="Hello ", style=style1),
            FormattedTextRun(content="World", style=style2),
        ]
        result = _consolidate_runs(runs)
        # Should consolidate because markdown style is the same
        assert len(result) == 1
        assert result[0].content == "Hello World"


class TestSameMarkdownStyle:
    """Tests for markdown style comparison."""

    def test_both_empty(self):
        """Empty styles should be equal."""
        a = MarkdownRenderableStyle()
        b = MarkdownRenderableStyle()
        assert _same_markdown_style(a, b) is True

    def test_both_bold(self):
        """Both bold should be equal."""
        a = MarkdownRenderableStyle(bold=True)
        b = MarkdownRenderableStyle(bold=True)
        assert _same_markdown_style(a, b) is True

    def test_bold_vs_not_bold(self):
        """Bold vs not bold should be different."""
        a = MarkdownRenderableStyle(bold=True)
        b = MarkdownRenderableStyle(bold=False)
        assert _same_markdown_style(a, b) is False

    def test_bold_vs_italic(self):
        """Bold vs italic should be different."""
        a = MarkdownRenderableStyle(bold=True)
        b = MarkdownRenderableStyle(italic=True)
        assert _same_markdown_style(a, b) is False

    def test_bold_and_italic(self):
        """Both bold+italic should be equal."""
        a = MarkdownRenderableStyle(bold=True, italic=True)
        b = MarkdownRenderableStyle(bold=True, italic=True)
        assert _same_markdown_style(a, b) is True


class TestFormatRunToMarkdown:
    """Tests for formatting individual runs to markdown."""

    def test_plain_text(self):
        """Plain text should return as-is."""
        style = FullTextStyle()
        result = _format_run_to_markdown("Hello", style)
        assert result == "Hello"

    def test_bold_text(self):
        """Bold text should be wrapped with **."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown("Hello", style)
        assert result == "**Hello**"

    def test_italic_text(self):
        """Italic text should be wrapped with *."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(italic=True))
        result = _format_run_to_markdown("Hello", style)
        assert result == "*Hello*"

    def test_bold_italic_text(self):
        """Bold+italic should be wrapped with ***."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True, italic=True))
        result = _format_run_to_markdown("Hello", style)
        assert result == "***Hello***"

    def test_strikethrough_text(self):
        """Strikethrough should be wrapped with ~~."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(strikethrough=True))
        result = _format_run_to_markdown("Hello", style)
        assert result == "~~Hello~~"

    def test_hyperlink(self):
        """Hyperlink should use markdown link format."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        result = _format_run_to_markdown("Click here", style)
        assert result == "[Click here](https://example.com)"

    def test_trailing_space_outside_markers(self):
        """Trailing spaces should be outside markers."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown("Hello ", style)
        assert result == "**Hello** "
        # Space should be AFTER closing **, not before (i.e., not "**Hello **")
        assert " **" not in result, "Space should not be inside closing markers"

    def test_leading_space_outside_markers(self):
        """Leading spaces should be outside markers."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown(" Hello", style)
        assert result == " **Hello**"
        assert not result.startswith("** ")

    def test_both_leading_and_trailing_spaces(self):
        """Both leading and trailing spaces should be outside markers."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown("  Hello  ", style)
        assert result == "  **Hello**  "

    def test_whitespace_only_not_formatted(self):
        """Whitespace-only content should not get markers."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown("   ", style)
        assert result == "   "
        assert "**" not in result

    def test_newline_preserved(self):
        """Trailing newline should be preserved."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        result = _format_run_to_markdown("Hello\n", style)
        assert result == "**Hello**\n"

    def test_code_span(self):
        """Code spans should use backticks."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(is_code=True))
        result = _format_run_to_markdown("code", style)
        assert result == "`code`"

    def test_hyperlink_with_trailing_space(self):
        """Hyperlink with trailing space should preserve it."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        result = _format_run_to_markdown("Click here ", style)
        assert result == "[Click here](https://example.com) "

    def test_hyperlink_with_leading_space(self):
        """Hyperlink with leading space should preserve it."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        result = _format_run_to_markdown(" Click here", style)
        assert result == " [Click here](https://example.com)"

    def test_hyperlink_with_both_leading_and_trailing_space(self):
        """Hyperlink with both leading and trailing spaces should preserve them."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        result = _format_run_to_markdown("  Click here  ", style)
        assert result == "  [Click here](https://example.com)  "

    def test_hyperlink_with_trailing_newline(self):
        """Hyperlink with trailing newline should preserve it."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(hyperlink="https://example.com")
        )
        result = _format_run_to_markdown("Click here\n", style)
        assert result == "[Click here](https://example.com)\n"

    def test_code_span_with_trailing_space(self):
        """Code span with trailing space should preserve it."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(is_code=True))
        result = _format_run_to_markdown("code ", style)
        assert result == "`code` "

    def test_code_span_with_leading_space(self):
        """Code span with leading space should preserve it."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(is_code=True))
        result = _format_run_to_markdown(" code", style)
        assert result == " `code`"

    def test_code_span_with_both_leading_and_trailing_space(self):
        """Code span with both leading and trailing spaces should preserve them."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(is_code=True))
        result = _format_run_to_markdown("  code  ", style)
        assert result == "  `code`  "

    def test_code_span_with_trailing_newline(self):
        """Code span with trailing newline should preserve it."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(is_code=True))
        result = _format_run_to_markdown("code\n", style)
        assert result == "`code`\n"


class TestIrToMarkdown:
    """Tests for full IR to markdown conversion."""

    def test_empty_document(self):
        """Empty document should produce empty string."""
        doc = FormattedDocument()
        result = ir_to_markdown(doc)
        assert result == ""

    def test_single_plain_paragraph(self):
        """Single plain paragraph."""
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Hello World")]
                )
            ]
        )
        result = ir_to_markdown(doc)
        assert result == "Hello World"

    def test_single_bold_paragraph(self):
        """Single bold paragraph."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[FormattedTextRun(content="Hello World", style=style)]
                )
            ]
        )
        result = ir_to_markdown(doc)
        assert result == "**Hello World**"

    def test_paragraph_with_adjacent_bold_runs(self):
        """Adjacent bold runs should be consolidated."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[
                        FormattedTextRun(content="Hello ", style=style),
                        FormattedTextRun(content="World", style=style),
                    ]
                )
            ]
        )
        result = ir_to_markdown(doc)
        # Should produce single bold section, not **Hello ****World**
        assert result == "**Hello World**"
        assert "****" not in result

    def test_paragraph_with_bold_then_italic(self):
        """Bold followed by italic should produce separate sections."""
        bold_style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        italic_style = FullTextStyle(markdown=MarkdownRenderableStyle(italic=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[
                        FormattedTextRun(content="Hello ", style=bold_style),
                        FormattedTextRun(content="World", style=italic_style),
                    ]
                )
            ]
        )
        result = ir_to_markdown(doc)
        assert result == "**Hello** *World*"

    def test_multiple_paragraphs(self):
        """Multiple paragraphs should be on separate lines."""
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(runs=[FormattedTextRun(content="First")]),
                FormattedParagraph(runs=[FormattedTextRun(content="Second")]),
            ]
        )
        result = ir_to_markdown(doc)
        assert result == "First\nSecond"

    def test_unordered_list(self):
        """Unordered list should use bullet markers."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    ordered=False,
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[FormattedTextRun(content="Item one")]
                                )
                            ]
                        ),
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[FormattedTextRun(content="Item two")]
                                )
                            ]
                        ),
                    ],
                )
            ]
        )
        result = ir_to_markdown(doc)
        assert "* Item one" in result
        assert "* Item two" in result

    def test_ordered_list(self):
        """Ordered list should use number markers."""
        doc = FormattedDocument(
            elements=[
                FormattedList(
                    ordered=True,
                    items=[
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[FormattedTextRun(content="First")]
                                )
                            ]
                        ),
                        FormattedListItem(
                            paragraphs=[
                                FormattedParagraph(
                                    runs=[FormattedTextRun(content="Second")]
                                )
                            ]
                        ),
                    ],
                )
            ]
        )
        result = ir_to_markdown(doc)
        assert "1. First" in result
        assert "2. Second" in result


class TestBugScenarios:
    """Tests specifically for the bug scenario from the issue."""

    def test_pptx_adjacent_bold_runs_bug(self):
        """Simulate the PPTX bug: adjacent bold runs producing ****."""
        style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[
                        FormattedTextRun(
                            content="Creating new profile tiers for ", style=style
                        ),
                        FormattedTextRun(content="{customer_name}", style=style),
                    ]
                )
            ]
        )
        result = ir_to_markdown(doc)

        # Should NOT have doubled asterisks
        assert "****" not in result, f"Doubled asterisks found: {result}"

        # Should produce single consolidated bold section
        assert "**Creating new profile tiers for {customer_name}**" in result

    def test_mixed_formatting_with_trailing_spaces(self):
        """Mixed formatting with trailing spaces should work correctly."""
        bold_style = FullTextStyle(markdown=MarkdownRenderableStyle(bold=True))
        plain_style = FullTextStyle()

        doc = FormattedDocument(
            elements=[
                FormattedParagraph(
                    runs=[
                        FormattedTextRun(content="Normal text ", style=plain_style),
                        FormattedTextRun(content="bold text", style=bold_style),
                        FormattedTextRun(content=" more normal", style=plain_style),
                    ]
                )
            ]
        )
        result = ir_to_markdown(doc)

        assert result == "Normal text **bold text** more normal"
        assert "** **" not in result
        assert "****" not in result
