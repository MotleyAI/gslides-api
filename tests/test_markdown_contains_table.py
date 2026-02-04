"""Tests for markdown_contains_table utility function."""

import pytest

from gslides_api.agnostic.markdown_parser import markdown_contains_table


class TestMarkdownContainsTable:
    """Tests for the markdown_contains_table function."""

    def test_detects_simple_table(self):
        """Should detect a simple markdown table."""
        content = "| A | B |\n|---|---|\n| 1 | 2 |"
        assert markdown_contains_table(content) is True

    def test_detects_table_with_surrounding_text(self):
        """Should detect a table even when surrounded by other content."""
        content = """
Some introductory text.

| Column 1 | Column 2 |
|----------|----------|
| Value 1  | Value 2  |

Some trailing text.
"""
        assert markdown_contains_table(content) is True

    def test_returns_false_for_plain_text(self):
        """Should return False for plain text without tables."""
        content = "Just some regular text without any tables."
        assert markdown_contains_table(content) is False

    def test_returns_false_for_bullet_list(self):
        """Should return False for bullet lists (not tables)."""
        content = "* Item 1\n* Item 2\n* Item 3"
        assert markdown_contains_table(content) is False

    def test_returns_false_for_numbered_list(self):
        """Should return False for numbered lists."""
        content = "1. First item\n2. Second item\n3. Third item"
        assert markdown_contains_table(content) is False

    def test_returns_false_for_empty_string(self):
        """Should return False for empty string."""
        assert markdown_contains_table("") is False

    def test_returns_false_for_none(self):
        """Should return False for None input."""
        assert markdown_contains_table(None) is False

    def test_returns_false_for_headings(self):
        """Should return False for markdown headings."""
        content = "# Heading 1\n## Heading 2\n### Heading 3"
        assert markdown_contains_table(content) is False

    def test_returns_false_for_formatted_text(self):
        """Should return False for bold/italic text."""
        content = "This has **bold** and *italic* and `code`."
        assert markdown_contains_table(content) is False

    def test_detects_table_with_alignment(self):
        """Should detect tables with alignment markers."""
        content = """
| Left | Center | Right |
|:-----|:------:|------:|
| L    |   C    |     R |
"""
        assert markdown_contains_table(content) is True

    def test_returns_false_for_pipe_in_text(self):
        """Should return False when pipes appear but not in table format."""
        content = "This | is | not | a table because there's no header separator."
        assert markdown_contains_table(content) is False

    def test_detects_multi_row_table(self):
        """Should detect tables with multiple rows."""
        content = """
| Name  | Age | City     |
|-------|-----|----------|
| Alice | 30  | New York |
| Bob   | 25  | Boston   |
| Carol | 35  | Chicago  |
"""
        assert markdown_contains_table(content) is True
