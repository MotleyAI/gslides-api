import logging
import pytest

from gslides_api.markdown.domain import (
    MarkdownDeck,
    MarkdownSlide,
)
from gslides_api.markdown.element import (
    ChartElement,
    ContentType,
    ImageElement,
    MarkdownSlideElement,
    TableData,
    TableElement,
    TextElement,
)


@pytest.fixture
def example_md():
    return """
---
# Slide Title

<!-- text: Text_1 -->
## Introduction

Content here...

<!-- text: Details -->
## Details

More content...

<!-- image: Image_1 -->
![Image](https://example.com/image.jpg)

<!-- chart: Chart_1 -->
```json
{
    "data": [1, 2, 3]
}
```

<!-- table: Table_1 -->
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |

---
# Next Slide

<!-- text: Summary -->
## Summary

Final thoughts
"""


class TestContentType:
    def test_content_type_enum_values(self):
        assert ContentType.TEXT == "text"
        assert ContentType.IMAGE == "image"
        assert ContentType.CHART == "chart"
        assert ContentType.TABLE == "table"


class TestTextElement:
    def test_create_element(self):
        element = TextElement(name="Test", content="Some content")
        assert element.name == "Test"
        assert element.content == "Some content"
        assert element.content_type == ContentType.TEXT
        assert element.metadata == {}

    def test_element_with_metadata(self):
        element = TextElement(name="Test", content="Some content", metadata={"key": "value"})
        assert element.metadata == {"key": "value"}

    def test_to_markdown_with_comment(self):
        element = TextElement(name="TestElement", content="## Header\n\nSome content")
        result = element.to_markdown()
        expected = "<!-- text: TestElement -->\n## Header\n\nSome content"
        assert result == expected

    def test_to_markdown_default_text_no_comment(self):
        element = TextElement(name="Default", content="# Title\n\nDefault content")
        result = element.to_markdown()
        expected = "# Title\n\nDefault content"
        assert result == expected

    def test_to_markdown_strips_trailing_whitespace(self):
        element = TextElement(name="Test", content="Content with trailing spaces   \n  ")
        result = element.to_markdown()
        expected = "<!-- text: Test -->\nContent with trailing spaces"
        assert result == expected


class TestImageElement:
    def test_create_valid_image_from_markdown(self):
        element = ImageElement.from_markdown(
            name="Image1", markdown_content="![alt text](https://example.com/image.jpg)"
        )
        assert element.name == "Image1"
        assert element.content_type == ContentType.IMAGE
        assert element.content == "https://example.com/image.jpg"  # URL in content
        assert element.metadata["alt_text"] == "alt text"  # Alt text in metadata
        assert element.metadata["original_markdown"] == "![alt text](https://example.com/image.jpg)"

    def test_create_valid_image_direct(self):
        element = ImageElement(name="Image1", content="![alt text](https://example.com/image.jpg)")
        assert element.name == "Image1"
        assert element.content_type == ContentType.IMAGE
        assert element.content == "https://example.com/image.jpg"  # URL extracted
        assert element.metadata["alt_text"] == "alt text"  # Alt text extracted

    def test_image_round_trip(self):
        original = "![alt text](https://example.com/image.jpg)"
        element = ImageElement.from_markdown("Test", original)
        reconstructed = element.to_markdown()
        assert "<!-- image: Test -->" in reconstructed
        assert original in reconstructed

    def test_invalid_image_content_raises(self):
        with pytest.raises(
            ValueError, match="Image element must contain at least one markdown image"
        ):
            ImageElement.from_markdown(name="BadImage", markdown_content="This is not an image")


class TestTableElement:
    def test_create_valid_table(self):
        table_md = """| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |"""

        element = TableElement(name="Table1", content=table_md)
        assert element.name == "Table1"
        assert element.content_type == ContentType.TABLE
        assert element.content.headers == ["Header 1", "Header 2"]
        assert element.content.rows == [["Cell 1", "Cell 2"]]

    def test_invalid_table_content_raises(self):
        with pytest.raises(ValueError, match="Table element must contain a valid markdown table"):
            TableElement(name="BadTable", content="This is not a table")

    def test_table_to_markdown(self):
        table_data = TableData(headers=["A", "B"], rows=[["1", "2"]])
        element = TableElement(name="Test", content=table_data)
        result = element.to_markdown()
        assert "<!-- table: Test -->" in result
        assert "| A | B |" in result

    def test_table_to_dataframe_functionality(self):
        """Test DataFrame conversion functionality."""
        table_data = TableData(
            headers=["Name", "Age", "City"],
            rows=[["Alice", "25", "NYC"], ["Bob", "30", "SF"], ["Carol", "35", "LA"]],
        )

        try:
            df = table_data.to_dataframe()

            # Test DataFrame properties
            assert list(df.columns) == ["Name", "Age", "City"]
            assert len(df) == 3

            # Test specific values
            assert df.loc[0, "Name"] == "Alice"
            assert df.loc[0, "Age"] == "25"
            assert df.loc[1, "City"] == "SF"
            assert df.loc[2, "Name"] == "Carol"

            # Test that all data is string type (as expected from markdown tables)
            assert all(df.dtypes == "object"), "All columns should be object/string type"

        except ImportError:
            pytest.skip("pandas not available")

    def test_table_element_to_df_functionality(self):
        """Test TableElement.to_df() method."""
        table_md = """| Product | Price | Stock |
|---------|-------|-------|
| Widget  | $10   | 50    |
| Gadget  | $25   | 30    |"""

        element = TableElement(name="Products", content=table_md)

        try:
            df = element.to_df()

            # Test DataFrame properties
            assert list(df.columns) == ["Product", "Price", "Stock"]
            assert len(df) == 2

            # Test values
            assert df.loc[0, "Product"] == "Widget"
            assert df.loc[0, "Price"] == "$10"
            assert df.loc[1, "Stock"] == "30"

        except ImportError:
            pytest.skip("pandas not available")

    def test_table_element_from_df_functionality(self):
        """Test TableElement.from_df() method."""
        try:
            import pandas as pd

            # Create a DataFrame
            data = {
                "Name": ["Alice", "Bob", "Carol"],
                "Age": [25, 30, 35],
                "City": ["NYC", "SF", "LA"],
            }
            df = pd.DataFrame(data)

            # Create TableElement from DataFrame
            element = TableElement.from_df(df, name="People")

            # Test element properties
            assert element.name == "People"
            assert element.content_type == ContentType.TABLE
            assert element.content.headers == ["Name", "Age", "City"]
            assert element.content.rows == [
                ["Alice", "25", "NYC"],
                ["Bob", "30", "SF"],
                ["Carol", "35", "LA"],
            ]
            assert element.metadata == {}

            # Test round-trip conversion
            df_roundtrip = element.to_df()
            assert list(df_roundtrip.columns) == ["Name", "Age", "City"]
            assert len(df_roundtrip) == 3
            assert df_roundtrip.loc[0, "Name"] == "Alice"
            assert df_roundtrip.loc[2, "City"] == "LA"

        except ImportError:
            pytest.skip("pandas not available")

    def test_table_element_from_df_with_metadata(self):
        """Test TableElement.from_df() with custom metadata."""
        try:
            import pandas as pd

            df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
            metadata = {"source": "test_data", "created": "2024"}

            element = TableElement.from_df(df, name="TestTable", metadata=metadata)

            assert element.metadata == {"source": "test_data", "created": "2024"}

        except ImportError:
            pytest.skip("pandas not available")

    def test_table_element_from_df_invalid_input(self):
        """Test TableElement.from_df() with invalid input."""
        try:
            import pandas as pd

            # Test with non-DataFrame input
            with pytest.raises(ValueError, match="Input must be a pandas DataFrame"):
                TableElement.from_df("not a dataframe", name="Test")

        except ImportError:
            pytest.skip("pandas not available")

    def test_table_element_from_df_without_pandas(self):
        """Test TableElement.from_df() when pandas is not available."""
        import sys
        from unittest.mock import patch

        # Mock pandas import to raise ImportError
        with patch.dict("sys.modules", {"pandas": None}):
            with patch("builtins.__import__") as mock_import:

                def side_effect(name, *args, **kwargs):
                    if name == "pandas":
                        raise ImportError("No module named 'pandas'")
                    return __import__(name, *args, **kwargs)

                mock_import.side_effect = side_effect

                with pytest.raises(ImportError, match="pandas is required"):
                    TableElement.from_df(None, name="Test")

    def test_table_with_empty_cells(self):
        """Test DataFrame conversion with empty/missing cells."""
        table_data = TableData(
            headers=["A", "B", "C"],
            rows=[["1", "2", "3"], ["4", "", "6"], ["7"]],  # Empty cell  # Missing cells
        )

        try:
            df = table_data.to_dataframe()

            # Check that empty cells are handled correctly
            assert df.loc[1, "B"] == ""
            # Missing cells should be filled with empty strings by pandas
            assert len(df) == 3

        except ImportError:
            pytest.skip("pandas not available")


class TestChartElement:
    def test_create_valid_chart(self):
        chart_md = """```json
{
    "data": [1, 2, 3]
}
```"""

        element = ChartElement(name="Chart1", content=chart_md)
        assert element.name == "Chart1"
        assert element.content_type == ContentType.CHART
        assert element.metadata["chart_data"] == {"data": [1, 2, 3]}

    def test_invalid_chart_content_raises(self):
        with pytest.raises(
            ValueError, match="Chart element must contain only a ```json code block"
        ):
            ChartElement(name="BadChart", content="This is not a JSON code block")

    def test_invalid_json_raises(self):
        with pytest.raises(ValueError, match="Chart element must contain valid JSON"):
            ChartElement(name="BadJSON", content="```json\n{invalid json\n```")


class TestMarkdownSlide:
    def test_create_empty_slide(self):
        slide = MarkdownSlide()
        assert slide.elements == []

    def test_slide_with_elements(self):
        elements = [
            TextElement(name="Default", content="# Title"),
            ImageElement(name="Image1", content="![alt](url)"),
        ]
        slide = MarkdownSlide(elements=elements)
        assert len(slide.elements) == 2

    def test_to_markdown(self):
        elements = [
            TextElement(name="Default", content="# Slide Title"),
            TextElement(name="Description", content="Some description text"),
        ]
        slide = MarkdownSlide(elements=elements)
        result = slide.to_markdown()
        expected = "# Slide Title\n\n<!-- text: Description -->\nSome description text"
        assert result == expected

    def test_from_markdown_simple(self):
        markdown = "# Title\n\nSome content"
        slide = MarkdownSlide.from_markdown(markdown)

        assert len(slide.elements) == 1
        assert slide.elements[0].name == "Default"
        assert slide.elements[0].content == "# Title\n\nSome content"
        assert slide.elements[0].content_type == ContentType.TEXT

    def test_from_markdown_with_comments(self):
        markdown = """# Title

<!-- text: Section1 -->
## Introduction

Content here

<!-- image: Img1 -->
![Image](url.jpg)"""

        slide = MarkdownSlide.from_markdown(markdown)

        assert len(slide.elements) == 3

        # Default text element
        assert slide.elements[0].name == "Default"
        assert slide.elements[0].content == "# Title"
        assert slide.elements[0].content_type == ContentType.TEXT

        # Section1 text element
        assert slide.elements[1].name == "Section1"
        assert slide.elements[1].content == "## Introduction\n\nContent here"
        assert slide.elements[1].content_type == ContentType.TEXT

        # Image element
        assert slide.elements[2].name == "Img1"
        assert slide.elements[2].content == "url.jpg"  # URL is stored in content
        assert slide.elements[2].content_type == ContentType.IMAGE
        assert slide.elements[2].metadata["alt_text"] == "Image"  # Alt text in metadata

    def test_from_markdown_invalid_element_type_warn(self, caplog):
        markdown = """# Title

<!-- invalid: BadType -->
Some content"""

        with caplog.at_level(logging.WARNING):
            slide = MarkdownSlide.from_markdown(markdown, on_invalid_element="warn")

        assert len(slide.elements) == 2
        assert slide.elements[1].content_type == ContentType.TEXT
        assert "Invalid element type 'invalid'" in caplog.text

    def test_from_markdown_invalid_element_type_raise(self):
        markdown = """# Title

<!-- invalid: BadType -->
Some content"""

        with pytest.raises(ValueError, match="Invalid element type: invalid"):
            MarkdownSlide.from_markdown(markdown, on_invalid_element="raise")

    def test_from_markdown_empty_content_after_comment_ignored(self):
        markdown = """# Title

<!-- text: EmptySection -->


<!-- text: ValidSection -->
Valid content"""

        slide = MarkdownSlide.from_markdown(markdown)

        # Should have Default + ValidSection only (EmptySection ignored due to empty content)
        assert len(slide.elements) == 2
        assert slide.elements[0].name == "Default"
        assert slide.elements[1].name == "ValidSection"


class TestMarkdownDeck:
    def test_create_empty_deck(self):
        deck = MarkdownDeck()
        assert deck.slides == []

    def test_deck_with_slides(self):
        slides = [MarkdownSlide(elements=[TextElement(name="Default", content="# Slide 1")])]
        deck = MarkdownDeck(slides=slides)
        assert len(deck.slides) == 1

    def test_dumps_single_slide(self):
        deck = MarkdownDeck(
            slides=[MarkdownSlide(elements=[TextElement(name="Default", content="# Title")])]
        )
        result = deck.dumps()
        expected = "---\n# Title\n"
        assert result == expected

    def test_dumps_multiple_slides(self):
        deck = MarkdownDeck(
            slides=[
                MarkdownSlide(elements=[TextElement(name="Default", content="# Slide 1")]),
                MarkdownSlide(elements=[TextElement(name="Default", content="# Slide 2")]),
            ]
        )
        result = deck.dumps()
        expected = "---\n# Slide 1\n\n---\n# Slide 2\n"
        assert result == expected

    def test_dumps_empty_slides_filtered_out(self):
        deck = MarkdownDeck(
            slides=[
                MarkdownSlide(elements=[]),  # Empty slide
                MarkdownSlide(elements=[TextElement(name="Default", content="# Valid Slide")]),
            ]
        )
        result = deck.dumps()
        expected = "---\n# Valid Slide\n"
        assert result == expected

    def test_loads_simple(self):
        markdown = "# Single Slide"
        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 1
        assert len(deck.slides[0].elements) == 1
        assert deck.slides[0].elements[0].content == "# Single Slide"

    def test_loads_multiple_slides(self):
        markdown = """# Slide 1

---
# Slide 2

Some content"""

        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2\n\nSome content"

    def test_loads_with_leading_separator(self):
        markdown = """---
# Slide 1

---
# Slide 2"""

        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2"

    def test_loads_ignores_empty_slides(self):
        markdown = """# Slide 1

---

---
# Slide 2"""

        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2"

    def test_loads_passes_invalid_element_option(self):
        markdown = """# Slide 1

<!-- invalid: BadType -->
Content"""

        with pytest.raises(ValueError):
            MarkdownDeck.loads(markdown, on_invalid_element="raise")


class TestSlideNames:
    """Test slide name functionality."""

    def test_slide_with_name(self):
        """Test parsing slide name from comment."""
        markdown = """<!-- slide: Summary -->
# Summary Slide

This is a summary slide with a name."""

        slide = MarkdownSlide.from_markdown(markdown)

        assert slide.name == "Summary"
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "Default"
        assert (
            slide.elements[0].content == "# Summary Slide\n\nThis is a summary slide with a name."
        )

    def test_slide_without_name(self):
        """Test slide without name comment."""
        markdown = """# Regular Slide

This slide has no name."""

        slide = MarkdownSlide.from_markdown(markdown)

        assert slide.name is None
        assert len(slide.elements) == 1
        assert slide.elements[0].content == "# Regular Slide\n\nThis slide has no name."

    def test_slide_name_with_whitespace(self):
        """Test slide name parsing with various whitespace."""
        markdown = """<!--   slide:   My Slide Name   -->
# Slide Content"""

        slide = MarkdownSlide.from_markdown(markdown)

        assert slide.name == "My Slide Name"
        assert len(slide.elements) == 1

    def test_slide_name_only_no_content(self):
        """Test slide with only name comment and no other content."""
        markdown = """<!-- slide: Empty Named Slide -->"""

        slide = MarkdownSlide.from_markdown(markdown)

        assert slide.name == "Empty Named Slide"
        assert len(slide.elements) == 0

    def test_slide_name_to_markdown(self):
        """Test serialization of slide name to markdown."""
        slide = MarkdownSlide(
            name="Test Slide", elements=[TextElement(name="Default", content="# Content")]
        )

        result = slide.to_markdown()
        expected = "<!-- slide: Test Slide -->\n# Content"
        assert result == expected

    def test_slide_name_to_markdown_no_content(self):
        """Test serialization of slide with name but no elements."""
        slide = MarkdownSlide(name="Empty Slide")

        result = slide.to_markdown()
        expected = "<!-- slide: Empty Slide -->"
        assert result == expected

    def test_slide_name_to_markdown_no_name(self):
        """Test serialization of slide without name."""
        slide = MarkdownSlide(elements=[TextElement(name="Default", content="# Content")])

        result = slide.to_markdown()
        expected = "# Content"
        assert result == expected

    def test_deck_with_named_slides(self):
        """Test deck with multiple named slides."""
        markdown = """---
<!-- slide: Introduction -->
# Welcome

Introduction content

---
<!-- slide: Details -->
# Details

Detail content

---
# Unnamed Slide

Regular content"""

        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 3
        assert deck.slides[0].name == "Introduction"
        assert deck.slides[1].name == "Details"
        assert deck.slides[2].name is None

    def test_deck_with_empty_named_slide(self):
        """Test that empty named slides are preserved."""
        markdown = """---
<!-- slide: Empty -->

---
# Regular Slide

Content"""

        deck = MarkdownDeck.loads(markdown)

        assert len(deck.slides) == 2
        assert deck.slides[0].name == "Empty"
        assert len(deck.slides[0].elements) == 0
        assert deck.slides[1].name is None
        assert len(deck.slides[1].elements) == 1


class TestSlideNameRoundTrip:
    """Test round-trip conversion preserves slide names exactly."""

    def test_named_slide_round_trip(self):
        """Test that named slides round-trip exactly."""
        original = """---
<!-- slide: Summary -->
# Summary

This is the summary.

---
<!-- slide: Details -->
# Details

These are the details.
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output

    def test_mixed_named_unnamed_slides_round_trip(self):
        """Test mixed named and unnamed slides."""
        original = """---
<!-- slide: First -->
# First Slide

Named slide content

---
# Second Slide

Unnamed slide content

---
<!-- slide: Third -->
# Third Slide

Another named slide
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output

    def test_empty_named_slide_round_trip(self):
        """Test empty named slide round-trip."""
        original = """---
<!-- slide: Empty -->

---
<!-- slide: Not Empty -->
# Content

Some content
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output


class TestFullCycleRoundTrip:
    """Test that loading and dumping preserves content exactly."""

    def test_example_md_round_trip(self, example_md):
        """Test that example_md can be loaded and dumped back to identical content."""
        deck = MarkdownDeck.loads(example_md)
        output = deck.dumps()

        # Strip whitespace for comparison (allow minor formatting differences)
        assert example_md.strip() == output.strip()

    def test_simple_content_round_trip(self):
        original = "---\n# Simple Slide\n\nJust some text\n"

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output

    def test_complex_content_round_trip(self):
        original = """---
# Slide 1

<!-- text: Intro -->
## Introduction

Welcome to the presentation

<!-- image: Logo -->
![Logo](https://example.com/logo.png)

---
# Slide 2

<!-- table: Data -->
| Name | Value |
|------|-------|
| A    | 1     |
| B    | 2     |
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output

    def test_no_leading_separator_round_trip(self):
        original = """# First Slide

Some content

---
# Second Slide

More content
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        # Should add leading --- when dumping
        expected = "---\n" + original
        assert expected == output

    def test_whitespace_preservation(self):
        """Test that internal whitespace in content is preserved."""
        original = """---
# Title

<!-- text: Content -->
Line 1

Line 2 with   spaces

    Indented line
"""

        deck = MarkdownDeck.loads(original)
        output = deck.dumps()

        assert original == output
