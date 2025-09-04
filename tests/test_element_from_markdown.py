"""
Tests for the from_markdown class methods on all MarkdownSlideElement classes.

This file tests the from_markdown methods that create element instances from raw markdown content.
Also tests the new markdown_element_to_requests method for TableElement.
"""

import pytest

from gslides_api.element.table import TableElement as GSlidesTableElement
from gslides_api.markdown.element import (
    MarkdownChartElement,
    ContentType,
    MarkdownImageElement,
    TableData,
    MarkdownTableElement,
    MarkdownTextElement,
)
from gslides_api.request.table import CreateTableRequest
from gslides_api.request.request import InsertTextRequest, UpdateTextStyleRequest
from gslides_api.request.domain import TableCellLocation


class TestTextElementFromMarkdown:
    """Test TextElement.from_markdown() method."""

    def test_simple_text_from_markdown(self):
        """Test creating TextElement from simple text."""
        element = MarkdownTextElement.from_markdown("MyText", "# Hello World")

        assert element.name == "MyText"
        assert element.content == "# Hello World"
        assert element.content_type == ContentType.TEXT
        assert element.metadata == {}

    def test_multiline_text_from_markdown(self):
        """Test creating TextElement from multiline markdown."""
        markdown_content = """# Title

This is a paragraph with **bold** text.

- List item 1
- List item 2

## Subtitle

More content here."""

        element = MarkdownTextElement.from_markdown("Details", markdown_content)

        assert element.name == "Details"
        assert element.content == markdown_content
        assert element.content_type == ContentType.TEXT

    def test_text_with_whitespace_from_markdown(self):
        """Test that TextElement.from_markdown strips whitespace."""
        element = MarkdownTextElement.from_markdown("Test", "  # Hello  \n\n  ")

        assert element.name == "Test"
        assert element.content == "# Hello"
        assert element.content_type == ContentType.TEXT

    def test_empty_text_from_markdown(self):
        """Test creating TextElement from empty content."""
        element = MarkdownTextElement.from_markdown("Empty", "")

        assert element.name == "Empty"
        assert element.content == ""
        assert element.content_type == ContentType.TEXT

    def test_text_from_markdown_round_trip(self):
        """Test that TextElement can be created from markdown and converted back."""
        original_markdown = "# Title\n\nSome **bold** text and *italic* text."
        element = MarkdownTextElement.from_markdown("Test", original_markdown)

        # Convert back to markdown
        result = element.to_markdown()

        # Should include comment for non-default text
        expected = "<!-- text: Test -->\n# Title\n\nSome **bold** text and *italic* text."
        assert result == expected


class TestImageElementFromMarkdown:
    """Test ImageElement.from_markdown() method."""

    def test_simple_image_from_markdown(self):
        """Test creating ImageElement from simple image markdown."""
        markdown = "![Alt text](https://example.com/image.jpg)"
        element = MarkdownImageElement.from_markdown("MyImage", markdown)

        assert element.name == "MyImage"
        assert element.content == "https://example.com/image.jpg"
        assert element.content_type == ContentType.IMAGE
        assert element.metadata["alt_text"] == "Alt text"
        assert element.metadata["original_markdown"] == markdown

    def test_image_without_alt_text_from_markdown(self):
        """Test creating ImageElement with empty alt text."""
        markdown = "![](https://example.com/image.png)"
        element = MarkdownImageElement.from_markdown("EmptyAlt", markdown)

        assert element.name == "EmptyAlt"
        assert element.content == "https://example.com/image.png"
        assert element.metadata["alt_text"] == ""
        assert element.metadata["original_markdown"] == markdown

    def test_image_with_complex_url_from_markdown(self):
        """Test creating ImageElement with complex URL."""
        markdown = "![Complex](https://example.com/path/to/image.jpg?param=value&size=large)"
        element = MarkdownImageElement.from_markdown("Complex", markdown)

        assert element.content == "https://example.com/path/to/image.jpg?param=value&size=large"
        assert element.metadata["alt_text"] == "Complex"

    def test_image_with_local_path_from_markdown(self):
        """Test creating ImageElement with local file path."""
        markdown = "![Local image](./assets/image.png)"
        element = MarkdownImageElement.from_markdown("Local", markdown)

        assert element.content == "./assets/image.png"
        assert element.metadata["alt_text"] == "Local image"

    def test_image_with_whitespace_from_markdown(self):
        """Test creating ImageElement with surrounding whitespace."""
        markdown = "  ![Test](https://example.com/test.jpg)  \n"
        element = MarkdownImageElement.from_markdown("Whitespace", markdown)

        assert element.content == "https://example.com/test.jpg"
        assert element.metadata["original_markdown"] == "![Test](https://example.com/test.jpg)"

    def test_invalid_image_markdown_raises(self):
        """Test that invalid image markdown raises ValueError."""
        with pytest.raises(
            ValueError, match="Image element must contain at least one markdown image"
        ):
            MarkdownImageElement.from_markdown("Invalid", "This is not an image")

    def test_text_without_image_raises(self):
        """Test that text without image markdown raises ValueError."""
        with pytest.raises(
            ValueError, match="Image element must contain at least one markdown image"
        ):
            MarkdownImageElement.from_markdown("NoImage", "# Just a title\n\nSome text here")

    def test_image_from_markdown_round_trip(self):
        """Test that ImageElement can be created from markdown and converted back."""
        original_markdown = "![Beautiful sunset](https://example.com/sunset.jpg)"
        element = MarkdownImageElement.from_markdown("Sunset", original_markdown)

        # Convert back to markdown
        result = element.to_markdown()

        # Should include comment and preserve original markdown
        expected = f"<!-- image: Sunset -->\n{original_markdown}"
        assert result == expected


class TestTableElementFromMarkdown:
    """Test TableElement.from_markdown() method."""

    def test_simple_table_from_markdown(self):
        """Test creating TableElement from simple table markdown."""
        markdown = """| Name | Age |
|------|-----|
| Alice | 25 |
| Bob   | 30 |"""

        element = MarkdownTableElement.from_markdown("People", markdown)

        assert element.name == "People"
        assert element.content_type == ContentType.TABLE
        assert isinstance(element.content, TableData)
        assert element.content.headers == ["Name", "Age"]
        assert element.content.rows == [["Alice", "25"], ["Bob", "30"]]

    def test_table_with_complex_content_from_markdown(self):
        """Test creating TableElement with complex table content."""
        markdown = """| Product | Price | In Stock | Notes |
|---------|-------|----------|-------|
| Widget  | $10.99| Yes      | Popular item |
| Gadget  | $25.00| No       | Out of stock |
| Tool    | $5.50 | Yes      | New arrival  |"""

        element = MarkdownTableElement.from_markdown("Inventory", markdown)

        assert element.name == "Inventory"
        assert element.content.headers == ["Product", "Price", "In Stock", "Notes"]
        assert len(element.content.rows) == 3
        assert element.content.rows[0] == ["Widget", "$10.99", "Yes", "Popular item"]
        assert element.content.rows[2] == ["Tool", "$5.50", "Yes", "New arrival"]

    def test_table_with_empty_cells_from_markdown(self):
        """Test creating TableElement with empty cells."""
        markdown = """| A | B | C |
|---|---|---|
| 1 |   | 3 |
|   | 2 | 3 |"""

        element = MarkdownTableElement.from_markdown("Sparse", markdown)

        assert element.content.headers == ["A", "B", "C"]
        assert element.content.rows == [["1", "", "3"], ["", "2", "3"]]

    def test_table_with_whitespace_from_markdown(self):
        """Test creating TableElement with surrounding whitespace."""
        markdown = """  
| X | Y |
|---|---|
| 1 | 2 |
  """

        element = MarkdownTableElement.from_markdown("Clean", markdown)

        assert element.content.headers == ["X", "Y"]
        assert element.content.rows == [["1", "2"]]

    def test_invalid_table_markdown_raises(self):
        """Test that invalid table markdown raises ValueError."""
        with pytest.raises(ValueError, match="Table element must contain a valid markdown table"):
            MarkdownTableElement.from_markdown("Invalid", "This is not a table")

    def test_text_without_table_raises(self):
        """Test that text without table markdown raises ValueError."""
        with pytest.raises(ValueError, match="Table element must contain a valid markdown table"):
            MarkdownTableElement.from_markdown("NoTable", "# Title\n\nJust some text")

    def test_table_from_markdown_round_trip(self):
        """Test that TableElement can be created from markdown and converted back."""
        original_markdown = """| Product | Qty |
|---------|-----|
| Apple   | 10  |
| Orange  | 5   |"""

        element = MarkdownTableElement.from_markdown("Fruits", original_markdown)

        # Convert back to markdown
        result = element.to_markdown()

        # Should include comment and properly formatted table
        assert "<!-- table: Fruits -->" in result
        assert "| Product | Qty |" in result
        assert "| Apple   | 10  |" in result


class TestChartElementFromMarkdown:
    """Test ChartElement.from_markdown() method."""

    def test_simple_chart_from_markdown(self):
        """Test creating ChartElement from JSON code block."""
        markdown = """```json
{
    "type": "bar",
    "data": [1, 2, 3, 4, 5]
}
```"""

        element = MarkdownChartElement.from_markdown("BarChart", markdown)

        assert element.name == "BarChart"
        assert element.content == markdown
        assert element.content_type == ContentType.CHART
        assert element.metadata["chart_data"]["type"] == "bar"
        assert element.metadata["chart_data"]["data"] == [1, 2, 3, 4, 5]

    def test_complex_chart_from_markdown(self):
        """Test creating ChartElement with complex JSON."""
        markdown = """```json
{
    "type": "line",
    "data": {
        "labels": ["Jan", "Feb", "Mar"],
        "datasets": [
            {
                "label": "Sales",
                "data": [100, 150, 200],
                "borderColor": "blue"
            }
        ]
    },
    "options": {
        "responsive": true,
        "scales": {
            "y": {
                "beginAtZero": true
            }
        }
    }
}
```"""

        element = MarkdownChartElement.from_markdown("SalesChart", markdown)

        assert element.name == "SalesChart"
        assert element.content_type == ContentType.CHART
        chart_data = element.metadata["chart_data"]
        assert chart_data["type"] == "line"
        assert chart_data["data"]["labels"] == ["Jan", "Feb", "Mar"]
        assert chart_data["options"]["responsive"] is True

    def test_chart_with_whitespace_from_markdown(self):
        """Test creating ChartElement with surrounding whitespace."""
        markdown = """  ```json
{"simple": "chart"}
```  """

        element = MarkdownChartElement.from_markdown("Simple", markdown)

        assert element.content_type == ContentType.CHART
        assert element.metadata["chart_data"]["simple"] == "chart"

    def test_invalid_chart_format_raises(self):
        """Test that non-JSON code block raises ValueError."""
        with pytest.raises(
            ValueError, match="Chart element must contain only a ```json code block"
        ):
            MarkdownChartElement.from_markdown("Invalid", "This is not a JSON code block")

    def test_non_json_code_block_raises(self):
        """Test that non-JSON code block raises ValueError."""
        with pytest.raises(
            ValueError, match="Chart element must contain only a ```json code block"
        ):
            MarkdownChartElement.from_markdown("Python", "```python\nprint('hello')\n```")

    def test_invalid_json_raises(self):
        """Test that invalid JSON raises ValueError."""
        with pytest.raises(ValueError, match="Chart element must contain valid JSON"):
            MarkdownChartElement.from_markdown("BadJSON", "```json\n{invalid json}\n```")

    def test_chart_from_markdown_round_trip(self):
        """Test that ChartElement can be created from markdown and converted back."""
        original_markdown = """```json
{
    "title": "My Chart",
    "values": [10, 20, 30]
}
```"""

        element = MarkdownChartElement.from_markdown("MyChart", original_markdown)

        # Convert back to markdown
        result = element.to_markdown()

        # Should include comment and preserve JSON block
        expected = f"<!-- chart: MyChart -->\n{original_markdown}"
        assert result == expected


class TestFromMarkdownIntegration:
    """Test integration between different from_markdown methods."""

    def test_all_element_types_from_markdown(self):
        """Test creating all element types using from_markdown methods."""
        # Create all element types
        text_elem = MarkdownTextElement.from_markdown("Title", "# Welcome")
        image_elem = MarkdownImageElement.from_markdown("Logo", "![Logo](logo.png)")
        table_elem = MarkdownTableElement.from_markdown("Data", "| A | B |\n|---|---|\n| 1 | 2 |")
        chart_elem = MarkdownChartElement.from_markdown("Chart", '```json\n{"data": [1, 2]}\n```')

        # Verify all have correct types
        elements = [text_elem, image_elem, table_elem, chart_elem]
        expected_types = [ContentType.TEXT, ContentType.IMAGE, ContentType.TABLE, ContentType.CHART]

        for elem, expected_type in zip(elements, expected_types):
            assert elem.content_type == expected_type
            assert hasattr(elem, "from_markdown")
            assert callable(getattr(elem.__class__, "from_markdown"))

    def test_from_markdown_methods_are_classmethods(self):
        """Test that all from_markdown methods are properly defined as classmethods."""
        classes = [
            MarkdownTextElement,
            MarkdownImageElement,
            MarkdownTableElement,
            MarkdownChartElement,
        ]

        for cls in classes:
            assert hasattr(cls, "from_markdown"), f"{cls.__name__} missing from_markdown method"
            method = getattr(cls, "from_markdown")

            # Check that it's a classmethod (bound to the class, not instance)
            assert hasattr(method, "__self__"), f"{cls.__name__}.from_markdown is not a classmethod"
            assert (
                method.__self__ is cls
            ), f"{cls.__name__}.from_markdown not bound to correct class"

    def test_from_markdown_consistent_signatures(self):
        """Test that all from_markdown methods have consistent signatures."""
        classes = [
            MarkdownTextElement,
            MarkdownImageElement,
            MarkdownTableElement,
            MarkdownChartElement,
        ]

        for cls in classes:
            method = getattr(cls, "from_markdown")

            # All should accept (name: str, markdown_content: str) -> Element
            # We can't easily inspect the exact signature, but we can test the call works
            result = None
            if cls == MarkdownTextElement:
                result = method("test", "content")
            elif cls == MarkdownImageElement:
                result = method("test", "![alt](url)")
            elif cls == MarkdownTableElement:
                result = method("test", "| A |\n|---|\n| 1 |")
            elif cls == MarkdownChartElement:
                result = method("test", '```json\n{"data": 1}\n```')

            assert result is not None
            assert isinstance(result, cls)
            assert result.name == "test"


class TestTableElementMarkdownToRequests:
    """Test TableElement.markdown_element_to_requests() method."""

    def test_simple_table_to_requests(self):
        """Test converting simple table to API requests."""
        markdown = """| Name | Age |
|------|-----|
| Alice | 25 |
| Bob   | 30 |"""

        # Create MarkdownTableElement
        markdown_elem = MarkdownTableElement.from_markdown("People", markdown)
        parent_id = "slide123"

        # Convert to requests
        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Should have at least one CreateTableRequest
        assert len(requests) > 0
        assert isinstance(requests[0], CreateTableRequest)

        # Verify table dimensions
        create_request = requests[0]
        assert create_request.rows == 3  # 2 data rows + 1 header
        assert create_request.columns == 2

        # Check that all requests have consistent object IDs
        object_ids = {
            getattr(req, "objectId", None) for req in requests if hasattr(req, "objectId")
        }
        assert len(object_ids) == 1  # All should have same objectId
        assert None not in object_ids  # All should have an objectId

    def test_table_with_custom_element_id(self):
        """Test table conversion with custom element ID."""
        markdown = """| A | B |
|---|---|
| 1 | 2 |"""

        markdown_elem = MarkdownTableElement.from_markdown("Test", markdown)
        parent_id = "slide456"
        custom_id = "my_custom_table_id"

        requests = GSlidesTableElement.markdown_element_to_requests(
            markdown_elem, parent_id, element_id=custom_id
        )

        # All requests should use the custom ID
        for request in requests:
            if hasattr(request, "objectId"):
                assert request.objectId == custom_id

    def test_table_with_formatted_content_to_requests(self):
        """Test table with markdown formatting generates style requests."""
        markdown = """| Product | Description |
|---------|-------------|
| **Widget** | *Popular* item |
| `Code` | Normal text |"""

        markdown_elem = MarkdownTableElement.from_markdown("Products", markdown)
        parent_id = "slide789"

        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Should have CreateTable + text insertion + style requests
        request_types = [type(req).__name__ for req in requests]
        assert "CreateTableRequest" in request_types
        assert "InsertTextRequest" in request_types
        # Note: UpdateTextStyleRequest may or may not be present depending on implementation

        # Verify cell locations are set correctly
        text_requests = [req for req in requests if isinstance(req, InsertTextRequest)]
        cell_locations = [req.cellLocation for req in text_requests if req.cellLocation]

        # Should have cell locations for non-empty cells
        assert len(cell_locations) > 0
        for location in cell_locations:
            assert isinstance(location, TableCellLocation)
            assert location.rowIndex >= 0
            assert location.columnIndex >= 0

    def test_table_with_empty_cells_to_requests(self):
        """Test table with empty cells doesn't generate requests for empty cells."""
        markdown = """| A | B | C |
|---|---|---|
| 1 |   | 3 |
|   | 2 |   |"""

        markdown_elem = MarkdownTableElement.from_markdown("Sparse", markdown)
        parent_id = "slide_empty"

        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Should only have requests for non-empty cells
        text_requests = [req for req in requests if isinstance(req, InsertTextRequest)]

        # We expect requests for: A, C (row 0), B (row 1), plus headers A, B, C = 6 total
        # But empty cells should not generate InsertTextRequest
        assert len(text_requests) <= 6  # At most 6 non-empty cells including headers

    def test_single_column_table_to_requests(self):
        """Test single column table conversion."""
        markdown = """| Item |
|------|
| Apple |
| Banana |"""

        markdown_elem = MarkdownTableElement.from_markdown("Items", markdown)
        parent_id = "slide_single"

        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Verify table dimensions
        create_request = requests[0]
        assert isinstance(create_request, CreateTableRequest)
        assert create_request.rows == 3  # 2 data + 1 header
        assert create_request.columns == 1

    def test_single_row_table_to_requests(self):
        """Test single row (header only) table conversion."""
        markdown = """| Col1 | Col2 | Col3 |
|------|------|------|"""

        markdown_elem = MarkdownTableElement.from_markdown("HeaderOnly", markdown)
        parent_id = "slide_header"

        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Should still create table structure
        create_request = requests[0]
        assert isinstance(create_request, CreateTableRequest)
        assert create_request.rows == 1  # Just header
        assert create_request.columns == 3

    def test_large_table_to_requests(self):
        """Test larger table conversion performance and correctness."""
        # Create a 5x5 table
        headers = ["Col1", "Col2", "Col3", "Col4", "Col5"]
        rows = [[f"R{r}C{c}" for c in range(5)] for r in range(4)]

        # Build markdown
        header_line = "| " + " | ".join(headers) + " |"
        separator_line = "|" + "".join(["------|" for _ in headers])
        data_lines = ["| " + " | ".join(row) + " |" for row in rows]
        markdown = "\n".join([header_line, separator_line] + data_lines)

        markdown_elem = MarkdownTableElement.from_markdown("LargeTable", markdown)
        parent_id = "slide_large"

        requests = GSlidesTableElement.markdown_element_to_requests(markdown_elem, parent_id)

        # Verify table dimensions
        create_request = requests[0]
        assert create_request.rows == 5  # 4 data + 1 header
        assert create_request.columns == 5

        # Should have reasonable number of requests (not exponential)
        assert len(requests) < 100  # Sanity check

    def test_requests_have_correct_structure(self):
        """Test that generated requests have correct API structure."""
        markdown = """| Test | Data |
|------|------|
| A    | B    |"""

        markdown_elem = MarkdownTableElement.from_markdown("Structure", markdown)
        parent_id = "slide_structure"
        element_id = "test_table_123"

        requests = GSlidesTableElement.markdown_element_to_requests(
            markdown_elem, parent_id, element_id=element_id
        )

        # First request should be CreateTableRequest
        create_req = requests[0]
        assert isinstance(create_req, CreateTableRequest)
        assert hasattr(create_req, "elementProperties")
        assert hasattr(create_req, "rows")
        assert hasattr(create_req, "columns")

        # Text requests should have proper structure
        for req in requests[1:]:
            if isinstance(req, InsertTextRequest):
                assert hasattr(req, "objectId")
                assert hasattr(req, "text")
                assert hasattr(req, "cellLocation")
                assert req.objectId == element_id
                assert isinstance(req.cellLocation, TableCellLocation)
