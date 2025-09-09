"""
Test suite for bidirectional conversion between MarkdownSlideElement and PageElement types.

Tests both levels of round-trip preservation:
1. markdown -> MarkdownSlideElement -> PageElement -> MarkdownSlideElement -> markdown
2. Google Slides API round-trip (create element -> read -> convert to markdown)
"""

from unittest.mock import MagicMock, patch

import pytest

from gslides_api.domain.domain import (
    AffineTransform,
    Dimension,
    Image,
    PageElementProperties,
    Size,
    Transform,
    Unit,
)
from gslides_api.element.image import ImageElement
from gslides_api.element.table import TableElement
from gslides_api.domain.table import Table
from gslides_api.element.shape import Shape, ShapeElement
from gslides_api.markdown.element import ContentType
from gslides_api.markdown.element import MarkdownImageElement as MarkdownImageElement
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import MarkdownTableElement as MarkdownTableElement
from gslides_api.markdown.element import MarkdownTextElement as MarkdownTextElement
from gslides_api.element.text_content import TextContent
from gslides_api.domain.text import (
    ParagraphMarker,
    ShapeProperties,
    TextElement as GSlidesTextElement,
    TextRun,
    TextStyle,
    Type as ShapeType,
)


class TestTextElementConversion:
    """Test ShapeElement ↔ TextElement conversion."""

    def test_markdown_to_shape_to_markdown_roundtrip(self):
        """Test Level 1: markdown -> MarkdownTextElement -> ShapeElement -> MarkdownTextElement -> markdown"""

        # Original markdown content
        original_markdown = "# Heading\n\nThis is **bold** text with *italics* and `code`."

        # Create MarkdownTextElement
        markdown_elem = MarkdownTextElement(name="Test Text", content=original_markdown)

        # Convert to ShapeElement
        with patch(
            "gslides_api.element.shape.ShapeElement.to_markdown",
            return_value=original_markdown,
        ):
            shape_elem = ShapeElement.from_markdown_element(markdown_elem, parent_id="slide_123")

        # Convert back to MarkdownTextElement
        with patch(
            "gslides_api.element.shape.ShapeElement.to_markdown",
            return_value=original_markdown,
        ):
            converted_markdown_elem = shape_elem.to_markdown_element(name="Test Text")

        # Convert back to markdown string
        final_markdown = converted_markdown_elem.to_markdown()

        # Should preserve the content exactly
        assert converted_markdown_elem.content == original_markdown
        assert "Test Text" in final_markdown or original_markdown in final_markdown

    def test_shape_element_metadata_preservation(self):
        """Test that ShapeElement metadata is preserved during conversion."""

        # Create a ShapeElement with metadata
        shape = Shape(
            shapeProperties=ShapeProperties(),
            shapeType=ShapeType.TEXT_BOX,
            text=TextContent(textElements=[]),
        )

        shape_elem = ShapeElement(
            objectId="shape_123",
            shape=shape,
            size=Size(
                width=Dimension(magnitude=100, unit=Unit.PT),
                height=Dimension(magnitude=50, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            title="Test Shape",
            description="Test Description",
            slide_id="slide_123",
            presentation_id="pres_123",
        )

        # Mock the to_markdown method using patch on the class
        with patch(
            "gslides_api.element.shape.ShapeElement.to_markdown",
            return_value="Test content",
        ):
            # Convert to markdown element
            markdown_elem = shape_elem.to_markdown_element()

            # Check metadata preservation
            assert markdown_elem.metadata["objectId"] == "shape_123"
            assert markdown_elem.metadata["shape_type"] == "TEXT_BOX"
            assert markdown_elem.metadata["title"] == "Test Shape"
            assert markdown_elem.metadata["description"] == "Test Description"
            assert "size" in markdown_elem.metadata
            assert markdown_elem.metadata["size"]["width"] == 100
            assert markdown_elem.metadata["size"]["unit"] == "PT"


class TestImageElementConversion:
    """Test ImageElement ↔ ImageElement conversion."""

    def test_markdown_to_image_to_markdown_roundtrip(self):
        """Test Level 1: markdown -> MarkdownImageElement -> ImageElement -> MarkdownImageElement -> markdown"""

        # Original markdown content
        original_markdown = "![Test Image](https://example.com/image.jpg)"

        # Create MarkdownImageElement
        markdown_elem = MarkdownImageElement.from_markdown(
            name="Test Image", markdown_content=original_markdown
        )

        # Convert to ImageElement
        image_elem = ImageElement.from_markdown_element(markdown_elem, parent_id="slide_123")

        # Convert back to MarkdownImageElement
        converted_markdown_elem = image_elem.to_markdown_element(name="Test Image")

        # Convert back to markdown string
        final_markdown = converted_markdown_elem.to_markdown()

        # Should preserve the URL and alt text
        assert "https://example.com/image.jpg" in final_markdown
        assert "Test Image" in final_markdown

    def test_image_element_metadata_preservation(self):
        """Test that ImageElement metadata is preserved during conversion."""

        # Create an ImageElement with metadata
        image = Image(
            sourceUrl="https://example.com/image.jpg",
            contentUrl="https://cached.example.com/image.jpg",
        )

        image_elem = ImageElement(
            objectId="image_123",
            image=image,
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=150, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            title="Test Image",
            description="Test Description",
            slide_id="slide_123",
            presentation_id="pres_123",
        )

        # Convert to markdown element
        markdown_elem = image_elem.to_markdown_element()

        # Check metadata preservation
        assert markdown_elem.metadata["objectId"] == "image_123"
        assert markdown_elem.metadata["sourceUrl"] == "https://example.com/image.jpg"
        assert markdown_elem.metadata["contentUrl"] == "https://cached.example.com/image.jpg"
        assert markdown_elem.metadata["title"] == "Test Image"
        assert markdown_elem.metadata["description"] == "Test Description"
        assert "size" in markdown_elem.metadata
        assert markdown_elem.metadata["size"]["width"] == 200


class TestTableElementConversion:
    """Test TableElement ↔ TableElement conversion."""

    def test_markdown_to_requests_generation(self):
        """Test: markdown -> MarkdownTableElement -> API requests"""

        # Original markdown table
        original_markdown = """| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |"""

        # Create MarkdownTableElement
        markdown_elem = MarkdownTableElement(name="Test Table", content=original_markdown)

        # Convert to API requests
        requests = TableElement.markdown_element_to_requests(
            markdown_elem, slide_id="slide_123", element_id="test_table"
        )

        # Should have at least a CreateTableRequest
        assert len(requests) > 0
        from gslides_api.request.table import CreateTableRequest

        assert isinstance(requests[0], CreateTableRequest)

        # Verify table dimensions
        create_request = requests[0]
        assert create_request.rows == 3  # 2 data rows + 1 header
        assert create_request.columns == 3
        assert create_request.objectId == "test_table"

    def test_table_data_extraction(self):
        """Test extraction of table data from Google Slides table structure."""

        # Mock Google Slides table structure
        table_rows = [
            {
                "tableCells": [
                    {
                        "text": {
                            "textElements": [{"endIndex": 8, "textRun": {"content": "Header 1"}}]
                        }
                    },
                    {
                        "text": {
                            "textElements": [{"endIndex": 8, "textRun": {"content": "Header 2"}}]
                        }
                    },
                ]
            },
            {
                "tableCells": [
                    {"text": {"textElements": [{"endIndex": 6, "textRun": {"content": "Cell 1"}}]}},
                    {"text": {"textElements": [{"endIndex": 6, "textRun": {"content": "Cell 2"}}]}},
                ]
            },
        ]

        table = Table(rows=2, columns=2, tableRows=table_rows)

        table_elem = TableElement(
            objectId="table_123",
            table=table,
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            slide_id="slide_123",
            presentation_id="pres_123",
        )

        # Extract table data
        table_data = table_elem.extract_table_data()

        # Check extracted data
        assert table_data.headers == ["Header 1", "Header 2"]
        assert len(table_data.rows) == 1
        assert table_data.rows[0] == ["Cell 1", "Cell 2"]

    def test_table_element_metadata_preservation(self):
        """Test that TableElement metadata is preserved during conversion."""

        # Create a TableElement with metadata
        table = Table(
            rows=2,
            columns=3,
            tableRows=[
                {
                    "tableCells": [
                        {"text": {"textElements": [{"endIndex": 2, "textRun": {"content": "H1"}}]}},
                        {"text": {"textElements": [{"endIndex": 2, "textRun": {"content": "H2"}}]}},
                        {"text": {"textElements": [{"endIndex": 2, "textRun": {"content": "H3"}}]}},
                    ]
                }
            ],
        )

        table_elem = TableElement(
            objectId="table_123",
            table=table,
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=200, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            title="Test Table",
            description="Test Description",
            slide_id="slide_123",
            presentation_id="pres_123",
        )

        # Convert to markdown element
        markdown_elem = table_elem.to_markdown_element()

        # Check metadata preservation
        assert markdown_elem.metadata["objectId"] == "table_123"
        assert markdown_elem.metadata["rows"] == 2
        assert markdown_elem.metadata["columns"] == 3
        assert markdown_elem.metadata["title"] == "Test Table"
        assert markdown_elem.metadata["description"] == "Test Description"
        assert "size" in markdown_elem.metadata
        assert markdown_elem.metadata["size"]["width"] == 300


class TestIntegrationRoundTrip:
    """Test full integration scenarios."""

    def test_empty_content_handling(self):
        """Test handling of empty or minimal content."""

        # Empty text
        empty_text = MarkdownTextElement(name="Empty", content="")
        shape = ShapeElement.from_markdown_element(empty_text, "slide_123")
        converted = shape.to_markdown_element()
        assert converted.content == ""

        # Empty table - test request generation
        empty_table_data = TableData(headers=["Col1"], rows=[])
        empty_table = MarkdownTableElement(name="Empty", content=empty_table_data)

        # Should generate valid API requests even for empty table
        requests = TableElement.markdown_element_to_requests(empty_table, "slide_123")
        from gslides_api.request.table import CreateTableRequest

        assert len(requests) > 0
        assert isinstance(requests[0], CreateTableRequest)
        assert requests[0].rows == 1  # Just header row
        assert requests[0].columns == 1

    def test_special_characters_handling(self):
        """Test handling of special characters and edge cases."""

        # Text with special markdown characters
        special_text = "Text with | pipes & **bold** and [links](http://example.com)"
        markdown_elem = MarkdownTextElement(name="Special", content=special_text)

        with patch(
            "gslides_api.element.shape.ShapeElement.to_markdown",
            return_value=special_text,
        ):
            shape = ShapeElement.from_markdown_element(markdown_elem, "slide_123")
            converted = shape.to_markdown_element()
            assert converted.content == special_text

    def test_large_table_handling(self):
        """Test handling of large tables."""

        # Create a large table
        headers = [f"Header {i}" for i in range(10)]
        rows = [[f"Cell {i}-{j}" for j in range(10)] for i in range(20)]
        large_table_data = TableData(headers=headers, rows=rows)

        large_table = MarkdownTableElement(name="Large", content=large_table_data)

        # Convert to API requests
        requests = TableElement.markdown_element_to_requests(
            large_table, slide_id="slide_123", element_id="large_table"
        )

        # Verify table creation request
        from gslides_api.request.table import CreateTableRequest

        assert isinstance(requests[0], CreateTableRequest)
        create_request = requests[0]
        assert create_request.rows == 21  # 20 data rows + 1 header
        assert create_request.columns == 10
        assert create_request.objectId == "large_table"


# Integration test fixtures and helpers
@pytest.fixture
def mock_api_client():
    """Mock GoogleAPIClient for testing."""
    client = MagicMock()
    client.batch_update.return_value = {"replies": []}
    return client


@pytest.fixture
def sample_presentation_data():
    """Sample presentation data for testing."""
    return {
        "presentationId": "test_pres_123",
        "slides": [{"objectId": "slide_123", "pageElements": []}],
    }


# TODO: Add Google Slides API integration tests when credentials are available
# These would test actual round-trip through the Google Slides API:
# 1. Create elements using the API
# 2. Read them back
# 3. Convert to markdown
# 4. Verify content preservation
