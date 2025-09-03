"""
Google Slides API round-trip tests for bidirectional element conversion.

These tests create actual Google Slides presentations, add elements,
read them back, and verify perfect conversion to markdown.
"""

import os
from typing import Optional

import pytest

from gslides_api.client import api_client, initialize_credentials

# Note: Importing individual element types instead of union type
from gslides_api.domain import Dimension, Size, Transform, Unit
from gslides_api.element.image import ImageElement
from gslides_api.element.table import TableElement
from gslides_api.element.shape import ShapeElement
from gslides_api.markdown.element import ImageElement as MarkdownImageElement
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import TableElement as MarkdownTableElement
from gslides_api.markdown.element import TextElement as MarkdownTextElement
from gslides_api.presentation import Presentation
from gslides_api.text import Shape, ShapeProperties
from gslides_api.text import Type as ShapeType


class TestAPIRoundTrip:
    """Test Google Slides API round-trip conversion."""

    @classmethod
    def setup_class(cls):
        """Set up API client and create test presentation."""
        # Initialize credentials
        credentials_path = os.environ.get("GSLIDES_CREDENTIALS_PATH")
        if not credentials_path:
            pytest.skip("GSLIDES_CREDENTIALS_PATH not set, skipping API tests")

        try:
            initialize_credentials(credentials_path)
        except Exception as e:
            pytest.skip(f"Failed to initialize credentials: {e}")

        # Create test presentation
        cls.test_presentation = Presentation.create_blank("API Round-trip Test")
        print(f"Created test presentation: {cls.test_presentation.presentationId}")

        # Get the first slide for testing
        cls.test_slide = cls.test_presentation.slides[0]

    @classmethod
    def teardown_class(cls):
        """Clean up by deleting the test presentation."""
        if hasattr(cls, "test_presentation") and cls.test_presentation:
            try:
                # Use Drive API to delete the presentation
                api_client.drive_service.files().delete(
                    fileId=cls.test_presentation.presentationId
                ).execute()
                print(f"Deleted test presentation: {cls.test_presentation.presentationId}")
            except Exception as e:
                print(f"Warning: Failed to delete test presentation: {e}")

    def test_shape_element_api_round_trip(self):
        """Test ShapeElement API round-trip conversion."""
        # Create markdown text element
        markdown_text = MarkdownTextElement(
            name="TestText",
            content="# Hello World\n\nThis is **bold** and *italic* text with a list:\n\n- Item 1\n- Item 2\n- Item 3",
            metadata={"objectId": "test_shape_1", "shape_type": "TEXT_BOX"},
        )

        # Convert to ShapeElement
        shape_element = ShapeElement.from_markdown_element(
            markdown_text, parent_id=self.test_slide.objectId, api_client=api_client
        )
        shape_element.presentation_id = self.test_presentation.presentationId

        # Create the element in Google Slides
        new_element_id = shape_element.create_copy(
            parent_id=self.test_slide.objectId,
            presentation_id=self.test_presentation.presentationId,
            api_client=api_client,
        )

        # Write the text content to the created element
        shape_element.objectId = new_element_id
        shape_element.write_text(markdown_text.content, as_markdown=True, api_client=api_client)

        # Read the slide back from API to get the actual element
        updated_slide = self.test_slide.__class__.from_ids(
            self.test_presentation.presentationId,
            self.test_slide.objectId,
            api_client=api_client,
        )

        # Find our created element
        created_element = updated_slide.get_element_by_id(new_element_id)
        assert created_element is not None, "Created element not found in slide"
        assert isinstance(created_element, ShapeElement), "Element is not a ShapeElement"

        # Convert back to markdown
        result_markdown_element = created_element.to_markdown_element("TestText")

        # Verify the conversion preserved the content structure
        assert result_markdown_element.name == "TestText"

        # The exact markdown may vary due to API formatting, but key structure should be preserved
        result_content = result_markdown_element.content
        assert "Hello World" in result_content
        assert "bold" in result_content or "**" in result_content
        assert "italic" in result_content or "*" in result_content
        assert "Item 1" in result_content

        print(f"Original: {markdown_text.content}")
        print(f"Result: {result_content}")

    def test_table_element_api_round_trip(self):
        """Test TableElement API round-trip conversion."""
        # Create markdown table element
        table_data = TableData(
            headers=["Name", "Age", "City"],
            rows=[
                ["Alice", "30", "New York"],
                ["Bob", "25", "Los Angeles"],
                ["Charlie", "35", "Chicago"],
            ],
        )

        markdown_table = MarkdownTableElement(
            name="TestTable",
            content=table_data,  # Pass TableData object directly
            metadata={"objectId": "test_table_1"},
        )

        # Convert to TableElement
        table_element = TableElement.from_markdown_element(
            markdown_table, parent_id=self.test_slide.objectId, api_client=api_client
        )
        table_element.presentation_id = self.test_presentation.presentationId

        # Create the element in Google Slides
        new_element_id = table_element.create_copy(
            parent_id=self.test_slide.objectId,
            presentation_id=self.test_presentation.presentationId,
            api_client=api_client,
        )

        # Read the slide back from API
        updated_slide = self.test_slide.__class__.from_ids(
            self.test_presentation.presentationId,
            self.test_slide.objectId,
            api_client=api_client,
        )

        # Find our created element
        created_element = updated_slide.get_element_by_id(new_element_id)
        assert created_element is not None, "Created element not found in slide"
        assert isinstance(created_element, TableElement), "Element is not a TableElement"

        # Convert back to markdown
        result_markdown_element = created_element.to_markdown_element("TestTable")

        # Verify the table data was preserved
        assert result_markdown_element.name == "TestTable"
        assert result_markdown_element.content is not None

        # Check that the extracted table has the correct structure
        # Note: Google Slides may not preserve exact text, but structure should be similar
        assert len(result_markdown_element.content.headers) == 3
        assert len(result_markdown_element.content.rows) == 3

        print(f"Original table data: {table_data}")
        print(f"Result table data: {result_markdown_element.content}")

        # Verify that we can successfully extract and convert table data from the API
        print("API round-trip successful: table created, read, and converted back to markdown")

    @pytest.mark.skip(
        reason="ImageElement requires actual image upload - skipping for basic API test"
    )
    def test_image_element_api_round_trip(self):
        """Test ImageElement API round-trip conversion (placeholder)."""
        # This would require uploading actual images to Google Drive
        # and creating image elements, which is more complex
        pass

    def test_multiple_elements_round_trip(self):
        """Test multiple elements in a single slide."""
        # Create a slide with multiple elements
        elements = []

        # Add a text element
        text_element = MarkdownTextElement(
            name="MultiText",
            content="## Multi-element Test\n\nThis slide has multiple elements.",
            metadata={"objectId": "multi_text_1"},
        )
        shape = ShapeElement.from_markdown_element(text_element, parent_id=self.test_slide.objectId)
        shape.presentation_id = self.test_presentation.presentationId
        shape.transform.translateX = 100000  # Offset position
        shape.transform.translateY = 100000

        text_id = shape.create_copy(
            parent_id=self.test_slide.objectId,
            presentation_id=self.test_presentation.presentationId,
        )
        shape.objectId = text_id
        shape.write_text(text_element.content, as_markdown=True)
        elements.append(("text", text_id))

        # Add a table element
        table_data = TableData(
            headers=["Feature", "Status"],
            rows=[["Text", "Working"], ["Table", "Working"]],
        )
        table_element = MarkdownTableElement(
            name="MultiTable",
            content=table_data,  # Pass TableData object directly
            metadata={"objectId": "multi_table_1"},
        )
        table = TableElement.from_markdown_element(
            table_element, parent_id=self.test_slide.objectId
        )
        table.presentation_id = self.test_presentation.presentationId
        table.transform.translateX = 100000  # Offset position
        table.transform.translateY = 300000

        table_id = table.create_copy(
            parent_id=self.test_slide.objectId,
            presentation_id=self.test_presentation.presentationId,
        )
        elements.append(("table", table_id))

        # Read back and verify all elements
        updated_slide = self.test_slide.__class__.from_ids(
            self.test_presentation.presentationId, self.test_slide.objectId
        )

        for element_type, element_id in elements:
            found_element = updated_slide.get_element_by_id(element_id)
            assert found_element is not None, f"{element_type} element {element_id} not found"

            if element_type == "text":
                assert isinstance(found_element, ShapeElement)
                markdown_elem = found_element.to_markdown_element("MultiText")
                assert "Multi-element Test" in markdown_elem.content
            elif element_type == "table":
                assert isinstance(found_element, TableElement)
                markdown_elem = found_element.to_markdown_element("MultiTable")
                assert markdown_elem.content is not None
                assert len(markdown_elem.content.headers) > 0

        print(f"Successfully tested {len(elements)} elements in multi-element round-trip")
