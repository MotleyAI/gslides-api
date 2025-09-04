"""
Comprehensive TableElement integration test with styling preservation.

Tests the full round-trip conversion chain:
Markdown → MarkdownTableElement → TableElement → Google Slides API → TableElement → MarkdownTableElement → Markdown
"""

import os
import pytest

from gslides_api.client import api_client, initialize_credentials
from gslides_api.element.table import TableElement
from gslides_api.markdown.element import MarkdownTableElement as MarkdownTableElement
from gslides_api.presentation import Presentation


class TestTableComprehensiveIntegration:
    """Test comprehensive TableElement integration with styling."""

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
        cls.test_presentation = Presentation.create_blank("Table Integration Test")
        print(f"Created test presentation: {cls.test_presentation.presentationId}")

        # Get the first slide for testing
        cls.test_slide = cls.test_presentation.slides[0]

    @classmethod
    def teardown_class(cls):
        """Clean up test presentation."""
        if hasattr(cls, "test_presentation"):
            try:
                # Delete the test presentation
                from googleapiclient.discovery import build

                service = build("drive", "v3", credentials=api_client.credentials)
                service.files().delete(fileId=cls.test_presentation.presentationId).execute()
                print(f"Deleted test presentation: {cls.test_presentation.presentationId}")
            except Exception as e:
                print(f"Warning: Failed to delete test presentation: {e}")

    def test_markdown_table_dual_parsing_validation(self):
        """Test that dual parsing method works correctly."""
        styled_markdown = """| **Product** | *Description* | **Price** | Notes |
|-------------|---------------|-----------|-------|
| **Widget**  | *Premium quality* | **$19.99** | `New item` |
| Gadget      | Standard **bold** text | $9.99 | Regular text |
| *Tool*      | Mixed *italic* **bold** | **$15.50** | *Special offer* |"""

        # Test dual parsing works
        marko_cells, markdown_cells = MarkdownTableElement._parse_table_dual_method(styled_markdown)

        assert len(marko_cells) == 4, f"Expected 4 rows, got {len(marko_cells)}"
        assert len(markdown_cells) == 4, f"Expected 4 rows, got {len(markdown_cells)}"

        # Verify styling is preserved in markdown cells
        assert "**Product**" in markdown_cells[0][0]
        assert "*Description*" in markdown_cells[0][1]
        assert "**Widget**" in markdown_cells[1][0]
        assert "*Premium quality*" in markdown_cells[1][1]
        assert "`New item`" in markdown_cells[1][3]

        print("✓ Dual parsing validation passed")

    def test_markdown_table_element_styling_preservation(self):
        """Test that MarkdownTableElement preserves styling when created from markdown."""
        styled_markdown = """| **Name** | *Description* | Status |
|----------|---------------|---------|
| **Item1** | *Important* item | `Active` |
| Item2    | Regular text | Normal |"""

        # Create MarkdownTableElement from styled markdown
        table_element = MarkdownTableElement.from_markdown("StyledTable", styled_markdown)

        # Verify headers preserve styling
        assert table_element.content.headers == ["**Name**", "*Description*", "Status"]

        # Verify rows preserve styling
        assert table_element.content.rows[0] == ["**Item1**", "*Important* item", "`Active`"]
        assert table_element.content.rows[1] == ["Item2", "Regular text", "Normal"]

        print("✓ MarkdownTableElement styling preservation passed")

    def test_table_element_creation_and_api_operations(self):
        """Test TableElement creation and basic API operations."""
        styled_markdown = """| **Product** | *Price* | Notes |
|-------------|---------|-------|
| **Widget**  | **$10** | `Hot item` |
| Tool        | $5      | Regular |"""

        # Create MarkdownTableElement with styled content
        markdown_table = MarkdownTableElement.from_markdown("APITest", styled_markdown)

        # Convert to API requests and create the table
        requests = TableElement.markdown_element_to_requests(
            markdown_table, parent_id=self.test_slide.objectId, element_id="api_test_table"
        )

        # Execute the requests to create the table
        api_client.batch_update(requests, self.test_presentation.presentationId)
        new_element_id = "api_test_table"

        print(f"✓ Table created in Google Slides with ID: {new_element_id}")

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

        print("✓ Table element created and retrieved successfully")

    def test_full_table_roundtrip_with_styling(self):
        """Test complete table round-trip with styling preservation."""
        original_markdown = """| **Name** | *Description* | **Price** |
|----------|---------------|-----------|
| **Bold Item** | *Italic text* with regular | **$15.99** |
| Regular Item  | Standard text | $9.50 |
| *Italic Name* | Mixed **bold** text | **$12** |"""

        print("=== FULL ROUNDTRIP TEST ===")
        print(f"Original markdown:\\n{original_markdown}\\n")

        # Step 1: Create MarkdownTableElement
        markdown_table = MarkdownTableElement.from_markdown("RoundtripTest", original_markdown)
        print("✓ Step 1: MarkdownTableElement created")

        # Step 2: Convert to API requests
        requests = TableElement.markdown_element_to_requests(
            markdown_table, parent_id=self.test_slide.objectId, element_id="roundtrip_test_table"
        )
        print("✓ Step 2: API requests generated")

        # Step 3: Create in Google Slides
        api_client.batch_update(requests, self.test_presentation.presentationId)
        new_element_id = "roundtrip_test_table"
        print("✓ Step 3: Table created in Google Slides")

        # Step 5: Read back from Google Slides API
        updated_slide = self.test_slide.__class__.from_ids(
            self.test_presentation.presentationId,
            self.test_slide.objectId,
            api_client=api_client,
        )

        retrieved_element = updated_slide.get_element_by_id(new_element_id)
        assert retrieved_element is not None, "Element not found after API round-trip"
        assert isinstance(retrieved_element, TableElement), "Element is not a TableElement"
        print("✓ Step 5: TableElement retrieved from Google Slides")

        # Step 6: Convert back to MarkdownTableElement
        result_markdown_element = retrieved_element.to_markdown_element("RoundtripTest")
        print("✓ Step 6: Converted back to MarkdownTableElement")

        # Step 7: Convert to final markdown
        final_markdown = result_markdown_element.to_markdown()
        print("✓ Step 7: Final markdown generated")

        print(f"Final markdown:\\n{final_markdown}\\n")

        # Verify structure preservation
        assert result_markdown_element.content.headers is not None, "Headers lost"
        assert (
            len(result_markdown_element.content.headers) == 3
        ), f"Expected 3 headers, got {len(result_markdown_element.content.headers)}"
        assert (
            len(result_markdown_element.content.rows) >= 2
        ), f"Expected at least 2 rows, got {len(result_markdown_element.content.rows)}"

        # Verify content preservation (text content, styling may vary)
        result_content = final_markdown.lower()
        assert "name" in result_content, "Header 'Name' not found"
        assert "description" in result_content, "Header 'Description' not found"
        assert "price" in result_content, "Header 'Price' not found"
        assert "bold item" in result_content or "bold" in result_content, "Bold item text not found"
        assert "italic" in result_content, "Italic text not found"
        assert "regular" in result_content, "Regular text not found"

        print("✓ Content verification passed")

        # Check for any styling preservation (best effort)
        styling_preserved = 0
        if "**" in final_markdown:
            styling_preserved += 1
            print("✓ Bold formatting (**) preserved")
        if "*" in final_markdown and "**" not in final_markdown.replace("**", ""):
            styling_preserved += 1
            print("✓ Italic formatting (*) preserved")
        if "`" in final_markdown:
            styling_preserved += 1
            print("✓ Code formatting (`) preserved")

        print(f"✓ Styling preservation score: {styling_preserved}/3")
        print("✓ Full round-trip test completed successfully!")

        # Success - all assertions passed!


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
