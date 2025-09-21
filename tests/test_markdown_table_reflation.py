"""
Test MarkdownTableElement reflation with model_dump and model_validate.

This tests the specific roundtrip serialization pattern requested by the user.
"""

import pytest

from gslides_api.agnostic.element import MarkdownTableElement, TableData


class TestMarkdownTableElementReflation:
    """Test that MarkdownTableElement correctly reflates from model_dump."""

    def test_simple_table_reflation(self):
        """Test that simple table element reflates correctly."""
        # Create original element
        table_data = TableData(headers=["Name", "Age"], rows=[["Alice", "25"], ["Bob", "30"]])
        original = MarkdownTableElement(name="Test Table", content=table_data)

        # Test roundtrip: model_dump -> model_validate
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        # Verify structure is preserved
        assert reloaded.name == original.name
        assert reloaded.content_type == original.content_type
        assert isinstance(reloaded.content, TableData)
        assert reloaded.content.headers == original.content.headers
        assert reloaded.content.rows == original.content.rows
        assert reloaded.metadata == original.metadata

    def test_complex_table_reflation(self):
        """Test that complex table with metadata reflates correctly."""
        table_data = TableData(
            headers=["Product", "Price", "In Stock", "Notes"],
            rows=[
                ["Widget", "$10.99", "Yes", "Popular item"],
                ["Gadget", "$25.00", "No", "Out of stock"],
                ["Tool", "$5.50", "Yes", "New arrival"],
            ],
        )
        original = MarkdownTableElement(
            name="Inventory",
            content=table_data,
            metadata={"source": "database", "updated": "2023-01-01"},
        )

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        # Verify everything is preserved
        assert reloaded.name == "Inventory"
        assert reloaded.content.headers == ["Product", "Price", "In Stock", "Notes"]
        assert len(reloaded.content.rows) == 3
        assert reloaded.content.rows[1] == ["Gadget", "$25.00", "No", "Out of stock"]
        assert reloaded.metadata == {"source": "database", "updated": "2023-01-01"}

    def test_empty_table_reflation(self):
        """Test that empty table reflates correctly."""
        table_data = TableData(headers=["A", "B"], rows=[])
        original = MarkdownTableElement(name="Empty", content=table_data)

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        assert reloaded.content.headers == ["A", "B"]
        assert reloaded.content.rows == []

    def test_single_column_table_reflation(self):
        """Test that single column table reflates correctly."""
        table_data = TableData(headers=["Item"], rows=[["Apple"], ["Banana"]])
        original = MarkdownTableElement(name="Items", content=table_data)

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        assert reloaded.content.headers == ["Item"]
        assert reloaded.content.rows == [["Apple"], ["Banana"]]

    def test_table_with_empty_cells_reflation(self):
        """Test that table with empty cells reflates correctly."""
        table_data = TableData(
            headers=["A", "B", "C"], rows=[["1", "", "3"], ["", "2", ""], ["4", "5", "6"]]
        )
        original = MarkdownTableElement(name="Sparse", content=table_data)

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        assert reloaded.content.headers == ["A", "B", "C"]
        assert reloaded.content.rows == [["1", "", "3"], ["", "2", ""], ["4", "5", "6"]]

    def test_table_from_markdown_reflation(self):
        """Test that table created from markdown reflates correctly."""
        markdown = """| Product | Qty |
|---------|-----|
| Apple   | 10  |
| Orange  | 5   |"""

        original = MarkdownTableElement.from_markdown("Fruits", markdown)

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        assert reloaded.name == "Fruits"
        assert reloaded.content.headers == ["Product", "Qty"]
        assert reloaded.content.rows == [["Apple", "10"], ["Orange", "5"]]

    def test_table_markdown_conversion_after_reflation(self):
        """Test that reflated table can still convert back to markdown correctly."""
        markdown = """| Name | Age |
|------|-----|
| Alice | 25 |
| Bob   | 30 |"""

        # Create from markdown
        original = MarkdownTableElement.from_markdown("People", markdown)

        # Reflate
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        # Convert back to markdown
        result_markdown = reloaded.to_markdown()

        # Should contain table structure
        assert "<!-- table: People -->" in result_markdown
        assert "Name" in result_markdown and "Age" in result_markdown
        assert "Alice" in result_markdown and "25" in result_markdown
        assert "Bob" in result_markdown and "30" in result_markdown

    def test_reflation_preserves_pandas_functionality(self):
        """Test that reflated table still works with pandas methods."""
        table_data = TableData(
            headers=["Name", "Age", "City"], rows=[["Alice", "25", "NYC"], ["Bob", "30", "SF"]]
        )
        original = MarkdownTableElement(name="People", content=table_data)

        # Test roundtrip
        dumped = original.model_dump(mode="json")
        reloaded = MarkdownTableElement.model_validate(dumped)

        # Try pandas conversion (will skip if pandas not available)
        try:
            import pandas as pd

            df = reloaded.to_df()
            assert list(df.columns) == ["Name", "Age", "City"]
            assert len(df) == 2
            assert df.iloc[0]["Name"] == "Alice"
            assert df.iloc[1]["City"] == "SF"
        except ImportError:
            # Skip pandas test if not available
            pass

    def test_multiple_reflation_cycles(self):
        """Test that multiple reflation cycles work correctly."""
        table_data = TableData(headers=["X", "Y"], rows=[["1", "2"], ["3", "4"]])
        original = MarkdownTableElement(name="Matrix", content=table_data)

        # First cycle
        dumped1 = original.model_dump(mode="json")
        reloaded1 = MarkdownTableElement.model_validate(dumped1)

        # Second cycle
        dumped2 = reloaded1.model_dump(mode="json")
        reloaded2 = MarkdownTableElement.model_validate(dumped2)

        # Third cycle
        dumped3 = reloaded2.model_dump(mode="json")
        reloaded3 = MarkdownTableElement.model_validate(dumped3)

        # Should be identical to original
        assert reloaded3.name == original.name
        assert reloaded3.content.headers == original.content.headers
        assert reloaded3.content.rows == original.content.rows
