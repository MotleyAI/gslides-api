"""
Test __getitem__ and __setitem__ methods for MarkdownTableElement.

Tests the new indexing functionality with headers as row 0 and Marko validation.
"""

import pytest

from gslides_api.markdown.element import MarkdownTableElement, TableData


class TestMarkdownTableElementGetSet:
    """Test __getitem__ and __setitem__ methods for MarkdownTableElement."""

    def test_getitem_direct_access(self):
        """Test direct cell access with [row, col] syntax."""
        table_data = TableData(
            headers=['Name', 'Age', 'City'],
            rows=[['Alice', '25', 'NYC'], ['Bob', '30', 'SF']]
        )
        table = MarkdownTableElement(name='People', content=table_data)

        # Test header access (row 0)
        assert table[0, 0] == 'Name'
        assert table[0, 1] == 'Age'
        assert table[0, 2] == 'City'

        # Test data row access (row 1+)
        assert table[1, 0] == 'Alice'
        assert table[1, 1] == '25'
        assert table[1, 2] == 'NYC'
        assert table[2, 0] == 'Bob'
        assert table[2, 1] == '30'
        assert table[2, 2] == 'SF'

    def test_getitem_chaining_access(self):
        """Test chained access with table[row][col] syntax."""
        table_data = TableData(
            headers=['Product', 'Price'],
            rows=[['Widget', '$10'], ['Gadget', '$20']]
        )
        table = MarkdownTableElement(name='Products', content=table_data)

        # Test header access via chaining
        assert table[0][0] == 'Product'
        assert table[0][1] == 'Price'

        # Test data row access via chaining
        assert table[1][0] == 'Widget'
        assert table[1][1] == '$10'
        assert table[2][0] == 'Gadget'
        assert table[2][1] == '$20'

    def test_getitem_negative_indices(self):
        """Test negative indexing support."""
        table_data = TableData(
            headers=['A', 'B'],
            rows=[['1', '2'], ['3', '4']]
        )
        table = MarkdownTableElement(name='Matrix', content=table_data)

        # Total rows: 1 header + 2 data = 3 rows (indices 0, 1, 2)
        # table[-1] should be row 2 (last row), table[-2] should be row 1, table[-3] should be row 0

        # Test negative row indexing
        assert table[-1, 0] == '3'  # Last row (row 2), first col
        assert table[-1, 1] == '4'  # Last row (row 2), second col  
        assert table[-2, 0] == '1'  # Second to last row (row 1), first col

        # Test negative column indexing  
        assert table[0, -1] == 'B'  # Header row, last col
        assert table[1, -1] == '2'  # First data row, last col

    def test_setitem_direct_assignment(self):
        """Test direct cell assignment with [row, col] = value syntax."""
        table_data = TableData(
            headers=['Name', 'Score'],
            rows=[['Alice', '85'], ['Bob', '92']]
        )
        table = MarkdownTableElement(name='Scores', content=table_data)

        # Test header modification
        table[0, 0] = 'Student'
        assert table[0, 0] == 'Student'
        assert table.content.headers[0] == 'Student'

        # Test data cell modification
        table[1, 1] = '88'
        assert table[1, 1] == '88'
        assert table.content.rows[0][1] == '88'

        # Test that other cells remain unchanged
        assert table[0, 1] == 'Score'
        assert table[1, 0] == 'Alice'
        assert table[2, 0] == 'Bob'

    def test_setitem_chaining_assignment(self):
        """Test chained assignment with table[row][col] = value syntax."""
        table_data = TableData(
            headers=['X', 'Y'],
            rows=[['1', '2'], ['3', '4']]
        )
        table = MarkdownTableElement(name='Points', content=table_data)

        # Test header modification via chaining
        table[0][0] = 'Longitude'
        assert table[0, 0] == 'Longitude'

        # Test data cell modification via chaining
        table[1][1] = '2.5'
        assert table[1, 1] == '2.5'

    def test_setitem_with_markdown_formatting(self):
        """Test that markdown formatting in cell values is preserved."""
        table_data = TableData(
            headers=['Item', 'Description'],
            rows=[['Apple', 'Red fruit'], ['Banana', 'Yellow fruit']]
        )
        table = MarkdownTableElement(name='Fruits', content=table_data)

        # Set cell with markdown formatting
        table[1, 1] = '**Bold** and *italic* text with `code`'
        assert table[1, 1] == '**Bold** and *italic* text with `code`'
        
        # Verify the table can still be converted to valid markdown
        markdown_output = table.to_markdown()
        assert '**Bold** and *italic* text with `code`' in markdown_output

    def test_setitem_marko_validation_success(self):
        """Test that valid changes pass Marko validation."""
        table_data = TableData(
            headers=['Name', 'Value'],
            rows=[['A', '1'], ['B', '2']]
        )
        table = MarkdownTableElement(name='Data', content=table_data)

        # These should all succeed and be validated
        table[0, 0] = 'New Header'
        table[1, 0] = 'New Name'
        table[1, 1] = 'New Value'
        table[2, 1] = '42'

        # Verify changes were applied
        assert table.content.headers[0] == 'New Header'
        assert table.content.rows[0][0] == 'New Name'
        assert table.content.rows[0][1] == 'New Value'
        assert table.content.rows[1][1] == '42'

    def test_setitem_preserves_table_structure(self):
        """Test that setting cells preserves overall table structure."""
        table_data = TableData(
            headers=['Col1', 'Col2', 'Col3'],
            rows=[['A', 'B', 'C'], ['D', 'E', 'F']]
        )
        table = MarkdownTableElement(name='Grid', content=table_data)

        # Modify various cells
        table[0, 1] = 'New Header'
        table[1, 0] = 'Changed'
        table[2, 2] = 'Modified'

        # Verify table structure is intact
        assert len(table.content.headers) == 3
        assert len(table.content.rows) == 2
        assert len(table.content.rows[0]) == 3
        assert len(table.content.rows[1]) == 3

        # Verify specific changes
        assert table[0, 1] == 'New Header'
        assert table[1, 0] == 'Changed'
        assert table[2, 2] == 'Modified'

        # Verify unchanged cells
        assert table[0, 0] == 'Col1'
        assert table[0, 2] == 'Col3'
        assert table[1, 1] == 'B'

    def test_getitem_index_errors(self):
        """Test that invalid indices raise appropriate errors."""
        table_data = TableData(
            headers=['A', 'B'],
            rows=[['1', '2']]
        )
        table = MarkdownTableElement(name='Small', content=table_data)

        # Test row index errors
        with pytest.raises(IndexError, match="Row index 3 out of range"):
            _ = table[3, 0]
        
        with pytest.raises(IndexError, match="Row index -3 out of range"):
            _ = table[-5, 0]

        # Test column index errors
        with pytest.raises(IndexError, match="Column index 3 out of range"):
            _ = table[0, 3]
            
        with pytest.raises(IndexError, match="Column index -3 out of range"):
            _ = table[0, -5]

    def test_setitem_index_errors(self):
        """Test that invalid indices in setitem raise appropriate errors."""
        table_data = TableData(
            headers=['A'],
            rows=[['1']]
        )
        table = MarkdownTableElement(name='Tiny', content=table_data)

        # Test row index errors
        with pytest.raises(IndexError, match="Row index 5 out of range"):
            table[5, 0] = 'value'

        # Test column index errors  
        with pytest.raises(IndexError, match="Column index 5 out of range"):
            table[0, 5] = 'value'

    def test_setitem_type_errors(self):
        """Test that invalid value types raise appropriate errors."""
        table_data = TableData(
            headers=['Name'],
            rows=[['Alice']]
        )
        table = MarkdownTableElement(name='Names', content=table_data)

        # Test non-string value
        with pytest.raises(TypeError, match="Cell value must be a string"):
            table[0, 0] = 123

        with pytest.raises(TypeError, match="Cell value must be a string"):
            table[1, 0] = None

    def test_getitem_type_errors(self):
        """Test that invalid key types raise appropriate errors."""
        table_data = TableData(
            headers=['A'],
            rows=[['1']]
        )
        table = MarkdownTableElement(name='Test', content=table_data)

        # Test invalid key types for __getitem__
        with pytest.raises(TypeError, match="Table indexing requires either"):
            _ = table['invalid']

        with pytest.raises(TypeError, match="Table indexing requires either"):
            _ = table[(1, 2, 3)]  # Too many indices

    def test_setitem_type_errors_keys(self):
        """Test that invalid key types in setitem raise appropriate errors."""
        table_data = TableData(
            headers=['A'],
            rows=[['1']]
        )
        table = MarkdownTableElement(name='Test', content=table_data)

        # Test invalid key types for __setitem__
        with pytest.raises(TypeError, match="Table assignment requires"):
            table['invalid'] = 'value'

        with pytest.raises(TypeError, match="Table assignment requires"):
            table[1] = 'value'  # Row assignment not supported directly

    def test_empty_table_errors(self):
        """Test behavior with empty tables.""" 
        empty_table = MarkdownTableElement(
            name='Empty',
            content=TableData(headers=[], rows=[])
        )

        # Test access to empty table
        with pytest.raises(IndexError, match="Table is empty"):
            _ = empty_table[0, 0]

        # Test assignment to empty table
        with pytest.raises(IndexError, match="Cannot set cell in empty table"):
            empty_table[0, 0] = 'value'

    def test_integration_with_existing_methods(self):
        """Test that new methods work well with existing MarkdownTableElement functionality."""
        # Create table from markdown
        markdown = """| Name | Age | City |
|------|-----|------|
| Alice | 25 | NYC |
| Bob   | 30 | SF  |"""

        table = MarkdownTableElement.from_markdown("People", markdown)

        # Test __getitem__ works
        assert table[0, 0].strip() == 'Name'
        assert table[1, 0].strip() == 'Alice'

        # Test __setitem__ works
        table[1, 2] = 'Boston'
        assert table[1, 2] == 'Boston'

        # Test conversion back to markdown still works
        result_markdown = table.to_markdown()
        assert 'Boston' in result_markdown
        assert 'People' in result_markdown

    def test_round_trip_with_indexing(self):
        """Test that tables modified via indexing can round-trip correctly."""
        original_data = TableData(
            headers=['Product', 'Qty', 'Price'],
            rows=[['Apple', '10', '$1.00'], ['Orange', '5', '$1.50']]
        )
        table = MarkdownTableElement(name='Inventory', content=original_data)

        # Modify via indexing
        table[0, 0] = 'Item'  # Change header
        table[1, 1] = '15'    # Change quantity
        table[2, 2] = '$2.00' # Change price

        # Convert to markdown and back
        markdown_str = table.to_markdown()
        
        # Verify the markdown contains our changes
        assert 'Item' in markdown_str
        assert '15' in markdown_str 
        assert '$2.00' in markdown_str

        # Test that we can create a new table from this markdown
        # (This tests that our changes maintained valid table structure)
        recreated = MarkdownTableElement.from_markdown("Recreated", markdown_str.split('\n', 1)[1])
        
        # Verify the recreated table has our changes
        assert recreated[0, 0].strip() == 'Item'
        assert recreated[1, 1].strip() == '15'
        assert recreated[2, 2].strip() == '$2.00'