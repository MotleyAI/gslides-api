"""
Tests for TableElement handling of cells with text=None.

This tests the fix for the AttributeError that occurred when a table cell
exists in the table structure but has text=None (empty cell from API).
"""

import pytest

from gslides_api.element.table import TableElement
from gslides_api.domain.table import (
    Table,
    TableCell,
    TableColumnProperties,
    TableRow,
    TableRowProperties,
)
from gslides_api.domain.table_cell import TableCellLocation
from gslides_api.domain.domain import Dimension, Size, Transform, Unit
from gslides_api.element.text_content import TextContent


class TestTableCellNoneText:
    """Test TableElement handles cells with text=None correctly."""

    def create_table_with_none_text_cells(self, rows: int = 2, columns: int = 2):
        """Create a test table where cells have text=None (empty cells)."""
        # Create table columns
        table_columns = []
        for i in range(columns):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        # Create table rows with cells that have text=None
        table_rows = []
        for row_idx in range(rows):
            cells = []
            for col_idx in range(columns):
                # Create cell with text=None (simulates empty cell from API)
                cells.append(TableCell(text=None))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(magnitude=50, unit=Unit.PT),
                    tableRowProperties=TableRowProperties(
                        minRowHeight=Dimension(magnitude=50, unit=Unit.PT)
                    ),
                )
            )

        table = Table(
            rows=rows,
            columns=columns,
            tableColumns=table_columns,
            tableRows=table_rows,
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def create_table_with_mixed_text_cells(self, rows: int = 2, columns: int = 2):
        """Create a test table with mix of cells with text and cells with text=None."""
        table_columns = []
        for i in range(columns):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        table_rows = []
        for row_idx in range(rows):
            cells = []
            for col_idx in range(columns):
                # Alternate between cells with text and cells without
                if (row_idx + col_idx) % 2 == 0:
                    cells.append(TableCell(text=TextContent(textElements=[])))
                else:
                    cells.append(TableCell(text=None))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(magnitude=50, unit=Unit.PT),
                    tableRowProperties=TableRowProperties(
                        minRowHeight=Dimension(magnitude=50, unit=Unit.PT)
                    ),
                )
            )

        table = Table(
            rows=rows,
            columns=columns,
            tableColumns=table_columns,
            tableRows=table_rows,
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_write_text_to_cell_with_none_text(self):
        """Test write_text_to_cell_requests handles cell with text=None."""
        table_element = self.create_table_with_none_text_cells()

        # Verify the cell's text is indeed None
        cell = table_element[0, 0]
        assert cell.text is None

        # This should NOT raise AttributeError
        location = TableCellLocation(rowIndex=0, columnIndex=0)
        requests = table_element.write_text_to_cell_requests(
            text="Test content",
            location=location,
        )

        # Should return valid requests
        assert len(requests) > 0
        # Verify objectId is set on all requests
        for r in requests:
            assert r.objectId == "test-table-id"

    def test_write_text_to_all_none_text_cells(self):
        """Test write_text_to_cell_requests for all cells with text=None."""
        table_element = self.create_table_with_none_text_cells(rows=2, columns=2)

        # Write to all cells - none should raise an error
        for row in range(2):
            for col in range(2):
                location = TableCellLocation(rowIndex=row, columnIndex=col)
                requests = table_element.write_text_to_cell_requests(
                    text=f"Cell ({row}, {col})",
                    location=location,
                )
                assert len(requests) > 0

    def test_write_text_to_mixed_cells(self):
        """Test write_text_to_cell_requests with mix of text and None cells."""
        table_element = self.create_table_with_mixed_text_cells(rows=2, columns=2)

        # Write to all cells - mixed cells should all work
        for row in range(2):
            for col in range(2):
                location = TableCellLocation(rowIndex=row, columnIndex=col)
                requests = table_element.write_text_to_cell_requests(
                    text=f"Cell ({row}, {col})",
                    location=location,
                )
                assert len(requests) > 0

    def test_delete_text_in_cell_with_none_text(self):
        """Test delete_text_in_cell_requests handles cell with text=None."""
        table_element = self.create_table_with_none_text_cells()

        # Verify the cell's text is indeed None
        cell = table_element[0, 0]
        assert cell.text is None

        # This should NOT raise AttributeError
        location = TableCellLocation(rowIndex=0, columnIndex=0)
        requests = table_element.delete_text_in_cell_requests(location=location)

        # Should return valid requests (even if empty list for empty cell)
        assert isinstance(requests, list)

    def test_delete_text_in_all_none_text_cells(self):
        """Test delete_text_in_cell_requests for all cells with text=None."""
        table_element = self.create_table_with_none_text_cells(rows=2, columns=2)

        # Delete from all cells - none should raise an error
        for row in range(2):
            for col in range(2):
                location = TableCellLocation(rowIndex=row, columnIndex=col)
                requests = table_element.delete_text_in_cell_requests(location=location)
                assert isinstance(requests, list)

    def test_delete_text_in_mixed_cells(self):
        """Test delete_text_in_cell_requests with mix of text and None cells."""
        table_element = self.create_table_with_mixed_text_cells(rows=2, columns=2)

        # Delete from all cells - mixed cells should all work
        for row in range(2):
            for col in range(2):
                location = TableCellLocation(rowIndex=row, columnIndex=col)
                requests = table_element.delete_text_in_cell_requests(location=location)
                assert isinstance(requests, list)

    def test_content_update_requests_with_none_text_cells(self):
        """Test content_update_requests handles table with text=None cells."""
        table_element = self.create_table_with_none_text_cells(rows=2, columns=2)

        # Create markdown content for the table
        markdown_table = """| Header 1 | Header 2 |
| --- | --- |
| Data 1 | Data 2 |"""

        # This should NOT raise AttributeError
        requests = table_element.content_update_requests(
            markdown_elem=markdown_table,
            check_shape=False,  # Don't check shape since our test table is smaller
        )

        # Should return valid requests
        assert len(requests) > 0

    def test_write_text_with_markdown_to_none_text_cell(self):
        """Test writing markdown content to a cell with text=None."""
        table_element = self.create_table_with_none_text_cells()

        location = TableCellLocation(rowIndex=0, columnIndex=0)
        requests = table_element.write_text_to_cell_requests(
            text="**Bold** and *italic*",
            location=location,
            as_markdown=True,
        )

        # Should return valid requests with formatting
        assert len(requests) > 0
