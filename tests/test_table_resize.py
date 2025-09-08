"""
Tests for TableElement resize functionality with fix_width parameter.
"""

import pytest
from unittest.mock import Mock, patch

from gslides_api.element.table import TableElement
from gslides_api.table import Table, TableColumnProperties, TableRow, TableCell
from gslides_api.table_cell import TableCellLocation
from gslides_api.domain_old import Dimension, Unit, Size, Transform
from gslides_api.element.text_content import TextContent
from gslides_api.request.table import (
    InsertTableColumnsRequest,
    DeleteTableColumnRequest,
    UpdateTableColumnPropertiesRequest,
    InsertTableRowsRequest,
    DeleteTableRowRequest,
)


class TestTableElementResize:
    """Test TableElement resize functionality with fix_width parameter."""

    def create_test_table(self, rows: int = 3, columns: int = 4):
        """Create a test table element with specified dimensions."""
        # Create table columns with width information
        table_columns = []
        for i in range(columns):
            table_columns.append(
                TableColumnProperties(columnWidth=Dimension(magnitude=100, unit=Unit.PT))
            )

        # Create table rows
        table_rows = []
        for row_idx in range(rows):
            cells = []
            for col_idx in range(columns):
                cells.append(TableCell(text=TextContent(textElements=[])))
            table_rows.append(TableRow(tableCells=cells))

        table = Table(rows=rows, columns=columns, tableColumns=table_columns, tableRows=table_rows)

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_resize_requests_no_changes(self):
        """Test that no requests are generated when dimensions don't change."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(3, 4)
        assert requests == []

    def test_resize_requests_add_rows_only(self):
        """Test adding rows doesn't generate column width requests."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(5, 4)

        assert len(requests) == 1
        assert isinstance(requests[0], InsertTableRowsRequest)
        assert requests[0].number == 2

    def test_resize_requests_delete_rows_only(self):
        """Test deleting rows doesn't generate column width requests."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(2, 4)

        assert len(requests) == 1
        assert isinstance(requests[0], DeleteTableRowRequest)

    def test_resize_requests_add_columns_fix_width_true(self):
        """Test adding columns with fix_width=True (default behavior)."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(3, 6, fix_width=True)

        # Should only have the insert request, no width adjustment
        assert len(requests) == 1
        assert isinstance(requests[0], InsertTableColumnsRequest)
        assert requests[0].number == 2

    def test_resize_requests_add_columns_fix_width_false(self):
        """Test adding columns with fix_width=False preserves original widths."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(3, 6, fix_width=False)

        # Should have insert request plus width adjustment requests
        insert_requests = [r for r in requests if isinstance(r, InsertTableColumnsRequest)]
        update_requests = [r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)]

        assert len(insert_requests) == 1
        assert insert_requests[0].number == 2

        # Should have update requests for all 6 columns (4 original + 2 new)
        assert len(update_requests) == 6

        # Check that original columns preserve their width
        for i in range(4):
            update_req = update_requests[i]
            assert update_req.columnIndices == [i]
            assert update_req.tableColumnProperties.columnWidth.magnitude == 100
            assert update_req.fields == "columnWidth"

        # Check that new columns get rightmost column width
        for i in range(4, 6):
            update_req = update_requests[i]
            assert update_req.columnIndices == [i]
            assert update_req.tableColumnProperties.columnWidth.magnitude == 100

    def test_resize_requests_delete_columns_fix_width_true(self):
        """Test deleting columns with fix_width=True maintains table width."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(3, 2, fix_width=True)

        delete_requests = [r for r in requests if isinstance(r, DeleteTableColumnRequest)]
        update_requests = [r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)]

        # Should have delete requests for 2 columns
        assert len(delete_requests) == 2

        # Should have width adjustment requests for remaining 2 columns
        assert len(update_requests) == 2

        # Each remaining column should be expanded to maintain total width
        # Original total: 4 * 100 = 400pt, remaining columns: 2, so each should be 200pt
        for update_req in update_requests:
            assert len(update_req.columnIndices) == 1
            assert update_req.tableColumnProperties.columnWidth.magnitude == 200.0
            assert update_req.fields == "columnWidth"

    def test_resize_requests_delete_columns_fix_width_false(self):
        """Test deleting columns with fix_width=False doesn't adjust widths."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(3, 2, fix_width=False)

        delete_requests = [r for r in requests if isinstance(r, DeleteTableColumnRequest)]
        update_requests = [r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)]

        # Should have delete requests for 2 columns
        assert len(delete_requests) == 2

        # Should NOT have width adjustment requests
        assert len(update_requests) == 0

    def test_resize_requests_no_column_width_info(self):
        """Test behavior when column width information is not available."""
        # Create table without column width information
        table = Table(rows=3, columns=4, tableColumns=None)
        table_element = TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Should only generate insert/delete requests, no width adjustments
        requests_add = table_element.resize_requests(3, 6, fix_width=False)
        requests_delete = table_element.resize_requests(3, 2, fix_width=True)

        insert_requests = [r for r in requests_add if isinstance(r, InsertTableColumnsRequest)]
        delete_requests = [r for r in requests_delete if isinstance(r, DeleteTableColumnRequest)]
        update_requests = [
            r
            for r in requests_add + requests_delete
            if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

        assert len(insert_requests) == 1
        assert len(delete_requests) == 2
        assert len(update_requests) == 0  # No width adjustments possible

    def test_resize_requests_invalid_dimensions(self):
        """Test that invalid dimensions raise ValueError."""
        table_element = self.create_test_table(3, 4)

        with pytest.raises(ValueError, match="Table must have at least 1 row and 1 column"):
            table_element.resize_requests(0, 4)

        with pytest.raises(ValueError, match="Table must have at least 1 row and 1 column"):
            table_element.resize_requests(3, 0)

    def test_resize_requests_mixed_changes(self):
        """Test simultaneous row and column changes."""
        table_element = self.create_test_table(3, 4)
        requests = table_element.resize_requests(5, 6, fix_width=False)

        row_requests = [
            r for r in requests if isinstance(r, (InsertTableRowsRequest, DeleteTableRowRequest))
        ]
        col_requests = [
            r
            for r in requests
            if isinstance(r, (InsertTableColumnsRequest, DeleteTableColumnRequest))
        ]
        update_requests = [r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)]

        # Should have row insertion
        assert len(row_requests) == 1
        assert isinstance(row_requests[0], InsertTableRowsRequest)

        # Should have column insertion
        assert len(col_requests) == 1
        assert isinstance(col_requests[0], InsertTableColumnsRequest)

        # Should have width preservation requests
        assert len(update_requests) == 6

    @patch("gslides_api.element.table.default_api_client")
    def test_resize_method_calls_resize_requests(self, mock_api_client):
        """Test that resize method properly calls resize_requests and executes."""
        mock_client_instance = Mock()
        mock_api_client.__bool__.return_value = True
        mock_api_client.batch_update.return_value = {"replies": []}

        table_element = self.create_test_table(3, 4)

        # Mock resize_requests to return some test requests
        with patch.object(TableElement, "resize_requests") as mock_resize_requests:
            mock_requests = [Mock()]
            mock_resize_requests.return_value = mock_requests

            # Test with fix_width=True
            table_element.resize(3, 6, fix_width=True, api_client=mock_client_instance)

            mock_resize_requests.assert_called_once_with(3, 6, True)
            mock_client_instance.batch_update.assert_called_once()

    @patch("gslides_api.element.table.default_api_client")
    def test_resize_method_no_requests(self, mock_api_client):
        """Test that resize method handles no requests case."""
        mock_client_instance = Mock()
        mock_api_client.__bool__.return_value = True

        table_element = self.create_test_table(3, 4)

        # Test with same dimensions (no changes needed)
        table_element.resize(3, 4, api_client=mock_client_instance)

        # Should not call batch_update
        mock_client_instance.batch_update.assert_not_called()

    def test_proportional_width_calculation_edge_cases(self):
        """Test edge cases in proportional width calculations."""
        # Test with uneven column widths
        table_columns = [
            TableColumnProperties(columnWidth=Dimension(magnitude=50, unit=Unit.PT)),
            TableColumnProperties(columnWidth=Dimension(magnitude=100, unit=Unit.PT)),
            TableColumnProperties(columnWidth=Dimension(magnitude=150, unit=Unit.PT)),
            TableColumnProperties(columnWidth=Dimension(magnitude=200, unit=Unit.PT)),
        ]

        table = Table(rows=3, columns=4, tableColumns=table_columns)
        table_element = TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=500, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Delete 2 columns, keep first 2
        requests = table_element.resize_requests(3, 2, fix_width=True)

        update_requests = [r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)]
        assert len(update_requests) == 2

        # Original total: 50 + 100 + 150 + 200 = 500pt
        # Remaining columns original widths: 50 + 100 = 150pt
        # Expansion factor: 500 / 150 = 3.333...
        # New widths should be: 50 * 3.333... ≈ 166.67, 100 * 3.333... ≈ 333.33

        widths = [req.tableColumnProperties.columnWidth.magnitude for req in update_requests]
        widths.sort()  # Sort to make comparison easier

        assert abs(widths[0] - 166.67) < 0.1  # Allow small floating point error
        assert abs(widths[1] - 333.33) < 0.1
