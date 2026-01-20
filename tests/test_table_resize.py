"""
Tests for TableElement resize functionality with fix_width parameter.
"""

import pytest
from unittest.mock import Mock, patch

from gslides_api.agnostic.units import OutputUnit
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
from gslides_api.request.table import (
    InsertTableColumnsRequest,
    DeleteTableColumnRequest,
    UpdateTableColumnPropertiesRequest,
    UpdateTableRowPropertiesRequest,
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
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        # Create table rows with height information
        table_rows = []
        for row_idx in range(rows):
            cells = []
            for col_idx in range(columns):
                cells.append(TableCell(text=TextContent(textElements=[])))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(
                        magnitude=50, unit=Unit.PT
                    ),  # Default row height
                    tableRowProperties=TableRowProperties(
                        minRowHeight=Dimension(magnitude=50, unit=Unit.PT)
                    ),
                )
            )

        table = Table(
            rows=rows, columns=columns, tableColumns=table_columns, tableRows=table_rows
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_resize_requests_no_changes(self):
        """Test that no requests are generated when dimensions don't change."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(3, 4)
        assert requests == []
        assert scale_factor == 1.0

    def test_resize_requests_add_rows_only(self):
        """Test adding rows doesn't generate column width requests."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(5, 4)

        assert len(requests) == 1
        assert isinstance(requests[0], InsertTableRowsRequest)
        assert requests[0].number == 2
        assert scale_factor == 1.0  # No fix_height, so scale is 1.0

    def test_resize_requests_delete_rows_only(self):
        """Test deleting rows doesn't generate column width requests."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(2, 4)

        assert len(requests) == 1
        assert isinstance(requests[0], DeleteTableRowRequest)
        assert scale_factor == 1.0

    def test_resize_requests_add_columns_fix_width_true(self):
        """Test adding columns with fix_width=True (default behavior)."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(3, 6, fix_width=True)

        # Should only have the insert request, no width adjustment
        assert len(requests) == 1
        assert isinstance(requests[0], InsertTableColumnsRequest)
        assert requests[0].number == 2
        assert scale_factor == 1.0

    def test_resize_requests_add_columns_fix_width_false(self):
        """Test adding columns with fix_width=False preserves original widths."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(3, 6, fix_width=False)

        # Should have insert request plus width adjustment requests
        insert_requests = [
            r for r in requests if isinstance(r, InsertTableColumnsRequest)
        ]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

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
        requests, scale_factor = table_element.resize_requests(3, 2, fix_width=True)

        delete_requests = [
            r for r in requests if isinstance(r, DeleteTableColumnRequest)
        ]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

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
        requests, scale_factor = table_element.resize_requests(3, 2, fix_width=False)

        delete_requests = [
            r for r in requests if isinstance(r, DeleteTableColumnRequest)
        ]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

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
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Should only generate insert/delete requests, no width adjustments
        requests_add, _ = table_element.resize_requests(3, 6, fix_width=False)
        requests_delete, _ = table_element.resize_requests(3, 2, fix_width=True)

        insert_requests = [
            r for r in requests_add if isinstance(r, InsertTableColumnsRequest)
        ]
        delete_requests = [
            r for r in requests_delete if isinstance(r, DeleteTableColumnRequest)
        ]
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

        with pytest.raises(
            ValueError, match="Table must have at least 1 row and 1 column"
        ):
            table_element.resize_requests(0, 4)

        with pytest.raises(
            ValueError, match="Table must have at least 1 row and 1 column"
        ):
            table_element.resize_requests(3, 0)

    def test_resize_requests_mixed_changes(self):
        """Test simultaneous row and column changes."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(5, 6, fix_width=False)

        row_requests = [
            r
            for r in requests
            if isinstance(r, (InsertTableRowsRequest, DeleteTableRowRequest))
        ]
        col_requests = [
            r
            for r in requests
            if isinstance(r, (InsertTableColumnsRequest, DeleteTableColumnRequest))
        ]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

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

        # Mock resize_requests to return some test requests (now returns tuple)
        with patch.object(TableElement, "resize_requests") as mock_resize_requests:
            mock_requests = [Mock()]
            mock_resize_requests.return_value = (mock_requests, 0.5)

            # Test with fix_width=True (fix_height uses default False)
            result = table_element.resize(3, 6, fix_width=True, api_client=mock_client_instance)

            mock_resize_requests.assert_called_once_with(3, 6, True, False, target_height_emu=None)
            mock_client_instance.batch_update.assert_called_once()
            assert result == 0.5  # resize now returns the scale factor

    @patch("gslides_api.element.table.default_api_client")
    def test_resize_method_no_requests(self, mock_api_client):
        """Test that resize method handles no requests case."""
        mock_client_instance = Mock()
        mock_api_client.__bool__.return_value = True

        table_element = self.create_test_table(3, 4)

        # Test with same dimensions (no changes needed)
        result = table_element.resize(3, 4, api_client=mock_client_instance)

        # Should not call batch_update
        mock_client_instance.batch_update.assert_not_called()
        assert result == 1.0  # No changes, should return default scale factor

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
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Delete 2 columns, keep first 2
        requests, _ = table_element.resize_requests(3, 2, fix_width=True)

        update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]
        assert len(update_requests) == 2

        # Original total: 50 + 100 + 150 + 200 = 500pt
        # Remaining columns original widths: 50 + 100 = 150pt
        # Expansion factor: 500 / 150 = 3.333...
        # New widths should be: 50 * 3.333... ≈ 166.67, 100 * 3.333... ≈ 333.33

        widths = [
            req.tableColumnProperties.columnWidth.magnitude for req in update_requests
        ]
        widths.sort()  # Sort to make comparison easier

        assert abs(widths[0] - 166.67) < 0.1  # Allow small floating point error
        assert abs(widths[1] - 333.33) < 0.1

    # ===== NEW TESTS FOR fix_height PARAMETER =====

    def test_resize_requests_add_rows_fix_height_false(self):
        """Test adding rows with fix_height=False (default behavior) doesn't adjust heights."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(5, 4, fix_height=False)

        insert_requests = [r for r in requests if isinstance(r, InsertTableRowsRequest)]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]

        assert len(insert_requests) == 1
        assert insert_requests[0].number == 2

        # Should NOT have height adjustment requests
        assert len(update_requests) == 0

    def test_resize_requests_add_rows_fix_height_true(self):
        """Test adding rows with fix_height=True maintains table height."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(5, 4, fix_height=True)

        insert_requests = [r for r in requests if isinstance(r, InsertTableRowsRequest)]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]

        assert len(insert_requests) == 1
        assert insert_requests[0].number == 2

        # Should have height adjustment requests for all 5 rows (3 original + 2 new)
        assert len(update_requests) == 5

        # Original total height: 3 * 50 = 150pt
        # New height per row: 150 / 5 = 30pt
        for update_req in update_requests:
            assert len(update_req.rowIndices) == 1
            assert update_req.tableRowProperties.minRowHeight.magnitude == 30.0
            assert update_req.fields == "minRowHeight"

    def test_resize_requests_delete_rows_fix_height_false(self):
        """Test deleting rows with fix_height=False (default) doesn't adjust heights."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(2, 4, fix_height=False)

        delete_requests = [r for r in requests if isinstance(r, DeleteTableRowRequest)]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]

        assert len(delete_requests) == 1

        # Should NOT have height adjustment requests
        assert len(update_requests) == 0

    def test_resize_requests_delete_rows_fix_height_true(self):
        """Test deleting rows with fix_height=True maintains table height."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(2, 4, fix_height=True)

        delete_requests = [r for r in requests if isinstance(r, DeleteTableRowRequest)]
        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]

        assert len(delete_requests) == 1

        # Should have height adjustment requests for remaining 2 rows
        assert len(update_requests) == 2

        # Original total height: 3 * 50 = 150pt, remaining rows: 2
        # Each remaining row should be expanded to maintain total height: 150 / 2 = 75pt
        for update_req in update_requests:
            assert len(update_req.rowIndices) == 1
            assert update_req.tableRowProperties.minRowHeight.magnitude == 75.0
            assert update_req.fields == "minRowHeight"

    def test_resize_requests_no_row_height_info(self):
        """Test behavior when row height information is not available."""
        # Create table without row height information
        table_columns = []
        for i in range(4):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        table = Table(rows=3, columns=4, tableColumns=table_columns, tableRows=None)
        table_element = TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Should only generate insert/delete requests, no height adjustments
        requests_add, _ = table_element.resize_requests(5, 4, fix_height=True)
        requests_delete, _ = table_element.resize_requests(2, 4, fix_height=True)

        insert_requests = [
            r for r in requests_add if isinstance(r, InsertTableRowsRequest)
        ]
        delete_requests = [
            r for r in requests_delete if isinstance(r, DeleteTableRowRequest)
        ]
        update_requests = [
            r
            for r in requests_add + requests_delete
            if isinstance(r, UpdateTableRowPropertiesRequest)
        ]

        assert len(insert_requests) == 1
        assert len(delete_requests) == 1
        assert len(update_requests) == 0  # No height adjustments possible

    def test_proportional_height_calculation_edge_cases(self):
        """Test edge cases in proportional height calculations."""
        # Create table with uneven row heights
        table_columns = []
        for i in range(4):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        table_rows = []
        row_heights = [30, 60, 90]  # Different heights
        for row_idx, height in enumerate(row_heights):
            cells = []
            for col_idx in range(4):
                cells.append(TableCell(text=TextContent(textElements=[])))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(magnitude=height, unit=Unit.PT),
                    tableRowProperties=TableRowProperties(
                        minRowHeight=Dimension(magnitude=height, unit=Unit.PT)
                    ),
                )
            )

        table = Table(
            rows=3, columns=4, tableColumns=table_columns, tableRows=table_rows
        )
        table_element = TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=180, unit=Unit.PT),  # 30 + 60 + 90
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Delete 1 row (keep first 2), should maintain total height
        requests, _ = table_element.resize_requests(2, 4, fix_height=True)

        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]
        assert len(update_requests) == 2

        # Original total: 30 + 60 + 90 = 180pt
        # Remaining rows original heights: 30 + 60 = 90pt
        # Expansion factor: 180 / 90 = 2.0
        # New heights should be: 30 * 2.0 = 60pt, 60 * 2.0 = 120pt

        heights = [
            req.tableRowProperties.minRowHeight.magnitude for req in update_requests
        ]
        heights.sort()  # Sort to make comparison easier

        assert heights[0] == 60.0
        assert heights[1] == 120.0

    def test_resize_requests_mixed_fix_width_fix_height(self):
        """Test simultaneous row and column changes with both fix_width and fix_height."""
        table_element = self.create_test_table(3, 4)
        requests, scale_factor = table_element.resize_requests(5, 6, fix_width=True, fix_height=True)

        row_requests = [
            r
            for r in requests
            if isinstance(r, (InsertTableRowsRequest, DeleteTableRowRequest))
        ]
        col_requests = [
            r
            for r in requests
            if isinstance(r, (InsertTableColumnsRequest, DeleteTableColumnRequest))
        ]
        row_update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]
        col_update_requests = [
            r for r in requests if isinstance(r, UpdateTableColumnPropertiesRequest)
        ]

        # Should have row insertion
        assert len(row_requests) == 1
        assert isinstance(row_requests[0], InsertTableRowsRequest)

        # Should have column insertion
        assert len(col_requests) == 1
        assert isinstance(col_requests[0], InsertTableColumnsRequest)

        # Should have height adjustment requests (fix_height=True)
        assert len(row_update_requests) == 5  # All 5 rows

        # Should NOT have width adjustment requests (fix_width=True is default for adding columns)
        assert len(col_update_requests) == 0

    @patch("gslides_api.element.table.default_api_client")
    def test_resize_method_with_fix_height(self, mock_api_client):
        """Test that resize method properly passes fix_height parameter."""
        mock_client_instance = Mock()
        mock_api_client.__bool__.return_value = True
        mock_api_client.batch_update.return_value = {"replies": []}

        table_element = self.create_test_table(3, 4)

        # Mock resize_requests to verify parameters (now returns tuple)
        with patch.object(TableElement, "resize_requests") as mock_resize_requests:
            mock_requests = [Mock()]
            mock_resize_requests.return_value = (mock_requests, 0.6)

            # Test with both fix_width and fix_height
            result = table_element.resize(
                5, 6, fix_width=False, fix_height=True, api_client=mock_client_instance
            )

            mock_resize_requests.assert_called_once_with(5, 6, False, True, target_height_emu=None)
            mock_client_instance.batch_update.assert_called_once()
            assert result == 0.6

    def test_resize_requests_height_adjustment_row_indices(self):
        """Test that height adjustment requests target the correct row indices."""
        table_element = self.create_test_table(4, 3)

        # Delete 2 rows (keep rows 0, 1)
        requests, _ = table_element.resize_requests(2, 3, fix_height=True)

        update_requests = [
            r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)
        ]
        assert len(update_requests) == 2

        # Check that the correct row indices are targeted
        targeted_rows = []
        for req in update_requests:
            targeted_rows.extend(req.rowIndices)

        targeted_rows.sort()
        assert targeted_rows == [0, 1]  # First two rows should remain


class TestTableBorderWeight:
    """Tests for border weight functionality in table resize."""

    def create_test_table_with_borders(self, rows: int = 3, columns: int = 4, border_weight_emu: float = 38100):
        """Create a test table element with border information."""
        from gslides_api.domain.table import (
            TableBorderRow,
            TableBorderCell,
            TableBorderProperties,
        )

        # Create table columns with width information
        table_columns = []
        for i in range(columns):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        # Create table rows with height information (in EMU for consistency)
        table_rows = []
        for row_idx in range(rows):
            cells = []
            for col_idx in range(columns):
                cells.append(TableCell(text=TextContent(textElements=[])))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(
                        magnitude=914400,  # 1 inch in EMU
                        unit=Unit.EMU
                    ),
                    tableRowProperties=TableRowProperties(
                        minRowHeight=Dimension(magnitude=914400, unit=Unit.EMU)
                    ),
                )
            )

        # Create horizontal border rows (N+1 rows of borders for N data rows)
        horizontal_border_rows = []
        for i in range(rows + 1):
            border_cells = []
            for j in range(columns):
                border_cells.append(
                    TableBorderCell(
                        tableBorderProperties=TableBorderProperties(
                            weight=Dimension(magnitude=border_weight_emu, unit=Unit.EMU)
                        )
                    )
                )
            horizontal_border_rows.append(
                TableBorderRow(tableBorderCells=border_cells)
            )

        table = Table(
            rows=rows,
            columns=columns,
            tableColumns=table_columns,
            tableRows=table_rows,
            horizontalBorderRows=horizontal_border_rows,
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_get_horizontal_border_weight_in_emu(self):
        """Test that get_horizontal_border_weight returns correct EMU value."""
        border_weight_emu = 38100
        table_element = self.create_test_table_with_borders(border_weight_emu=border_weight_emu)

        weight = table_element.get_horizontal_border_weight(units=OutputUnit.EMU)
        assert weight == border_weight_emu

    def test_get_horizontal_border_weight_in_inches(self):
        """Test that get_horizontal_border_weight converts to inches correctly."""
        border_weight_emu = 914400  # 1 inch
        table_element = self.create_test_table_with_borders(border_weight_emu=border_weight_emu)

        weight = table_element.get_horizontal_border_weight(units=OutputUnit.IN)
        assert weight == pytest.approx(1.0)

    def test_get_horizontal_border_weight_in_points(self):
        """Test that get_horizontal_border_weight converts to points correctly."""
        border_weight_emu = 12700  # 1 point
        table_element = self.create_test_table_with_borders(border_weight_emu=border_weight_emu)

        weight = table_element.get_horizontal_border_weight(units=OutputUnit.PT)
        assert weight == pytest.approx(1.0)

    def test_get_horizontal_border_weight_no_borders(self):
        """Test that get_horizontal_border_weight returns 0 when no borders exist."""
        from gslides_api.element.table import TableElement as TE

        # Create table without borders
        table_columns = [
            TableColumnProperties(columnWidth=Dimension(magnitude=100, unit=Unit.PT))
            for _ in range(4)
        ]
        table_rows = []
        for _ in range(3):
            cells = [TableCell(text=TextContent(textElements=[])) for _ in range(4)]
            table_rows.append(TableRow(
                tableCells=cells,
                rowHeight=Dimension(magnitude=50, unit=Unit.PT),
            ))

        table = Table(
            rows=3, columns=4, tableColumns=table_columns, tableRows=table_rows
        )
        table_element = TE(
            objectId="test", table=table, slide_id="", presentation_id="",
            size=Size(width=Dimension(magnitude=400, unit=Unit.PT), height=Dimension(magnitude=300, unit=Unit.PT)),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
        )

        weight = table_element.get_horizontal_border_weight(units=OutputUnit.IN)
        assert weight == 0.0

    def test_generate_border_weight_requests(self):
        """Test that _generate_border_weight_requests creates correct API requests."""
        from gslides_api.request.table import UpdateTableBorderPropertiesRequest

        table_element = self.create_test_table_with_borders()
        new_weight_emu = 19050

        requests = table_element._generate_border_weight_requests(new_weight_emu)

        # Should generate 3 requests: TOP, BOTTOM, INNER_HORIZONTAL
        assert len(requests) == 3
        positions = {req.borderPosition for req in requests}
        assert positions == {"TOP", "BOTTOM", "INNER_HORIZONTAL"}

        for req in requests:
            assert isinstance(req, UpdateTableBorderPropertiesRequest)
            assert req.objectId == "test-table-id"
            assert req.tableBorderProperties.weight.magnitude == new_weight_emu
            assert req.fields == "weight"

    def test_resize_requests_with_target_total_height_scales_rows_and_borders(self):
        """Test that resize_requests with target_height_emu scales both rows and borders."""
        from gslides_api.request.table import UpdateTableBorderPropertiesRequest

        # 3 rows x 1 inch each = 3 inches of row height
        # 4 borders x 38100 EMU (0.0417 inches) = ~0.167 inches of borders
        # Total expected ~3.167 inches
        table_element = self.create_test_table_with_borders(
            rows=3,
            columns=4,
            border_weight_emu=38100,  # ~0.0417 inches
        )

        # Add 3 rows (total 6 rows), but constrain to 4 inches total
        # Expected 6 rows at 1 inch each = 6 inches row height
        # Expected 7 borders at 38100 EMU = ~0.29 inches
        # Total expected without constraint = ~6.29 inches
        # With 4 inch constraint, scale factor = 4 / 6.29 = ~0.636
        target_height_emu = 4 * 914400  # 4 inches in EMU

        requests, scale_factor = table_element.resize_requests(
            n_rows=6, n_columns=4, target_height_emu=target_height_emu
        )

        # Should have: insert row request, row height requests, border weight request
        insert_requests = [r for r in requests if isinstance(r, InsertTableRowsRequest)]
        row_height_requests = [r for r in requests if isinstance(r, UpdateTableRowPropertiesRequest)]
        border_requests = [r for r in requests if isinstance(r, UpdateTableBorderPropertiesRequest)]

        assert len(insert_requests) == 1
        assert len(row_height_requests) == 6  # All 6 rows get height adjustments
        assert len(border_requests) == 3  # Border weight requests (TOP, BOTTOM, INNER_HORIZONTAL)

        # The border weights should all be scaled down
        for border_req in border_requests:
            new_border_weight = border_req.tableBorderProperties.weight.magnitude
            # New weight should be less than original 38100 EMU
            assert new_border_weight < 38100

        # Scale factor should be target_height / expected_total (including borders)
        # Expected: 6 rows at 914400 EMU = 5486400 EMU row height
        # Expected: 7 borders at 38100 EMU = 266700 EMU
        # Expected total = 5486400 + 266700 = 5753100 EMU
        # Target = 4 * 914400 = 3657600 EMU
        # Scale = 3657600 / 5753100 ≈ 0.6357
        expected_row_heights = 3 * 914400 * (6 / 3)  # 5486400
        expected_border_height = 38100 * 7  # 266700
        expected_total = expected_row_heights + expected_border_height  # 5753100
        target_total = 4 * 914400  # 3657600
        expected_scale = target_total / expected_total
        assert scale_factor == pytest.approx(expected_scale, rel=1e-6)

    def test_resize_requests_no_border_request_when_border_weight_zero(self):
        """Test that no border request is generated when border weight is zero."""
        from gslides_api.request.table import UpdateTableBorderPropertiesRequest

        # Create table with zero border weight
        table_element = self.create_test_table_with_borders(
            rows=3,
            columns=4,
            border_weight_emu=0,  # No borders
        )

        target_height_emu = 2 * 914400  # 2 inches

        requests, _ = table_element.resize_requests(
            n_rows=6, n_columns=4, target_height_emu=target_height_emu
        )

        # Should NOT have border request when border weight is 0
        border_requests = [r for r in requests if isinstance(r, UpdateTableBorderPropertiesRequest)]
        assert len(border_requests) == 0


class TestCellPropertyPropagation:
    """Tests for cell property propagation when adding rows."""

    def create_test_table_with_cell_properties(self, rows: int = 2, columns: int = 3):
        """Create a test table element with cell properties (background, alignment) set."""
        from gslides_api.domain.table import TableCellProperties, TableCellBackgroundFill
        from gslides_api.domain.domain import (
            SolidFill,
            Color,
            RgbColor,
        )

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
                # Last row has specific cell properties (simulating a styled header/data row)
                cell_props = None
                if row_idx == rows - 1:
                    cell_props = TableCellProperties(
                        tableCellBackgroundFill=TableCellBackgroundFill(
                            solidFill=SolidFill(
                                color=Color(rgbColor=RgbColor(red=0.0, green=0.0, blue=0.5))
                            )
                        ),
                        contentAlignment="MIDDLE",
                    )
                cells.append(
                    TableCell(
                        text=TextContent(textElements=[]),
                        tableCellProperties=cell_props,
                    )
                )
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
            rows=rows, columns=columns, tableColumns=table_columns, tableRows=table_rows
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_resize_requests_add_rows_copies_cell_properties(self):
        """Test that adding rows copies cell properties from last row to new rows."""
        from gslides_api.request.table import UpdateTableCellPropertiesRequest

        table_element = self.create_test_table_with_cell_properties(rows=2, columns=3)

        requests, scale_factor = table_element.resize_requests(
            n_rows=5, n_columns=3, fix_height=True
        )

        # Find UpdateTableCellPropertiesRequest instances
        cell_prop_requests = [
            r for r in requests if isinstance(r, UpdateTableCellPropertiesRequest)
        ]

        # Should have 3 columns × 3 new rows = 9 cell property update requests
        assert len(cell_prop_requests) == 9

        # Each new row (2, 3, 4) × each column (0, 1, 2) should have a request
        for req in cell_prop_requests:
            assert req.tableRange.location.rowIndex in [2, 3, 4]  # New rows only
            assert req.tableRange.location.columnIndex in [0, 1, 2]
            assert "tableCellBackgroundFill" in req.fields
            assert "contentAlignment" in req.fields

    def test_resize_requests_no_cell_properties_when_shrinking(self):
        """Test that shrinking rows doesn't add cell property requests."""
        from gslides_api.request.table import UpdateTableCellPropertiesRequest

        table_element = self.create_test_table_with_cell_properties(rows=5, columns=3)

        requests, _ = table_element.resize_requests(n_rows=2, n_columns=3, fix_height=True)

        # Should NOT have cell property requests when shrinking
        cell_prop_requests = [
            r for r in requests if isinstance(r, UpdateTableCellPropertiesRequest)
        ]
        assert len(cell_prop_requests) == 0


class TestScaleFactorReturn:
    """Tests for scale factor return from resize_requests."""

    def create_test_table(self, rows: int = 2, columns: int = 3):
        """Create a simple test table."""
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
                cells.append(TableCell(text=TextContent(textElements=[])))
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
            rows=rows, columns=columns, tableColumns=table_columns, tableRows=table_rows
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_resize_requests_returns_scale_factor_with_fix_height(self):
        """Test that resize_requests returns correct font scale factor when fix_height=True."""
        table_element = self.create_test_table(rows=2, columns=3)

        requests, scale = table_element.resize_requests(
            n_rows=11, n_columns=3, fix_height=True
        )

        # Scale should be original_rows / new_rows = 2/11
        assert scale == pytest.approx(2 / 11, rel=1e-6)

    def test_resize_requests_returns_1_without_fix_height(self):
        """Test that resize_requests returns scale=1.0 when fix_height=False."""
        table_element = self.create_test_table(rows=2, columns=3)

        requests, scale = table_element.resize_requests(
            n_rows=11, n_columns=3, fix_height=False
        )

        assert scale == 1.0

    def test_resize_requests_returns_1_when_shrinking_rows(self):
        """Test that resize_requests returns scale=1.0 when reducing rows."""
        table_element = self.create_test_table(rows=10, columns=3)

        requests, scale = table_element.resize_requests(
            n_rows=2, n_columns=3, fix_height=True
        )

        # Shrinking doesn't need font scaling
        assert scale == 1.0

    def test_resize_requests_returns_1_when_same_rows(self):
        """Test that resize_requests returns scale=1.0 when rows don't change."""
        table_element = self.create_test_table(rows=5, columns=3)

        requests, scale = table_element.resize_requests(
            n_rows=5, n_columns=3, fix_height=True
        )

        assert scale == 1.0


class TestFontScaling:
    """Tests for font scaling in write_text_to_cell_requests."""

    def create_test_table_with_text_styles(self, rows: int = 2, columns: int = 2):
        """Create a test table with text styles (font size, color) in cells."""
        from gslides_api.domain.text import TextElement as TE, TextRun, TextStyle, ParagraphMarker
        from gslides_api.domain.domain import OptionalColor, Color, RgbColor

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
                # Create text content with specific font size (10pt)
                text_style = TextStyle(
                    fontSize=Dimension(magnitude=10, unit=Unit.PT),
                    foregroundColor=OptionalColor(
                        opaqueColor=Color(
                            rgbColor=RgbColor(red=1.0, green=1.0, blue=1.0)
                        )
                    ),
                )
                text_elements = [
                    TE(
                        startIndex=0,
                        endIndex=0,
                        paragraphMarker=ParagraphMarker(),
                    ),
                    TE(
                        startIndex=0,
                        endIndex=12,
                        textRun=TextRun(
                            content="Sample text\n",
                            style=text_style,
                        )
                    ),
                ]
                cells.append(
                    TableCell(
                        text=TextContent(textElements=text_elements),
                    )
                )
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
            rows=rows, columns=columns, tableColumns=table_columns, tableRows=table_rows
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

    def test_write_text_to_cell_requests_scales_fonts(self):
        """Test that font_scale_factor scales font sizes when writing text."""
        from gslides_api.request.request import UpdateTextStyleRequest

        table_element = self.create_test_table_with_text_styles(rows=2, columns=2)

        requests = table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=0, columnIndex=0),
            font_scale_factor=0.5,
        )

        # Find UpdateTextStyleRequest and check font size
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # At least one style request should exist
        assert len(style_requests) > 0

        # Check that font size is scaled
        for req in style_requests:
            if req.style and req.style.fontSize:
                # 10pt * 0.5 = 5pt
                assert req.style.fontSize.magnitude == 5

    def test_write_text_to_cell_requests_no_scaling_when_factor_is_1(self):
        """Test that font sizes are not modified when scale factor is 1.0."""
        from gslides_api.request.request import UpdateTextStyleRequest

        table_element = self.create_test_table_with_text_styles(rows=2, columns=2)

        requests = table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=0, columnIndex=0),
            font_scale_factor=1.0,
        )

        # Find UpdateTextStyleRequest and check font size
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # Font sizes should remain at original 10pt
        for req in style_requests:
            if req.style and req.style.fontSize:
                assert req.style.fontSize.magnitude == 10

    def test_write_text_to_cell_requests_deepcopy_preserves_original(self):
        """Test that font scaling uses deepcopy and doesn't modify original styles."""
        table_element = self.create_test_table_with_text_styles(rows=2, columns=2)

        # Get original styles
        original_styles = table_element.table.tableRows[0].tableCells[0].text.styles()
        original_font_size = original_styles[0].font_size_pt

        # Apply scaling
        table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=0, columnIndex=0),
            font_scale_factor=0.5,
        )

        # Original style should be unchanged
        current_styles = table_element.table.tableRows[0].tableCells[0].text.styles()
        current_font_size = current_styles[0].font_size_pt
        assert current_font_size == original_font_size

    def test_content_update_requests_passes_font_scale_factor(self):
        """Test that content_update_requests passes font_scale_factor to write_text_to_cell_requests."""
        table_element = self.create_test_table_with_text_styles(rows=3, columns=2)

        markdown = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |"

        # This should work without error
        requests = table_element.content_update_requests(
            markdown_elem=markdown,
            font_scale_factor=0.5,
        )

        assert len(requests) > 0


class TestTemplateStyles:
    """Tests for template_styles fallback in write_text_to_cell_requests."""

    def create_test_table_with_styled_rows(self, styled_rows: int = 2, empty_rows: int = 3, columns: int = 2):
        """Create a test table where first N rows have styles, rest are empty cells.

        Args:
            styled_rows: Number of rows with text content and styles
            empty_rows: Number of rows with empty cells (no text content)
            columns: Number of columns
        """
        from gslides_api.domain.text import TextElement as TE, TextRun, TextStyle, ParagraphMarker
        from gslides_api.domain.domain import OptionalColor, Color, RgbColor

        total_rows = styled_rows + empty_rows

        table_columns = []
        for i in range(columns):
            table_columns.append(
                TableColumnProperties(
                    columnWidth=Dimension(magnitude=100, unit=Unit.PT)
                )
            )

        table_rows = []
        for row_idx in range(total_rows):
            cells = []
            for col_idx in range(columns):
                if row_idx < styled_rows:
                    # Rows with text content and styles
                    text_style = TextStyle(
                        fontSize=Dimension(magnitude=10, unit=Unit.PT),
                        foregroundColor=OptionalColor(
                            opaqueColor=Color(
                                rgbColor=RgbColor(red=1.0, green=1.0, blue=1.0)
                            )
                        ),
                    )
                    text_elements = [
                        TE(
                            startIndex=0,
                            endIndex=0,
                            paragraphMarker=ParagraphMarker(),
                        ),
                        TE(
                            startIndex=0,
                            endIndex=12,
                            textRun=TextRun(
                                content="Sample text\n",
                                style=text_style,
                            )
                        ),
                    ]
                    cells.append(
                        TableCell(
                            text=TextContent(textElements=text_elements),
                        )
                    )
                else:
                    # Empty cells (simulating newly added rows)
                    cells.append(
                        TableCell(
                            text=TextContent(textElements=[]),
                        )
                    )
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
            rows=total_rows, columns=columns, tableColumns=table_columns, tableRows=table_rows
        )

        return TableElement(
            objectId="test-table-id",
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=250, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

    def test_write_text_uses_template_styles_for_empty_cell(self):
        """Test that template_styles is used when cell has no existing styles."""
        from gslides_api.request.request import UpdateTextStyleRequest
        from gslides_api.agnostic.text import RichStyle

        table_element = self.create_test_table_with_styled_rows(styled_rows=2, empty_rows=3, columns=2)

        # Define template styles with 10pt font
        template_styles = [RichStyle(font_size_pt=10.0, font_family="Arial")]

        # Write to an empty cell (row 3, which is in the empty_rows section)
        requests = table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=3, columnIndex=0),
            font_scale_factor=0.5,
            template_styles=template_styles,
        )

        # Find UpdateTextStyleRequest and check font size is scaled
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # At least one style request should exist
        assert len(style_requests) > 0

        # Font should be scaled: 10pt * 0.5 = 5pt
        for req in style_requests:
            if req.style and req.style.fontSize:
                assert req.style.fontSize.magnitude == 5

    def test_write_text_cell_styles_take_priority_over_template(self):
        """Test that existing cell styles take priority over template_styles."""
        from gslides_api.request.request import UpdateTextStyleRequest
        from gslides_api.agnostic.text import RichStyle

        table_element = self.create_test_table_with_styled_rows(styled_rows=2, empty_rows=3, columns=2)

        # Define template styles with DIFFERENT font size (20pt)
        template_styles = [RichStyle(font_size_pt=20.0, font_family="Arial")]

        # Write to a styled cell (row 0, which has 10pt font)
        requests = table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=0, columnIndex=0),
            font_scale_factor=0.5,
            template_styles=template_styles,
        )

        # Find UpdateTextStyleRequest and check font size
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # Font should be scaled from cell's 10pt, not template's 20pt
        # 10pt * 0.5 = 5pt
        for req in style_requests:
            if req.style and req.style.fontSize:
                assert req.style.fontSize.magnitude == 5

    def test_content_update_requests_extracts_template_from_styled_rows(self):
        """Test that content_update_requests extracts template styles from last styled row."""
        from gslides_api.request.request import UpdateTextStyleRequest

        # Create table with 2 styled rows and 3 empty rows
        table_element = self.create_test_table_with_styled_rows(styled_rows=2, empty_rows=3, columns=2)

        # Markdown to fill all 5 rows (GFM format with header separator)
        markdown = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n| 5 | 6 |\n| 7 | 8 |"

        requests = table_element.content_update_requests(
            markdown_elem=markdown,
            font_scale_factor=0.5,
        )

        # Find UpdateTextStyleRequest for cells in originally empty rows (rows 2, 3, 4)
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # Should have style requests for all cells (styles extracted from styled rows)
        assert len(style_requests) > 0

        # All font sizes should be scaled: 10pt * 0.5 = 5pt
        for req in style_requests:
            if req.style and req.style.fontSize:
                assert req.style.fontSize.magnitude == 5

    def test_content_update_requests_iterates_from_last_row_upward(self):
        """Test that template styles are extracted by iterating from last row upward."""
        from gslides_api.domain.text import TextElement as TE, TextRun, TextStyle, ParagraphMarker
        from gslides_api.domain.domain import OptionalColor, Color, RgbColor
        from gslides_api.request.request import UpdateTextStyleRequest

        # Create table where ONLY the FIRST row has styles, rest are empty
        # This tests that we iterate upward from last row until we find styles
        table_columns = [
            TableColumnProperties(columnWidth=Dimension(magnitude=100, unit=Unit.PT))
            for _ in range(2)
        ]

        table_rows = []
        for row_idx in range(5):
            cells = []
            for col_idx in range(2):
                if row_idx == 0:  # Only first row has styles
                    text_style = TextStyle(
                        fontSize=Dimension(magnitude=12, unit=Unit.PT),
                        foregroundColor=OptionalColor(
                            opaqueColor=Color(
                                rgbColor=RgbColor(red=1.0, green=0.0, blue=0.0)
                            )
                        ),
                    )
                    text_elements = [
                        TE(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
                        TE(startIndex=0, endIndex=5, textRun=TextRun(content="text\n", style=text_style)),
                    ]
                    cells.append(TableCell(text=TextContent(textElements=text_elements)))
                else:
                    # Empty cells
                    cells.append(TableCell(text=TextContent(textElements=[])))
            table_rows.append(
                TableRow(
                    tableCells=cells,
                    rowHeight=Dimension(magnitude=50, unit=Unit.PT),
                    tableRowProperties=TableRowProperties(minRowHeight=Dimension(magnitude=50, unit=Unit.PT)),
                )
            )

        table = Table(rows=5, columns=2, tableColumns=table_columns, tableRows=table_rows)
        table_element = TableElement(
            objectId="test-table-id",
            size=Size(width=Dimension(magnitude=200, unit=Unit.PT), height=Dimension(magnitude=250, unit=Unit.PT)),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Even though we iterate from last row, we should find the first row's styles
        # GFM format with header separator
        markdown = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n| 5 | 6 |\n| 7 | 8 |"

        requests = table_element.content_update_requests(
            markdown_elem=markdown,
            font_scale_factor=0.5,
        )

        # Find UpdateTextStyleRequest
        style_requests = [r for r in requests if isinstance(r, UpdateTextStyleRequest)]

        # Should have style requests (extracted from row 0)
        assert len(style_requests) > 0

        # All font sizes should be scaled: 12pt * 0.5 = 6pt
        for req in style_requests:
            if req.style and req.style.fontSize:
                assert req.style.fontSize.magnitude == 6

    def test_no_template_styles_no_scaling(self):
        """Test that when no template_styles and cell has no styles, no scaling is applied."""
        table_element = self.create_test_table_with_styled_rows(styled_rows=0, empty_rows=3, columns=2)

        # Write to empty cell without template styles
        requests = table_element.write_text_to_cell_requests(
            text="Test",
            location=TableCellLocation(rowIndex=0, columnIndex=0),
            font_scale_factor=0.5,
            template_styles=None,  # No template styles
        )

        # Should still generate requests (for text insertion), just no styled ones
        assert len(requests) > 0
