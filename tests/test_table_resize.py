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
        requests = table_element.resize_requests(3, 2, fix_width=True)

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
        requests = table_element.resize_requests(3, 2, fix_width=False)

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
        requests_add = table_element.resize_requests(3, 6, fix_width=False)
        requests_delete = table_element.resize_requests(3, 2, fix_width=True)

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
        requests = table_element.resize_requests(5, 6, fix_width=False)

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

        # Mock resize_requests to return some test requests
        with patch.object(TableElement, "resize_requests") as mock_resize_requests:
            mock_requests = [Mock()]
            mock_resize_requests.return_value = mock_requests

            # Test with fix_width=True (fix_height uses default False)
            table_element.resize(3, 6, fix_width=True, api_client=mock_client_instance)

            mock_resize_requests.assert_called_once_with(3, 6, True, False, target_height_emu=None)
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
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            ),
            table=table,
            slide_id="test-slide-id",
            presentation_id="test-presentation-id",
        )

        # Delete 2 columns, keep first 2
        requests = table_element.resize_requests(3, 2, fix_width=True)

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
        requests = table_element.resize_requests(5, 4, fix_height=False)

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
        requests = table_element.resize_requests(5, 4, fix_height=True)

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
        requests = table_element.resize_requests(2, 4, fix_height=False)

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
        requests = table_element.resize_requests(2, 4, fix_height=True)

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
        requests_add = table_element.resize_requests(5, 4, fix_height=True)
        requests_delete = table_element.resize_requests(2, 4, fix_height=True)

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
        requests = table_element.resize_requests(2, 4, fix_height=True)

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
        requests = table_element.resize_requests(5, 6, fix_width=True, fix_height=True)

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

        # Mock resize_requests to verify parameters
        with patch.object(TableElement, "resize_requests") as mock_resize_requests:
            mock_requests = [Mock()]
            mock_resize_requests.return_value = mock_requests

            # Test with both fix_width and fix_height
            table_element.resize(
                5, 6, fix_width=False, fix_height=True, api_client=mock_client_instance
            )

            mock_resize_requests.assert_called_once_with(5, 6, False, True, target_height_emu=None)
            mock_client_instance.batch_update.assert_called_once()

    def test_resize_requests_height_adjustment_row_indices(self):
        """Test that height adjustment requests target the correct row indices."""
        table_element = self.create_test_table(4, 3)

        # Delete 2 rows (keep rows 0, 1)
        requests = table_element.resize_requests(2, 3, fix_height=True)

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

        requests = table_element.resize_requests(
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

        requests = table_element.resize_requests(
            n_rows=6, n_columns=4, target_height_emu=target_height_emu
        )

        # Should NOT have border request when border weight is 0
        border_requests = [r for r in requests if isinstance(r, UpdateTableBorderPropertiesRequest)]
        assert len(border_requests) == 0
