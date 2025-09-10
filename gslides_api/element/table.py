from typing import List, Optional, Sequence, Tuple, Any
import uuid

from pydantic import Field, field_validator
from typeguard import typechecked

from gslides_api.client import GoogleAPIClient, api_client as default_api_client
from gslides_api.domain.domain import Dimension, OutputUnit, Size, Transform
from gslides_api.domain.table_cell import TableCellLocation
from gslides_api.domain.table import Table, TableColumnProperties
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import MarkdownTableElement as MarkdownTableElement
from gslides_api.request.request import UpdatePageElementAltTextRequest
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.table import (
    CreateTableRequest,
    InsertTableRowsRequest,
    DeleteTableRowRequest,
    InsertTableColumnsRequest,
    DeleteTableColumnRequest,
    UpdateTableColumnPropertiesRequest,
)


from gslides_api.domain.text import TextStyle


@typechecked
class TableElement(PageElementBase):
    """Represents a table element on a slide."""

    table: Table
    type: ElementKind = Field(
        default=ElementKind.TABLE, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.TABLE

    def absolute_size(
        self, units: OutputUnit, location: Optional[TableCellLocation] = None
    ) -> Tuple[float, float]:

        if location is None:
            return super().absolute_size(units)
        else:
            # Get cell-specific dimensions
            if location.rowIndex is None or location.columnIndex is None:
                raise ValueError(
                    "Both rowIndex and columnIndex must be specified in TableCellLocation"
                )

            # Get row height from tableRows
            if (
                not self.table.tableRows
                or location.rowIndex >= len(self.table.tableRows)
                or not self.table.tableRows[location.rowIndex].rowHeight
            ):
                raise ValueError(f"Row height not available for row {location.rowIndex}")

            row_height_dim = self.table.tableRows[location.rowIndex].rowHeight

            # Get column width from tableColumns
            if (
                not self.table.tableColumns
                or location.columnIndex >= len(self.table.tableColumns)
                or not self.table.tableColumns[location.columnIndex].columnWidth
            ):
                raise ValueError(f"Column width not available for column {location.columnIndex}")

            column_width_dim = self.table.tableColumns[location.columnIndex].columnWidth

            # Extract EMU values from Dimension objects
            width_emu = column_width_dim.magnitude
            height_emu = row_height_dim.magnitude

            # Apply transform scaling (cells inherit table transforms)
            actual_width_emu = width_emu * self.transform.scaleX
            actual_height_emu = height_emu * self.transform.scaleY

            # Convert from EMUs to requested units
            width_result = self._convert_emu_to_units(actual_width_emu, units)
            height_result = self._convert_emu_to_units(actual_height_emu, units)

            return width_result, height_result

    def _read_text(self, location: Optional[TableCellLocation] = None) -> str | list[list[str]]:
        if location is not None:
            cell = self.table.tableRows[location.rowIndex].tableCells[location.columnIndex]
            if cell.text is not None:
                return cell.text.read_text()
            else:
                return ""
        else:
            out = []
            for row in self.table.tableRows:
                this_row = []
                for cell in row.tableCells:
                    if cell.text is not None:
                        this_row.append(cell.text.read_text() or "")
                    else:
                        this_row.append("")
                out.append(this_row)
            return out

    def __getitem__(self, key):
        """Get a table cell by row/column indices or TableCellLocation.

        Supports multiple access patterns:
        - table[row, col] - tuple syntax (most common)
        - table[location] - TableCellLocation object
        """
        if isinstance(key, TableCellLocation):
            row_idx, col_idx = key.rowIndex, key.columnIndex
        elif isinstance(key, tuple) and len(key) == 2:
            row_idx, col_idx = key
        else:
            raise TypeError("Table indexing requires either (row, col) tuple or TableCellLocation")

        return self.table.tableRows[row_idx].tableCells[col_idx]

    def write_text_to_cell_requests(
        self,
        text: str,
        location: TableCellLocation | Sequence[int],
        as_markdown: bool = True,
        styles: List[TextStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        check_shape: bool = True,
    ) -> List[GSlidesAPIRequest]:
        if isinstance(location, Sequence):
            location = TableCellLocation(rowIndex=location[0], columnIndex=location[1])

        # Validate cell location is within table bounds
        if check_shape and (
            location.rowIndex >= self.table.rows
            or location.columnIndex >= self.table.columns
            or location.rowIndex < 0
            or location.columnIndex < 0
        ):
            raise ValueError(
                f"Cell location ({location.rowIndex}, {location.columnIndex}) "
                f"is outside table bounds ({self.table.rows}, {self.table.columns})"
            )

        # If table structure is populated, use the existing cell's text content
        if (
            self.table.tableRows is not None
            and location.rowIndex < len(self.table.tableRows)
            and self.table.tableRows[location.rowIndex].tableCells is not None
            and location.columnIndex < len(self.table.tableRows[location.rowIndex].tableCells)
        ):
            cell = self[location.rowIndex, location.columnIndex]
            size_inches = self.absolute_size(OutputUnit.IN, location)
            requests = cell.text.write_text_requests(
                text=text,
                as_markdown=as_markdown,
                styles=styles,
                overwrite=overwrite,
                autoscale=autoscale,
                size_inches=size_inches,
            )
        else:
            # Table structure not populated yet (e.g., during creation from markdown)
            # Create a temporary TextContent to generate the requests
            from gslides_api.element.text_content import TextContent

            temp_text_content = TextContent(textElements=[])

            # Calculate size if possible, otherwise use default
            try:
                size_inches = self.absolute_size(OutputUnit.IN, location)
            except (ValueError, AttributeError):
                size_inches = (4.0, 1.0) if autoscale else None

            requests = temp_text_content.write_text_requests(
                text=text,
                as_markdown=as_markdown,
                styles=styles,
                overwrite=overwrite,
                autoscale=autoscale,
                size_inches=size_inches,
            )

        # Set objectId and cellLocation on all requests
        for r in requests:
            r.objectId = self.objectId
            if hasattr(r, "cellLocation"):
                r.cellLocation = location
        return requests

    def create_request(
        self, parent_id: str, object_id: Optional[str] = None
    ) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to a create request for the Google Slides API.
        We can supply an object_id so we know it for other operations,
        even before the request is executed.
        """
        element_properties = self.element_properties(parent_id)
        if object_id is None:
            object_id = f"table_{uuid.uuid4().hex[:8]}"

        request = CreateTableRequest(
            objectId=object_id,
            elementProperties=element_properties,
            rows=self.table.rows,
            columns=self.table.columns,
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to an update request for the Google Slides API.
        Meant for copying one's content to another element
        """
        requests = self.alt_text_update_request(element_id)
        requests += self.content_update_requests(self.to_markdown_element("temp"))
        return requests

    def content_update_requests(
        self, markdown_elem: MarkdownTableElement | str, check_shape: bool = True
    ) -> List[GSlidesAPIRequest]:
        """
        Update the table's content with the provided markdown table.
        :param markdown_elem: if a str, must contain only a valid markdown table
        :return: Request to update the table's content
        """
        requests = []

        if isinstance(markdown_elem, str):
            markdown_elem = MarkdownTableElement.from_markdown("temp", markdown_elem)

        for row in range(markdown_elem.shape[0]):
            for col in range(markdown_elem.shape[1]):
                cell_content = markdown_elem[row, col]
                cell_location = TableCellLocation(rowIndex=row, columnIndex=col)
                requests.extend(
                    self.write_text_to_cell_requests(
                        cell_content.strip(), cell_location, check_shape=check_shape
                    )
                )

        return requests

    def write_text_to_cell(
        self,
        text: str,
        location: TableCellLocation | Sequence[int],
        as_markdown: bool = True,
        styles: List[TextStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> dict[str, Any] | None:
        requests = self.write_text_to_cell_requests(
            text=text,
            location=location,
            as_markdown=as_markdown,
            styles=styles,
            overwrite=overwrite,
            autoscale=autoscale,
        )
        if requests:
            client = api_client or default_api_client
            return client.batch_update(requests, self.presentation_id)

    def delete_text_in_cell_requests(self, location: TableCellLocation) -> List[GSlidesAPIRequest]:
        # Validate cell location is within table bounds
        if (
            location.rowIndex >= self.table.rows
            or location.columnIndex >= self.table.columns
            or location.rowIndex < 0
            or location.columnIndex < 0
        ):
            raise ValueError(
                f"Cell location ({location.rowIndex}, {location.columnIndex}) "
                f"is outside table bounds ({self.table.rows}, {self.table.columns})"
            )

        # If table structure is populated, use the existing cell's text content
        if (
            self.table.tableRows is not None
            and location.rowIndex < len(self.table.tableRows)
            and self.table.tableRows[location.rowIndex].tableCells is not None
            and location.columnIndex < len(self.table.tableRows[location.rowIndex].tableCells)
        ):
            cell = self[location.rowIndex, location.columnIndex]
            requests = cell.text.delete_text_request(self.objectId)
        else:
            # Table structure not populated yet - create generic delete requests
            from gslides_api.element.text_content import TextContent

            temp_text_content = TextContent(textElements=[])
            requests = temp_text_content.delete_text_request(self.objectId)

        for r in requests:
            if hasattr(r, "cellLocation"):
                r.cellLocation = location
        return requests

    def extract_table_data(self) -> TableData:
        """Extract table data from Google Slides Table structure into simple TableData format."""
        if not self.table.tableRows or not self.table.rows or not self.table.columns:
            raise ValueError("Table has no data to extract")

        # Use _read_text() to get all cell text at once
        try:
            all_cell_text = self._read_text()
            if not isinstance(all_cell_text, list) or not all_cell_text:
                raise ValueError("No data extracted from table")

            # First row is typically headers
            headers = [cell.strip() for cell in all_cell_text[0]]

            # Remaining rows are data
            rows = []
            for row_data in all_cell_text[1:]:
                row_cells = [cell.strip() for cell in row_data]
                # Pad row with empty strings if it's shorter than headers
                while len(row_cells) < len(headers):
                    row_cells.append("")
                # Trim if longer than headers
                rows.append(row_cells[: len(headers)])

        except (AttributeError, IndexError):
            raise ValueError("Could not extract table data - table structure may be invalid")

        if not headers:
            raise ValueError("No headers found in table")

        return TableData(headers=headers, rows=rows)

    def to_markdown_element(self, name: str = "Table") -> MarkdownTableElement:
        """Convert TableElement to MarkdownTableElement for round-trip conversion."""

        # Check if we have stored table data from markdown conversion
        if hasattr(self, "_markdown_table_data") and self._markdown_table_data:
            table_data = self._markdown_table_data
        else:
            # Extract table data from Google Slides structure
            try:
                table_data = self.extract_table_data()
            except ValueError as e:
                # If we can't extract data, create an empty table
                table_data = TableData(headers=["Column 1"], rows=[])

        # Store all necessary metadata for perfect reconstruction
        metadata = {
            "objectId": self.objectId,
            "rows": self.table.rows,
            "columns": self.table.columns,
        }

        # Store element properties (position, size, etc.) if available
        if hasattr(self, "size") and self.size:
            metadata["size"] = {
                "width": self.size.width.magnitude,
                "height": self.size.height.magnitude,
                "unit": self.size.width.unit.value,
            }

        if hasattr(self, "transform") and self.transform:
            metadata["transform"] = (
                self.transform.to_api_format() if hasattr(self.transform, "to_api_format") else None
            )

        # Store title and description if available
        if hasattr(self, "title") and self.title:
            metadata["title"] = self.title
        if hasattr(self, "description") and self.description:
            metadata["description"] = self.description

        # Store raw table structure for perfect reconstruction
        if self.table.tableRows:
            metadata["tableRows"] = self.table.tableRows
        if self.table.tableColumns:
            metadata["tableColumns"] = self.table.tableColumns
        if self.table.horizontalBorderRows:
            metadata["horizontalBorderRows"] = self.table.horizontalBorderRows
        if self.table.verticalBorderRows:
            metadata["verticalBorderRows"] = self.table.verticalBorderRows

        return MarkdownTableElement(name=name, content=table_data, metadata=metadata)

    def to_markdown(self) -> str | None:
        """Convert the table's content back to markdown format."""
        return self.to_markdown_element("Table").to_markdown()

    @classmethod
    def create_element_from_markdown_requests(
        cls,
        markdown_elem: MarkdownTableElement,
        slide_id: str,
        description: Optional[str] = None,
        element_id: Optional[str] = None,
    ) -> List[GSlidesAPIRequest]:
        """Convert MarkdownTableElement to a sequence of API requests that will create the table."""

        # Generate object_id if not provided
        if element_id is None:
            element_id = f"table_{uuid.uuid4().hex[:8]}"

        # Get table data
        table_data = markdown_elem.content

        # Calculate table dimensions
        num_rows = len(table_data.rows) + 1  # +1 for header
        num_cols = len(table_data.headers)

        # Create temporary TableElement to generate the structure creation request
        from gslides_api.domain.domain import Unit

        # Basic sizing: 100pt per column, 30pt per row
        default_width = max(300, num_cols * 100)
        default_height = max(150, num_rows * 30)

        temp_table_element = cls(
            objectId=element_id,
            size=Size(
                width=Dimension(magnitude=default_width, unit=Unit.PT),
                height=Dimension(magnitude=default_height, unit=Unit.PT),
            ),
            transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"),
            table=Table(rows=num_rows, columns=num_cols),
            slide_id=slide_id,
            presentation_id="",
        )

        # Start with table creation request
        requests = temp_table_element.create_request(slide_id, element_id)

        requests.append(
            UpdatePageElementAltTextRequest(
                objectId=element_id, title=markdown_elem.name, description=description
            )
        )

        # Generate text requests for each cell
        requests += temp_table_element.content_update_requests(markdown_elem)

        return requests

    def _calculate_proportional_widths_after_deletion(
        self, remaining_column_indices: List[int]
    ) -> List[GSlidesAPIRequest]:
        """Generate requests to proportionally expand remaining columns to maintain total table width.

        Args:
            remaining_column_indices: List of column indices that will remain after deletion

        Returns:
            List of UpdateTableColumnPropertiesRequest to adjust column widths
        """
        requests = []

        # Only proceed if we have column width information
        if not self.table.tableColumns or len(self.table.tableColumns) != self.table.columns:
            return requests

        # Calculate total width of remaining columns
        total_remaining_width = 0
        remaining_widths = []

        for col_idx in remaining_column_indices:
            if (
                col_idx < len(self.table.tableColumns)
                and self.table.tableColumns[col_idx].columnWidth
            ):
                width = self.table.tableColumns[col_idx].columnWidth.magnitude
                remaining_widths.append(width)
                total_remaining_width += width
            else:
                # If we don't have width info for some columns, can't do proportional adjustment
                return requests

        if total_remaining_width == 0:
            return requests

        # Calculate total width of all original columns
        total_original_width = 0
        for col in self.table.tableColumns:
            if col.columnWidth:
                total_original_width += col.columnWidth.magnitude

        if total_original_width == 0:
            return requests

        # Calculate proportional expansion factor
        expansion_factor = total_original_width / total_remaining_width

        # Generate update requests for each remaining column
        for i, col_idx in enumerate(remaining_column_indices):
            new_width = remaining_widths[i] * expansion_factor
            original_col = self.table.tableColumns[col_idx]

            # Create new column properties with adjusted width
            new_column_props = TableColumnProperties(
                columnWidth=Dimension(magnitude=new_width, unit=original_col.columnWidth.unit)
            )

            requests.append(
                UpdateTableColumnPropertiesRequest(
                    objectId=self.objectId,
                    columnIndices=[col_idx],
                    tableColumnProperties=new_column_props,
                    fields="columnWidth",
                )
            )

        return requests

    def _generate_width_preserving_requests_after_addition(
        self, original_columns: int, new_columns: int
    ) -> List[GSlidesAPIRequest]:
        """Generate requests to preserve original column widths and set new columns to rightmost width.

        Args:
            original_columns: Number of columns before addition
            new_columns: Total number of columns after addition

        Returns:
            List of UpdateTableColumnPropertiesRequest to set column widths
        """
        from gslides_api.request.table import UpdateTableColumnPropertiesRequest
        from gslides_api.domain.table import TableColumnProperties

        requests = []

        # Only proceed if we have column width information
        if not self.table.tableColumns or len(self.table.tableColumns) != original_columns:
            return requests

        # Get the width of the rightmost original column
        rightmost_col = self.table.tableColumns[original_columns - 1]
        if not rightmost_col.columnWidth:
            return requests

        rightmost_width = rightmost_col.columnWidth

        # Set all original columns to their current widths (preserving them)
        for col_idx in range(original_columns):
            col = self.table.tableColumns[col_idx]
            if col.columnWidth:
                requests.append(
                    UpdateTableColumnPropertiesRequest(
                        objectId=self.objectId,
                        columnIndices=[col_idx],
                        tableColumnProperties=TableColumnProperties(columnWidth=col.columnWidth),
                        fields="columnWidth",
                    )
                )

        # Set all new columns to the rightmost column width
        for col_idx in range(original_columns, new_columns):
            requests.append(
                UpdateTableColumnPropertiesRequest(
                    objectId=self.objectId,
                    columnIndices=[col_idx],
                    tableColumnProperties=TableColumnProperties(columnWidth=rightmost_width),
                    fields="columnWidth",
                )
            )

        return requests

    def resize_requests(
        self, n_rows: int, n_columns: int, fix_width: bool = True
    ) -> List[GSlidesAPIRequest]:
        """Generate requests to resize the table to the specified dimensions.

        Args:
            n_rows: Target number of rows
            n_columns: Target number of columns
            fix_width: If True (default), maintain constant table width when adding/deleting columns.
                      If False, preserve original column widths when adding columns and allow
                      table width to change when deleting columns.

        Returns:
            List of API requests to resize the table

        Raises:
            ValueError: If target dimensions are less than 1
        """
        if n_rows < 1 or n_columns < 1:
            raise ValueError("Table must have at least 1 row and 1 column")

        requests = []
        current_rows = self.table.rows
        current_columns = self.table.columns

        # Handle row changes
        if n_rows > current_rows:
            rows_to_add = n_rows - current_rows
            requests.append(
                InsertTableRowsRequest(
                    tableObjectId=self.objectId,
                    cellLocation=TableCellLocation(rowIndex=current_rows - 1, columnIndex=0),
                    insertBelow=True,
                    number=rows_to_add,
                )
            )
        elif n_rows < current_rows:
            rows_to_delete = current_rows - n_rows
            for i in range(rows_to_delete):
                row_index = current_rows - 1 - i
                requests.append(
                    DeleteTableRowRequest(
                        tableObjectId=self.objectId,
                        cellLocation={"rowIndex": row_index, "columnIndex": 0},
                    )
                )

        # Handle column changes
        if n_columns > current_columns:
            # Adding columns
            columns_to_add = n_columns - current_columns
            requests.append(
                InsertTableColumnsRequest(
                    tableObjectId=self.objectId,
                    cellLocation=TableCellLocation(rowIndex=0, columnIndex=current_columns - 1),
                    insertRight=True,
                    number=columns_to_add,
                )
            )

            # Add width adjustment requests if fix_width=False
            if not fix_width:
                width_requests = self._generate_width_preserving_requests_after_addition(
                    current_columns, n_columns
                )
                requests.extend(width_requests)

        elif n_columns < current_columns:
            # Deleting columns
            columns_to_delete = current_columns - n_columns

            # Calculate which columns will remain for width adjustment if fix_width=True
            remaining_column_indices = list(range(n_columns)) if fix_width else []

            for i in range(columns_to_delete):
                column_index = current_columns - 1 - i
                requests.append(
                    DeleteTableColumnRequest(
                        tableObjectId=self.objectId,
                        cellLocation=TableCellLocation(rowIndex=0, columnIndex=column_index),
                    )
                )

            # Add width adjustment requests if fix_width=True
            if fix_width and remaining_column_indices:
                width_requests = self._calculate_proportional_widths_after_deletion(
                    remaining_column_indices
                )
                requests.extend(width_requests)

        return requests

    def resize(self, n_rows: int, n_columns: int, fix_width: bool = True, api_client=None) -> None:
        """Resize the table to the specified dimensions.

        Args:
            n_rows: Target number of rows
            n_columns: Target number of columns
            fix_width: If True (default), maintain constant table width when adding/deleting columns.
                      If False, preserve original column widths when adding columns and allow
                      table width to change when deleting columns.
            api_client: Optional GoogleAPIClient instance. If None, uses the default client.

        Raises:
            ValueError: If target dimensions are less than 1 or if no API client is available
        """
        requests = self.resize_requests(n_rows, n_columns, fix_width)

        if not requests:
            # No changes needed
            return

        if api_client is None:

            api_client = default_api_client
            if api_client is None:
                raise ValueError(
                    "No API client available. Please provide an api_client parameter or initialize the default client."
                )

        # Execute the resize requests
        api_client.batch_update(requests, presentation_id=self.presentation_id)

        # Update local table dimensions
        self.table.rows = n_rows
        self.table.columns = n_columns
