from typing import List, Optional, Tuple
import uuid

from pydantic import Field, field_validator

from gslides_api.domain import OutputUnit
from gslides_api.element.text_container import TextContainer
from gslides_api.request.domain import TableCellLocation
from gslides_api.table import Table
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import MarkdownTableElement as MarkdownTableElement
from gslides_api.request.request import GSlidesAPIRequest, UpdatePageElementAltTextRequest
from gslides_api.request.table import CreateTableRequest


class TableElement(TextContainer):
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

    def create_request(
        self, parent_id: str, object_id: Optional[str] = None
    ) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to a create request for the Google Slides API."""
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
        """Convert a TableElement to an update request for the Google Slides API."""
        requests = self.alt_text_update_request(element_id)
        for row in self.table.tableRows:
            for cell in row.tableCells:
                requests += cell.text.to_requests(element_id, location=cell.location)
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
    def markdown_element_to_requests(
        cls,
        markdown_elem: MarkdownTableElement,
        parent_id: str,
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
        from gslides_api.domain import Dimension, Size, Unit, Transform

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
            slide_id=parent_id,
            presentation_id="",
        )

        # Start with table creation request
        requests = temp_table_element.create_request(parent_id, element_id)

        requests.append(
            UpdatePageElementAltTextRequest(
                objectId=element_id, title=markdown_elem.name, description=description
            )
        )

        # Generate text requests for each cell
        from gslides_api.markdown.from_markdown import markdown_to_text_elements

        # Process header row first (row 0)
        for col_idx, header_content in enumerate(table_data.headers):
            if header_content.strip():
                cell_location = TableCellLocation(rowIndex=0, columnIndex=col_idx)
                cell_requests = markdown_to_text_elements(header_content.strip())

                # Set objectId and cellLocation for each request
                for request in cell_requests:
                    request.objectId = element_id
                    if hasattr(request, "cellLocation"):
                        request.cellLocation = cell_location

                requests.extend(cell_requests)

        # Process data rows (row 1+)
        for row_idx, row_data in enumerate(table_data.rows):
            for col_idx, cell_content in enumerate(row_data):
                if cell_content.strip():
                    cell_location = TableCellLocation(rowIndex=row_idx + 1, columnIndex=col_idx)
                    cell_requests = markdown_to_text_elements(cell_content.strip())

                    # Set objectId and cellLocation for each request
                    for request in cell_requests:
                        request.objectId = element_id
                        if hasattr(request, "cellLocation"):
                            request.cellLocation = cell_location

                    requests.extend(cell_requests)

        return requests
