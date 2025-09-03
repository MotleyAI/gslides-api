from typing import List, Optional

from pydantic import Field, field_validator

from gslides_api.client import GoogleAPIClient
from gslides_api.domain import Table
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import TableElement as MarkdownTableElement
from gslides_api.request.request import GSlidesAPIRequest
from gslides_api.request.table import CreateTableRequest


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

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        request = CreateTableRequest(
            elementProperties=element_properties.to_api_format(),
            rows=self.table.rows,
            columns=self.table.columns,
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to an update request for the Google Slides API."""
        # Tables don't have specific properties to update beyond base properties
        return self.alt_text_update_request(element_id)

    def extract_table_data(self) -> TableData:
        """Extract table data from Google Slides Table structure into simple TableData format."""
        if not self.table.tableRows or not self.table.rows or not self.table.columns:
            raise ValueError("Table has no data to extract")

        headers = []
        rows = []

        # Extract data from tableRows
        for row_idx, table_row in enumerate(self.table.tableRows):
            row_cells = []

            # Each table row contains tableCells
            if "tableCells" in table_row:
                for cell in table_row["tableCells"]:
                    # Extract text content from cell
                    cell_text = ""
                    if "text" in cell and "textElements" in cell["text"]:
                        text_parts = []
                        for text_element in cell["text"]["textElements"]:
                            if (
                                "textRun" in text_element
                                and "content" in text_element["textRun"]
                            ):
                                text_parts.append(text_element["textRun"]["content"])
                        cell_text = "".join(text_parts).strip()
                    row_cells.append(cell_text)

            # First row is typically headers
            if row_idx == 0:
                headers = row_cells
            else:
                # Pad row with empty strings if it's shorter than headers
                while len(row_cells) < len(headers):
                    row_cells.append("")
                rows.append(row_cells[: len(headers)])  # Trim if longer than headers

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
                self.transform.to_api_format()
                if hasattr(self.transform, "to_api_format")
                else None
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

    @classmethod
    def from_markdown_element(
        cls,
        markdown_elem: MarkdownTableElement,
        parent_id: str,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "TableElement":
        """Create TableElement from MarkdownTableElement with preserved metadata."""

        # Extract metadata
        metadata = markdown_elem.metadata or {}
        object_id = metadata.get("objectId")

        # Get table data
        table_data = markdown_elem.content

        # Create the Table domain object
        table = Table(
            rows=metadata.get("rows", len(table_data.rows) + 1),  # +1 for header
            columns=metadata.get("columns", len(table_data.headers)),
        )

        # Restore full table structure if available in metadata
        if "tableRows" in metadata:
            table.tableRows = metadata["tableRows"]
        if "tableColumns" in metadata:
            table.tableColumns = metadata["tableColumns"]
        if "horizontalBorderRows" in metadata:
            table.horizontalBorderRows = metadata["horizontalBorderRows"]
        if "verticalBorderRows" in metadata:
            table.verticalBorderRows = metadata["verticalBorderRows"]

        # Create element properties from metadata
        from gslides_api.domain import PageElementProperties

        element_props = PageElementProperties(pageObjectId=parent_id)

        # Restore size if available, otherwise provide default
        if "size" in metadata:
            size_data = metadata["size"]
            from gslides_api.domain import Dimension, Size, Unit

            element_props.size = Size(
                width=Dimension(
                    magnitude=size_data["width"], unit=Unit(size_data["unit"])
                ),
                height=Dimension(
                    magnitude=size_data["height"], unit=Unit(size_data["unit"])
                ),
            )
        else:
            # Provide default size for tables
            from gslides_api.domain import Dimension, Size, Unit

            # Calculate size based on table dimensions
            num_rows = len(table_data.rows) + 1  # +1 for header
            num_cols = len(table_data.headers)

            # Basic sizing: 100pt per column, 30pt per row
            default_width = max(300, num_cols * 100)
            default_height = max(150, num_rows * 30)

            element_props.size = Size(
                width=Dimension(magnitude=default_width, unit=Unit.PT),
                height=Dimension(magnitude=default_height, unit=Unit.PT),
            )

        # Restore transform if available, otherwise create default
        if "transform" in metadata and metadata["transform"]:
            from gslides_api.domain import Transform

            transform_data = metadata["transform"]
            element_props.transform = Transform(**transform_data)
        else:
            # Create a default identity transform
            from gslides_api.domain import Transform

            element_props.transform = Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            )

        # Create the table element
        table_element = cls(
            objectId=object_id
            or "table_" + str(hash(str(markdown_elem.content.headers)))[:8],
            size=element_props.size,
            transform=element_props.transform,
            title=metadata.get("title"),
            description=metadata.get("description"),
            table=table,
            slide_id=parent_id,
            presentation_id="",  # Will need to be set by caller
        )

        # Store the markdown table data for round-trip conversion
        table_element._markdown_table_data = table_data

        return table_element
