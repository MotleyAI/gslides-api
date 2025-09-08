from enum import Enum
from typing import List, Optional

from gslides_api.client import GoogleAPIClient, api_client as default_api_client
from gslides_api.domain import Dimension, GSlidesBaseModel, SolidFill, DashStyle
from gslides_api.element.text_content import TextContent
from gslides_api.text import TextStyle


class TableCellLocation(GSlidesBaseModel):
    """Represents the location of a table cell."""

    rowIndex: Optional[int] = None
    columnIndex: Optional[int] = None


class ContentAlignment(Enum):
    """Enumeration of possible content alignment values for table cells."""

    TOP = "TOP"
    MIDDLE = "MIDDLE"
    BOTTOM = "BOTTOM"


class TableRowProperties(GSlidesBaseModel):
    """Represents properties of a table row."""

    minRowHeight: Optional[Dimension] = None


class TableColumnProperties(GSlidesBaseModel):
    """Represents properties of a table column."""

    columnWidth: Optional[Dimension] = None


class TableCellBackgroundFill(GSlidesBaseModel):
    """Represents the background fill of a table cell."""

    solidFill: Optional[SolidFill] = None
    propertyState: Optional[str] = None


class TableCellProperties(GSlidesBaseModel):
    """Represents properties of a table cell."""

    tableCellBackgroundFill: Optional[TableCellBackgroundFill] = None
    contentAlignment: Optional[ContentAlignment] = None


class TableCell(GSlidesBaseModel):
    """Represents an individual table cell."""

    location: Optional[TableCellLocation] = None
    rowSpan: Optional[int] = None
    columnSpan: Optional[int] = None
    text: Optional[TextContent] = None
    tableCellProperties: Optional[TableCellProperties] = None

    def read_text(self, as_markdown: bool = True) -> str:
        if self.text is None:
            return ""
        return self.text.read_text(as_markdown=as_markdown)

    def write_text(self, *args, **kwargs):
        raise NotImplementedError("Use TableElement.write_text_to_cell instead.")

    # def write_text(
    #     self,
    #     text: str,
    #     as_markdown: bool = True,
    #     styles: List[TextStyle] | None = None,
    #     overwrite: bool = True,
    #     autoscale: bool = False,
    #     api_client: Optional[GoogleAPIClient] = None,
    # ) -> None:
    #     if autoscale:
    #         raise NotImplementedError(
    #             "Autoscale not implemented for table cells yet, call TableElement.write_text_to_cell instead."
    #         )
    #     if self.text is None:
    #         self.text = TextContent()
    #     requests = self.text.write_text_requests(
    #         text=text,
    #         as_markdown=as_markdown,
    #         styles=styles,
    #         overwrite=overwrite,
    #         autoscale=autoscale,
    #     )
    #     for r in requests:
    #         r.objectId = self.objectId
    #         if hasattr(r, "cellLocation"):
    #             r.cellLocation = self.location
    #     if requests:
    #         client = api_client or default_api_client
    #         return client.batch_update(requests, self.presentation_id)


class TableRow(GSlidesBaseModel):
    """Represents a table row."""

    rowHeight: Optional[Dimension] = None
    tableRowProperties: Optional[TableRowProperties] = None
    tableCells: Optional[List[TableCell]] = None


class TableBorderFill(GSlidesBaseModel):
    """Represents the fill of a table border."""

    solidFill: Optional[SolidFill] = None


class TableBorderProperties(GSlidesBaseModel):
    """Represents properties of a table border."""

    tableBorderFill: Optional[TableBorderFill] = None
    weight: Optional[Dimension] = None
    dashStyle: Optional[DashStyle] = None


class TableBorderCell(GSlidesBaseModel):
    """Represents a table border cell."""

    location: Optional[TableCellLocation] = None
    tableBorderProperties: Optional[TableBorderProperties] = None


class TableBorderRow(GSlidesBaseModel):
    """Represents a table border row."""

    tableBorderCells: Optional[List[TableBorderCell]] = None


class Table(GSlidesBaseModel):
    """Represents a table in a slide."""

    rows: Optional[int] = None
    columns: Optional[int] = None
    tableRows: Optional[List[TableRow]] = None
    tableColumns: Optional[List[TableColumnProperties]] = None
    horizontalBorderRows: Optional[List[TableBorderRow]] = None
    verticalBorderRows: Optional[List[TableBorderRow]] = None
