from enum import Enum
from typing import List, Optional

from gslides_api.domain import Dimension, GSlidesBaseModel, SolidFill, DashStyle
from gslides_api.element.text_container import TextContent


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
