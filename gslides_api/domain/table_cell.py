from typing import Optional

from gslides_api.domain.domain import GSlidesBaseModel


class TableCellLocation(GSlidesBaseModel):
    """Represents the location of a table cell."""

    rowIndex: Optional[int] = None
    columnIndex: Optional[int] = None
