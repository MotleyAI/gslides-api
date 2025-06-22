from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api.request.request import GSlidesAPIRequest


class CreateTableRequest(GSlidesAPIRequest):
    """Creates a new table.

    This request creates a new table on the specified page with the given number of rows and columns.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createtablerequest
    """

    elementProperties: Dict[str, Any] = Field(description="The element properties for the table")
    rows: int = Field(description="Number of rows in the table")
    columns: int = Field(description="Number of columns in the table")


class InsertTableRowsRequest(GSlidesAPIRequest):
    """Inserts rows into a table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#inserttablerowsrequest
    """

    tableObjectId: str = Field(description="The table to insert rows into")
    cellLocation: Optional[Dict[str, Any]] = Field(
        default=None,
        description="The reference table cell location from which rows will be inserted",
    )
    insertBelow: Optional[bool] = Field(
        default=None, description="Whether to insert new rows below the reference cell location"
    )
    number: int = Field(description="The number of rows to be inserted")


class InsertTableColumnsRequest(GSlidesAPIRequest):
    """Inserts columns into a table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#inserttablecolumnsrequest
    """

    tableObjectId: str = Field(description="The table to insert columns into")
    cellLocation: Optional[Dict[str, Any]] = Field(
        default=None,
        description="The reference table cell location from which columns will be inserted",
    )
    insertRight: Optional[bool] = Field(
        default=None,
        description="Whether to insert new columns to the right of the reference cell location",
    )
    number: int = Field(description="The number of columns to be inserted")


class DeleteTableRowRequest(GSlidesAPIRequest):
    """Deletes a row from a table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deletetablerowrequest
    """

    tableObjectId: str = Field(description="The table to delete rows from")
    cellLocation: Dict[str, Any] = Field(
        description="The reference table cell location from which a row will be deleted"
    )


class DeleteTableColumnRequest(GSlidesAPIRequest):
    """Deletes a column from a table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#deletetablecolumnrequest
    """

    tableObjectId: str = Field(description="The table to delete columns from")
    cellLocation: Dict[str, Any] = Field(
        description="The reference table cell location from which a column will be deleted"
    )


class UpdateTableCellPropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a TableCell.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetablecellpropertiesrequest
    """

    objectId: str = Field(description="The object ID of the table")
    tableRange: Dict[str, Any] = Field(
        description="The table range representing the subset of the table to which the updates are applied"
    )
    tableCellProperties: Dict[str, Any] = Field(description="The table cell properties to update")
    fields: str = Field(description="The fields that should be updated")


class UpdateTableBorderPropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of the table borders in a Table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetableborderpropertiesrequest
    """

    objectId: str = Field(description="The object ID of the table")
    tableRange: Optional[Dict[str, Any]] = Field(
        default=None,
        description="The table range representing the subset of the table to which the updates are applied",
    )
    borderPosition: Optional[str] = Field(
        default=None,
        description="The border position in the table range the updates should apply to",
    )
    tableBorderProperties: Dict[str, Any] = Field(
        description="The table border properties to update"
    )
    fields: str = Field(description="The fields that should be updated")


class UpdateTableColumnPropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Table column.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetablecolumnpropertiesrequest
    """

    objectId: str = Field(description="The object ID of the table")
    columnIndices: Optional[List[int]] = Field(
        default=None,
        description="The list of zero-based indices specifying which columns to update",
    )
    tableColumnProperties: Dict[str, Any] = Field(
        description="The table column properties to update"
    )
    fields: str = Field(description="The fields that should be updated")


class UpdateTableRowPropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Table row.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetablerowpropertiesrequest
    """

    objectId: str = Field(description="The object ID of the table")
    rowIndices: Optional[List[int]] = Field(
        default=None, description="The list of zero-based indices specifying which rows to update"
    )
    tableRowProperties: Dict[str, Any] = Field(description="The table row properties to update")
    fields: str = Field(description="The fields that should be updated")


class MergeTableCellsRequest(GSlidesAPIRequest):
    """Merges cells in a Table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#mergetablecellsrequest
    """

    objectId: str = Field(description="The object ID of the table")
    tableRange: Dict[str, Any] = Field(
        description="The table range specifying which cells of the table to merge"
    )


class UnmergeTableCellsRequest(GSlidesAPIRequest):
    """Unmerges cells in a Table.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#unmergetablecellsrequest
    """

    objectId: str = Field(description="The object ID of the table")
    tableRange: Dict[str, Any] = Field(
        description="The table range specifying which cells of the table to unmerge"
    )
