from typing import Optional, Dict, Any, List
from enum import Enum
from pydantic import Field, model_validator

from gslides_api.domain import GSlidesBaseModel, BulletGlyphPreset


class GslidesAPIRequest(GSlidesBaseModel):
    """Base class for all requests to the Google Slides API."""

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to the format expected by the Google Slides API."""
        request_name = self.__class__.__name__.replace("Request", "")
        # make first letter lowercase
        request_name = request_name[0].lower() + request_name[1:]

        return [{request_name: self.to_api_format()}]


class RangeType(Enum):
    """Enumeration of possible range types for text selection."""

    ALL = "ALL"
    """Selects all the text in the target shape or table cell."""

    FIXED_RANGE = "FIXED_RANGE"
    """Selects a specific range of text using start and end indexes."""

    FROM_START_INDEX = "FROM_START_INDEX"
    """Selects text from a start index to the end of the text."""


class Range(GSlidesBaseModel):
    """Represents a range of text within a shape or table cell.

    The range can be specified in different ways:
    - ALL: Select all text
    - FIXED_RANGE: Select text between startIndex and endIndex
    - FROM_START_INDEX: Select text from startIndex to the end
    """

    type: RangeType = Field(description="The type of range selection")
    startIndex: Optional[int] = Field(
        default=None, description="The starting index of the range (0-based)"
    )
    endIndex: Optional[int] = Field(
        default=None, description="The ending index of the range (exclusive)"
    )

    @model_validator(mode="after")
    def validate_range_consistency(self) -> "Range":
        """Validate that the range parameters are consistent with the type."""
        if self.type == RangeType.ALL:
            if self.startIndex is not None or self.endIndex is not None:
                raise ValueError("startIndex and endIndex must be None when type is ALL")
        elif self.type == RangeType.FIXED_RANGE:
            if self.startIndex is None or self.endIndex is None:
                raise ValueError(
                    "Both startIndex and endIndex must be provided when type is FIXED_RANGE"
                )
            if self.startIndex >= self.endIndex:
                raise ValueError("startIndex must be less than endIndex")
        elif self.type == RangeType.FROM_START_INDEX:
            if self.startIndex is None:
                raise ValueError("startIndex must be provided when type is FROM_START_INDEX")
            if self.endIndex is not None:
                raise ValueError("endIndex must be None when type is FROM_START_INDEX")
        return self


class TableCellLocation(GSlidesBaseModel):
    """Represents the location of a cell in a table.

    Used to specify which table cell contains the text to be modified
    when the target object is a table rather than a shape.
    """

    rowIndex: int = Field(description="The 0-based row index of the table cell")
    columnIndex: int = Field(description="The 0-based column index of the table cell")

    @model_validator(mode="after")
    def validate_indexes(self) -> "TableCellLocation":
        """Validate that row and column indexes are non-negative."""
        if self.rowIndex < 0:
            raise ValueError("rowIndex must be non-negative")
        if self.columnIndex < 0:
            raise ValueError("columnIndex must be non-negative")
        return self


class CreateParagraphBulletsRequest(GslidesAPIRequest):
    """Creates bullets for paragraphs in a shape or table cell.

    This request converts plain paragraphs into bulleted lists using a specified
    bullet preset pattern. The bullets are applied to all paragraphs that overlap
    with the given text range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createparagraphbulletsrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to add bullets to"
    )
    textRange: Range = Field(
        description="The range of text to add bullets to, based on TextElement indexes"
    )
    bulletPreset: Optional[BulletGlyphPreset] = Field(
        default=None, description="The kinds of bullet glyphs to be used"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text to be modified is in a table cell. If present, the objectId must refer to a table.",
    )

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert the request to a format suitable for the Google Slides API."""
        return [{"createParagraphBullets": self.to_api_format()}]
