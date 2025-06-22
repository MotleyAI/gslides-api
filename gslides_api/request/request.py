from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api.domain import BulletGlyphPreset, GSlidesBaseModel, TextStyle
from gslides_api.request.domain import Range, TableCellLocation


class GslidesAPIRequest(GSlidesBaseModel):
    """Base class for all requests to the Google Slides API."""

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to the format expected by the Google Slides API."""
        request_name = self.__class__.__name__.replace("Request", "")
        # make first letter lowercase
        request_name = request_name[0].lower() + request_name[1:]

        return [{request_name: self.to_api_format()}]


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


class InsertTextRequest(GslidesAPIRequest):
    """Inserts text into a shape or table cell.

    This request inserts text at the specified insertion index within a shape or table cell.
    The text is inserted at the given index, and all existing text at and after that index
    is shifted to accommodate the new text.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#inserttextrequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table containing the text to insert into"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The optional table cell location if the text is to be inserted into a table cell. If present, the objectId must refer to a table.",
    )
    text: str = Field(description="The text to insert")
    insertionIndex: Optional[int] = Field(
        default=None,
        description="The index where the text will be inserted, in Unicode code units. Text is inserted before the character currently at this index. An insertion index of 0 will insert the text at the beginning of the text.",
    )


class UpdateTextStyleRequest(GslidesAPIRequest):
    """Updates the styling of text within a Shape or Table.

    This request updates the text style for the specified range of text within a shape or table cell.
    The style changes are applied to all text elements that overlap with the given text range.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updatetextstylerequest
    """

    objectId: str = Field(
        description="The object ID of the shape or table with the text to be styled"
    )
    cellLocation: Optional[TableCellLocation] = Field(
        default=None,
        description="The location of the cell in the table containing the text to style. If objectId refers to a table, cellLocation must have a value. Otherwise, it must not.",
    )
    style: TextStyle = Field(
        description="The style(s) to set on the text. If the value for a particular style matches that of the parent, that style will be set to inherit."
    )
    textRange: Range = Field(
        description="The range of text to style. The range may be extended to include adjacent newlines. If the range fully contains a paragraph belonging to a list, the paragraph's bullet is also updated with the matching text style."
    )
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'style' is implied and should not be specified. A single '*' can be used as short-hand for listing every field. For example, to update the text style to bold, set fields to 'bold'."
    )
