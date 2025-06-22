from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api.domain import BulletGlyphPreset, GSlidesBaseModel
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

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert the request to a format suitable for the Google Slides API."""
        return [{"createParagraphBullets": self.to_api_format()}]
