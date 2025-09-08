from typing import Optional

from pydantic import Field

from gslides_api.domain.domain import GSlidesBaseModel
from gslides_api.page.notes import Notes
from gslides_api.request.parent import GSlidesAPIRequest


class SlideProperties(GSlidesBaseModel):
    """Represents properties of a slide."""

    layoutObjectId: Optional[str] = None
    masterObjectId: Optional[str] = None
    notesPage: Notes = None
    isSkipped: Optional[bool] = None


class UpdateSlidePropertiesRequest(GSlidesAPIRequest):
    """Updates the properties of a Slide.

    This request updates the slide properties for the specified slide.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#updateslidepropertiesrequest
    """

    objectId: str = Field(description="The object ID of the slide to update")
    slideProperties: SlideProperties = Field(description="The slide properties to update")
    fields: str = Field(
        description="The fields that should be updated. At least one field must be specified. The root 'slideProperties' is implied and should not be specified. A single '*' can be used as short-hand for listing every field."
    )
