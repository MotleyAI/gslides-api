from typing import Optional

from gslides_api.page.notes import Notes
from gslides_api.domain import GSlidesBaseModel


class SlideProperties(GSlidesBaseModel):
    """Represents properties of a slide."""

    layoutObjectId: Optional[str] = None
    masterObjectId: Optional[str] = None
    notesPage: Notes = None
    isSkipped: Optional[bool] = None
