from pydantic import Field, field_validator

from gslides_api.page.base import BasePage, PageType
from gslides_api.domain import GSlidesBaseModel


class NotesProperties(GSlidesBaseModel):
    """Represents properties of notes."""

    speakerNotesObjectId: str


class Notes(BasePage):
    """Represents a notes page in a presentation."""

    notesProperties: NotesProperties
    pageType: PageType = Field(default=PageType.NOTES, description="The type of page", exclude=True)

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.NOTES
