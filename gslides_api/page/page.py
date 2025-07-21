import logging
from typing import Annotated, Any, Optional, Union

from pydantic import Discriminator, Field, Tag, field_validator

from gslides_api.page.slide import Slide
from gslides_api.page.notes import Notes
from gslides_api.page.slide_properties import SlideProperties
from gslides_api.page.base import BasePage, PageType
from gslides_api.domain import GSlidesBaseModel

logger = logging.getLogger(__name__)


def page_discriminator(v: Any) -> str:
    """Discriminator function to determine which Page subclass to use based on which properties field is present."""
    if isinstance(v, dict):
        if v.get("slideProperties") is not None:
            return "slide"
        elif v.get("layoutProperties") is not None:
            return "layout"
        elif v.get("notesProperties") is not None:
            return "notes"
        elif v.get("masterProperties") is not None:
            return "master"
        # Handle notes master case - it has pageType but no specific properties
        elif v.get("pageType") == "NOTES_MASTER":
            return "notes_master"
    else:
        # Handle model instances
        if hasattr(v, "slideProperties") and v.slideProperties is not None:
            return "slide"
        elif hasattr(v, "layoutProperties") and v.layoutProperties is not None:
            return "layout"
        elif hasattr(v, "notesProperties") and v.notesProperties is not None:
            return "notes"
        elif hasattr(v, "masterProperties") and v.masterProperties is not None:
            return "master"
        elif hasattr(v, "pageType") and v.pageType == PageType.NOTES_MASTER:
            return "notes_master"

    # If no discriminator found, raise an error
    raise ValueError("Cannot determine page type - no valid properties found")


class LayoutProperties(GSlidesBaseModel):
    """Represents properties of a layout."""

    masterObjectId: Optional[str] = None
    name: Optional[str] = None
    displayName: Optional[str] = None


class Layout(BasePage):
    """Represents a layout page in a presentation."""

    layoutProperties: LayoutProperties
    pageType: PageType = Field(
        default=PageType.LAYOUT, description="The type of page", exclude=True
    )

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.LAYOUT


class MasterProperties(GSlidesBaseModel):
    """Represents properties of a master slide."""

    displayName: Optional[str] = None


class Master(BasePage):
    """Represents a master page in a presentation."""

    masterProperties: MasterProperties
    pageType: PageType = Field(
        default=PageType.MASTER, description="The type of page", exclude=True
    )

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.MASTER


class NotesMaster(BasePage):
    """Represents a notes master page in a presentation."""

    pageType: PageType = Field(
        default=PageType.NOTES_MASTER, description="The type of page", exclude=True
    )

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.NOTES_MASTER


Page = Annotated[
    Union[
        Annotated[Slide, Tag("slide")],
        Annotated[Layout, Tag("layout")],
        Annotated[Notes, Tag("notes")],
        Annotated[Master, Tag("master")],
        Annotated[NotesMaster, Tag("notes_master")],
    ],
    Discriminator(page_discriminator),
]

# Rebuild models to resolve forward references
SlideProperties.model_rebuild()
Slide.model_rebuild()
Layout.model_rebuild()
Notes.model_rebuild()
Master.model_rebuild()
NotesMaster.model_rebuild()
