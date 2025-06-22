from typing import Optional, List, Union, ForwardRef, Dict, Any, Annotated

import logging

from pydantic import Field, model_validator, Discriminator, Tag, field_validator

from gslides_api.element.base import ElementKind

from gslides_api.domain import (
    GSlidesBaseModel,
    MasterProperties,
    NotesProperties,
    PageType,
    LayoutReference,
    PageBackgroundFill,
    ColorScheme,
)

# Import PageElement and ElementKind directly to avoid circular imports
from gslides_api.element.element import PageElement
from gslides_api.execute import api_client
from gslides_api.utils import dict_to_dot_separated_field_list

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


class SlideProperties(GSlidesBaseModel):
    """Represents properties of a slide."""

    layoutObjectId: Optional[str] = None
    masterObjectId: Optional[str] = None
    notesPage: Optional[ForwardRef("BasePage")] = None
    isSkipped: Optional[bool] = None


class LayoutProperties(GSlidesBaseModel):
    """Represents properties of a layout."""

    masterObjectId: Optional[str] = None
    name: Optional[str] = None
    displayName: Optional[str] = None


# https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages#pageproperties
# The page will inherit properties from the parent page.
# Depending on the page type the hierarchy is defined in either SlideProperties or LayoutProperties.


class SlidePageProperties(SlideProperties):
    """Represents properties of a page."""

    pageBackgroundFill: Optional[PageBackgroundFill] = None
    colorScheme: Optional[ColorScheme] = None


class LayoutPageProperties(LayoutProperties):
    """Represents properties of a page."""

    pageBackgroundFill: Optional[PageBackgroundFill] = None
    colorScheme: Optional[ColorScheme] = None


class PageProperties(GSlidesBaseModel):
    """Represents properties of a page."""

    pageBackgroundFill: Optional[PageBackgroundFill] = None
    colorScheme: Optional[ColorScheme] = None


class BasePage(GSlidesBaseModel):
    """Base class for all page types in a presentation."""

    objectId: Optional[str] = None
    pageElements: Optional[List[PageElement]] = (
        None  # Make optional to preserve original JSON exactly
    )
    revisionId: Optional[str] = None
    pageProperties: Optional[Union[SlidePageProperties, LayoutPageProperties]] = None
    pageType: Optional[PageType] = Field(default=None, exclude=True)

    # Store the presentation ID for reference but exclude from model_dump
    presentation_id: Optional[str] = Field(default=None, exclude=True)

    def _propagate_presentation_id(self, presentation_id: Optional[str] = None) -> None:
        """Helper method to set presentation_id on all pageElements."""
        target_id = presentation_id if presentation_id is not None else self.presentation_id
        if target_id is not None and self.pageElements is not None:
            for element in self.pageElements:
                element.presentation_id = target_id

    @model_validator(mode="after")
    def set_presentation_id_on_elements(self) -> "BasePage":
        """Automatically set presentation_id on all pageElements after model creation."""
        self._propagate_presentation_id()
        return self

    def __setattr__(self, name: str, value) -> None:
        """Override setattr to propagate presentation_id when it's set directly."""
        super().__setattr__(name, value)
        # If presentation_id was just set, propagate it to pageElements
        if name == "presentation_id" and hasattr(self, "pageElements"):
            self._propagate_presentation_id(value)

    @classmethod
    def create_blank(
        cls,
        presentation_id: str,
        insertion_index: Optional[int] = None,
        slide_layout_reference: Optional[LayoutReference] = None,
        layoout_placeholder_id_mapping: Optional[dict] = None,
    ) -> "BasePage":
        """Create a blank slide in a Google Slides presentation.

        Args:
            presentation_id: The ID of the presentation to create the slide in.
            insertion_index: The index to insert the slide at. If not provided, the slide will be added at the end.
            slide_layout_reference: The layout reference to use for the slide.
            layoout_placeholder_id_mapping: The mapping of placeholder IDs to use for the slide.
        """

        # https://developers.google.com/slides/api/reference/rest/v1/presentations/request#CreateSlideRequest
        base = {} if insertion_index is None else {"insertionIndex": insertion_index}
        if slide_layout_reference is not None:
            base["slideLayoutReference"] = slide_layout_reference.to_api_format()

        out = api_client.batch_update([{"createSlide": base}], presentation_id)
        new_slide_id = out["replies"][0]["createSlide"]["objectId"]

        return cls.from_ids(presentation_id, new_slide_id)

    @classmethod
    def from_ids(cls, presentation_id: str, slide_id: str) -> "BasePage":
        # To avoid circular imports
        json = api_client.get_slide_json(presentation_id, slide_id)
        new_slide = cls.model_validate(json)
        new_slide.presentation_id = presentation_id
        return new_slide

    def write_copy(
        self,
        insertion_index: Optional[int] = None,
        presentation_id: Optional[str] = None,
    ) -> "BasePage":
        """Write the slide to a Google Slides presentation.

        Args:
            presentation_id: The ID of the presentation to write to.
            insertion_index: The index to insert the slide at. If not provided, the slide will be added at the end.
        """
        presentation_id = presentation_id or self.presentation_id

        # This method is primarily for slides, so we need to check if we have slide properties
        if not hasattr(self, "slideProperties") or self.slideProperties is None:
            raise ValueError("write_copy is only supported for slide pages")

        new_slide = self.create_blank(
            presentation_id,
            insertion_index,
            slide_layout_reference=LayoutReference(layoutId=self.slideProperties.layoutObjectId),
        )
        slide_id = new_slide.objectId

        # Set the page properties
        try:
            # TODO: this raises an InternalError sometimes, need to debug
            page_properties = self.pageProperties.to_api_format()
            request = [
                {
                    "updatePageProperties": {
                        "objectId": slide_id,
                        "pageProperties": page_properties,
                        "fields": ",".join(dict_to_dot_separated_field_list(page_properties)),
                    }
                }
            ]
            api_client.batch_update(request, presentation_id)
        except Exception as e:
            logger.error(f"Error writing page properties: {e}")

        # Set the slid properties that hadn't been set when creating the slide
        slide_properties = self.slideProperties.to_api_format()
        # Not clear with which call this can be set, but updateSlideProperties rejects it
        slide_properties.pop("masterObjectId", None)
        # This has already been set when creating the slide
        slide_properties.pop("layoutObjectId", None)
        request = [
            {
                "updateSlideProperties": {
                    "objectId": slide_id,
                    "slideProperties": slide_properties,
                    "fields": ",".join(dict_to_dot_separated_field_list(slide_properties)),
                }
            }
        ]
        api_client.batch_update(request, presentation_id)

        if self.pageElements is not None:
            # Some elements came from layout, some were created manually
            # Let's first match those that came from layout, before creating new ones
            for kind in ElementKind:
                my_elements = self.select_elements(kind)
                layout_elements = new_slide.select_elements(kind)
                for i, element in enumerate(my_elements):
                    if i < len(layout_elements):
                        element_id = layout_elements[i].objectId
                    else:
                        element_id = element.create_copy(slide_id, presentation_id)
                    element.update(presentation_id=presentation_id, element_id=element_id)

        return self.from_ids(presentation_id, slide_id)

    def select_elements(self, kind: ElementKind) -> List[PageElement]:
        if self.pageElements is None:
            return []
        return [e for e in self.pageElements if e.type == kind]

    @property
    def image_elements(self):
        if self.pageElements is None:
            return []
        return [e for e in self.pageElements if e.image is not None]

    def get_element_by_id(self, element_id: str) -> PageElement:
        if self.pageElements is None:
            return None
        return next((e for e in self.pageElements if e.objectId == element_id), None)

    def duplicate(self, id_map: Dict[str, str] = None) -> "Page":
        """
        Duplicates the slide in the same presentation.

        :return:
        """
        assert (
            self.presentation_id is not None
        ), "self.presentation_id must be set when calling duplicate()"
        new_id = api_client.duplicate_object(self.objectId, self.presentation_id, id_map)
        return self.from_ids(self.presentation_id, new_id)

    def delete(self) -> None:
        assert (
            self.presentation_id is not None
        ), "self.presentation_id must be set when calling delete()"

        return api_client.delete_object(self.objectId, self.presentation_id)

    def move(self, insertion_index: int) -> None:
        """
        Move the slide to a new position in the presentation.

        Args:
            insertion_index: The index to insert the slide at.
        """
        request = [
            {
                "updateSlidesPosition": {
                    "slideObjectIds": [self.objectId],
                    "insertionIndex": insertion_index,
                }
            }
        ]
        api_client.batch_update(request, self.presentation_id)


class Slide(BasePage):
    """Represents a slide page in a presentation."""

    slideProperties: SlideProperties
    pageType: PageType = Field(default=PageType.SLIDE, description="The type of page", exclude=True)

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.SLIDE


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


class Notes(BasePage):
    """Represents a notes page in a presentation."""

    notesProperties: NotesProperties
    pageType: PageType = Field(default=PageType.NOTES, description="The type of page", exclude=True)

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.NOTES


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


# Create the discriminated union type
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


# SlidePageProperties.model_rebuild()
# Rebuild all page models to handle forward references
Slide.model_rebuild()
Layout.model_rebuild()
Notes.model_rebuild()
Master.model_rebuild()
NotesMaster.model_rebuild()
