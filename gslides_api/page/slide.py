from typing import Dict, Optional
import logging

from pydantic import Field, field_validator

from gslides_api.page.base import BasePage, ElementKind, PageType
from gslides_api.domain import GSlidesBaseModel, LayoutReference
from gslides_api.execute import api_client
from gslides_api.utils import dict_to_dot_separated_field_list


logger = logging.getLogger(__name__)


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


class SlideProperties(GSlidesBaseModel):
    """Represents properties of a slide."""

    layoutObjectId: Optional[str] = None
    masterObjectId: Optional[str] = None
    notesPage: Notes = None
    isSkipped: Optional[bool] = None


class Slide(BasePage):
    """Represents a slide page in a presentation."""

    slideProperties: SlideProperties
    pageType: PageType = Field(default=PageType.SLIDE, description="The type of page", exclude=True)

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.SLIDE

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

    @property
    def speaker_notes_element_id(self):
        return self.slideProperties.notesPage.notesProperties.speakerNotesObjectId

    def read_speaker_notes(self) -> str | None:
        """Read the speaker notes for the slide."""
        # Apparently the API guarantees this will always be set
        for e in self.notes_page.pageElements:
            if e.objectId == self.speaker_notes_element_id:
                return e.to_markdown()
        return None

    def write_speaker_notes(self, text: str):
        """Write speaker notes for the slide."""
        # make sure notes page exists

        for e in self.notes_page.pageElements:
            if e.objectId == self.speaker_notes_element_id:
                e.write_text(text, as_markdown=False)
                return
