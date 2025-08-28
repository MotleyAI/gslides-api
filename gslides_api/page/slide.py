import logging
import uuid
from typing import Optional

from gslides_api.client import GoogleAPIClient, api_client
from gslides_api.domain import LayoutReference, ThumbnailProperties, ThumbnailSize, Transform
from gslides_api.element.shape import ShapeElement
from gslides_api.page.base import BasePage, ElementKind, PageType
from gslides_api.page.slide_properties import SlideProperties
from gslides_api.text import Shape, ShapeProperties
from gslides_api.request.domain import Range, RangeType
from gslides_api.request.request import (
    CreateSlideRequest,
    InsertTextRequest,
    UpdatePagePropertiesRequest,
    UpdateSlidePropertiesRequest,
    UpdateSlidesPositionRequest,
    DeleteTextRequest,
)
from gslides_api.response import ImageThumbnail
from gslides_api.utils import dict_to_dot_separated_field_list
from pydantic import Field, field_validator

logger = logging.getLogger(__name__)


class Slide(BasePage):
    """Represents a slide page in a presentation."""

    slideProperties: SlideProperties
    pageType: PageType = Field(default=PageType.SLIDE, description="The type of page", exclude=True)

    @field_validator("pageType")
    @classmethod
    def validate_page_type(cls, v):
        return PageType.SLIDE

    def duplicate(self, api_client: Optional[GoogleAPIClient] = None) -> "Page":
        """
        Duplicates the slide in the same presentation.

        :return:
        """
        assert (
            self.presentation_id is not None
        ), "self.presentation_id must be set when calling duplicate()"
        client = api_client or globals()["api_client"]

        id_map = {self.objectId: uuid.uuid4().hex}  # to avoid force flushing
        for e in self.page_elements_flat:
            id_map[e.objectId] = uuid.uuid4().hex

        new_id = client.duplicate_object(
            object_id=self.objectId, presentation_id=self.presentation_id, id_map=id_map
        )
        new_slide = self.model_copy(deep=True)
        new_slide.objectId = new_id
        for e in new_slide.page_elements_flat:
            e.objectId = id_map[e.objectId]
            e.slide_id = new_slide.objectId

        return new_slide

    def delete(self, api_client: Optional[GoogleAPIClient] = None) -> None:
        assert (
            self.presentation_id is not None
        ), "self.presentation_id must be set when calling delete()"

        client = api_client or globals()["api_client"]
        return client.delete_object(self.objectId, self.presentation_id)

    def move(self, insertion_index: int, api_client: Optional[GoogleAPIClient] = None) -> None:
        """
        Move the slide to a new position in the presentation.

        Args:
            insertion_index: The index to insert the slide at.
        """
        client = api_client or globals()["api_client"]
        request = UpdateSlidesPositionRequest(
            slideObjectIds=[self.objectId], insertionIndex=insertion_index
        )
        client.batch_update([request], self.presentation_id)

    def write_copy(
        self,
        insertion_index: Optional[int] = None,
        presentation_id: Optional[str] = None,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "BasePage":
        """Write the slide to a Google Slides presentation.

        Args:
            presentation_id: The ID of the presentation to write to.
            insertion_index: The index to insert the slide at. If not provided, the slide will be added at the end.
        """
        client = api_client or globals()["api_client"]
        presentation_id = presentation_id or self.presentation_id

        # This method is primarily for slides, so we need to check if we have slide properties
        if not hasattr(self, "slideProperties") or self.slideProperties is None:
            raise ValueError("write_copy is only supported for slide pages")

        new_slide = self.create_blank(
            presentation_id,
            insertion_index,
            slide_layout_reference=LayoutReference(layoutId=self.slideProperties.layoutObjectId),
            api_client=api_client,
        )
        slide_id = new_slide.objectId

        # Set the page properties
        try:
            # TODO: this raises an InternalError sometimes, need to debug
            page_properties = self.pageProperties.to_api_format()
            request = UpdatePagePropertiesRequest(
                objectId=slide_id,
                pageProperties=page_properties,
                fields=",".join(dict_to_dot_separated_field_list(page_properties)),
            )
            client.batch_update([request], presentation_id)
        except Exception as e:
            logger.error(f"Error writing page properties: {e}")

        # Set the slid properties that hadn't been set when creating the slide
        slide_properties = self.slideProperties.to_api_format()
        # Not clear with which call this can be set, but updateSlideProperties rejects it
        slide_properties.pop("masterObjectId", None)
        # This has already been set when creating the slide
        slide_properties.pop("layoutObjectId", None)
        request = UpdateSlidePropertiesRequest(
            objectId=slide_id,
            slideProperties=slide_properties,
            fields=",".join(dict_to_dot_separated_field_list(slide_properties)),
        )
        client.batch_update([request], presentation_id)

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

        return self.from_ids(presentation_id, slide_id, api_client=api_client)

    @classmethod
    def create_blank(
        cls,
        presentation_id: str,
        insertion_index: Optional[int] = None,
        slide_layout_reference: Optional[LayoutReference] = None,
        layoout_placeholder_id_mapping: Optional[dict] = None,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "BasePage":
        """Create a blank slide in a Google Slides presentation.

        Args:
            presentation_id: The ID of the presentation to create the slide in.
            insertion_index: The index to insert the slide at. If not provided, the slide will be added at the end.
            slide_layout_reference: The layout reference to use for the slide.
            layoout_placeholder_id_mapping: The mapping of placeholder IDs to use for the slide.
        """

        client = api_client or globals()["api_client"]
        request = CreateSlideRequest(
            insertionIndex=insertion_index, slideLayoutReference=slide_layout_reference
        )
        out = client.batch_update([request], presentation_id)
        new_slide_id = out["replies"][0]["createSlide"]["objectId"]

        return cls.from_ids(presentation_id, new_slide_id, api_client=api_client)

    @property
    def speaker_notes(self) -> ShapeElement:
        id = self.slideProperties.notesPage.notesProperties.speakerNotesObjectId

        for e in self.slideProperties.notesPage.pageElements:
            if e.objectId == id:
                return e

        # The element must not have been created yet
        return ShapeElement(
            objectId=id,
            slide_id=self.objectId,
            presentation_id=self.presentation_id,
            transform=Transform(),
            shape=Shape(shapeProperties=ShapeProperties()),
        )

    # TODO: do we ever need this?
    def create_speaker_notes(
        self, text: str | None = None, api_client: Optional[GoogleAPIClient] = None
    ) -> None:
        # This assumes the speaker notes don't exist yet
        # Apparently even if the notes element doesn't exist, the API creates it upon the first insertTextRequest
        api_client = api_client or globals()["api_client"]

        notes_id = self.slideProperties.notesPage.notesProperties.speakerNotesObjectId

        req = InsertTextRequest(
            objectId=notes_id,
            text=text if text is not None else "...",
            insertionIndex=0,
        )

        del_req = DeleteTextRequest(objectId=notes_id, textRange=Range(type=RangeType.ALL))

        # Use the global api_client for this internal operation
        api_client.batch_update([req, del_req], self.presentation_id)

    def sync_from_cloud(self):
        new_state = Slide.from_ids(self.presentation_id, self.objectId, api_client=api_client)
        self.__dict__ = new_state.__dict__

    def thumbnail(
        self, size: Optional[ThumbnailSize] = None, api_client: Optional[GoogleAPIClient] = None
    ) -> ImageThumbnail:
        client = api_client or globals()["api_client"]
        props = ThumbnailProperties(thumbnailSize=size)
        return client.slide_thumbnail(self.presentation_id, self.objectId, props)
