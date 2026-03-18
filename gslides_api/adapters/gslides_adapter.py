"""
Concrete implementation of abstract slides using gslides-api.
This module provides the actual implementation that maps abstract slide operations to gslides-api calls.
"""

import io
import logging
from typing import Annotated, Any, Optional, Union

import httpx
from pydantic import BaseModel, Discriminator, Field, Tag, TypeAdapter, model_validator

import gslides_api
from gslides_api import GoogleAPIClient, Presentation
from gslides_api import Slide as GSlide
from gslides_api.agnostic.element import MarkdownTableElement
from gslides_api.agnostic.units import EMU_PER_CM, EMU_PER_INCH, OutputUnit, from_emu
from gslides_api.domain.domain import ThumbnailSize
from gslides_api.domain.request import SubstringMatchCriteria
from gslides_api.domain.text import TextStyle
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.element import ImageElement, PageElement, TableElement
from gslides_api.element.shape import ShapeElement
from gslides_api.page.notes import Notes
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.request import ReplaceAllTextRequest

from gslides_api.common.log_time import log_time
from gslides_api.common.retry import retry
from gslides_api.adapters.abstract_slides import (
    AbstractAltText,
    AbstractElement,
    AbstractImageElement,
    AbstractPresentation,
    AbstractShapeElement,
    AbstractSlide,
    AbstractSlideProperties,
    AbstractSlidesAPIClient,
    AbstractSpeakerNotes,
    AbstractTableElement,
    AbstractThumbnail,
)

logger = logging.getLogger(__name__)


def concrete_element_discriminator(v: Any) -> str:
    """Discriminator to determine which ConcreteElement subclass based on type field."""
    if hasattr(v, "type"):
        # Handle ElementKind enum
        element_type = v.type
        if hasattr(element_type, "value"):
            element_type = element_type.value

        if element_type in ["SHAPE", "shape"]:
            return "shape"
        elif element_type in ["IMAGE", "image"]:
            return "image"
        elif element_type in ["TABLE", "table"]:
            return "table"
        else:
            # Fallback for other element types (LINE, VIDEO, WORD_ART, etc.)
            return "generic"

    raise ValueError(f"Cannot determine element type from: {v}")


class GSlidesAPIClient(AbstractSlidesAPIClient):
    def __init__(self, gslides_client: GoogleAPIClient | None = None):
        if gslides_client is None:
            gslides_client = gslides_api.client.api_client
        self.gslides_client = gslides_client

    @property
    def auto_flush(self):
        return self.gslides_client.auto_flush

    @auto_flush.setter
    def auto_flush(self, value: bool):
        self.gslides_client.auto_flush = value

    def flush_batch_update(self):
        pending_count = len(self.gslides_client.pending_batch_requests)
        pending_presentation = self.gslides_client.pending_presentation_id
        logger.info(
            f"FLUSH_BATCH_UPDATE: {pending_count} pending requests "
            f"for presentation {pending_presentation}"
        )
        result = self.gslides_client.flush_batch_update()
        replies = result.get("replies", []) if result else []
        logger.info(
            f"FLUSH_BATCH_UPDATE: completed, {len(replies)} replies, "
            f"result keys: {result.keys() if result else 'None'}"
        )
        # Log any non-empty replies (errors or meaningful responses)
        for i, reply in enumerate(replies):
            if reply:  # Non-empty reply
                logger.debug(f"FLUSH_BATCH_UPDATE reply[{i}]: {reply}")
        return result

    def copy_presentation(
        self, presentation_id: str, copy_title: str, folder_id: Optional[str] = None
    ) -> dict:
        return self.gslides_client.copy_presentation(
            presentation_id=presentation_id, copy_title=copy_title, folder_id=folder_id
        )

    def create_folder(
        self, name: str, ignore_existing: bool = True, parent_folder_id: Optional[str] = None
    ) -> dict:
        return self.gslides_client.create_folder(
            name, ignore_existing=ignore_existing, parent_folder_id=parent_folder_id
        )

    def delete_file(self, file_id: str):
        self.gslides_client.delete_file(file_id)

    def trash_file(self, file_id: str):
        self.gslides_client.trash_file(file_id)

    def upload_image_to_drive(self, image_path: str) -> str:
        return self.gslides_client.upload_image_to_drive(image_path)

    def set_credentials(self, credentials):
        from google.oauth2.credentials import Credentials

        # Store the abstract credentials for later retrieval
        self._abstract_credentials = credentials
        # Convert abstract credentials to concrete Google credentials
        google_creds = Credentials(
            token=credentials.token,
            refresh_token=credentials.refresh_token,
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
            token_uri=credentials.token_uri,
        )
        self.gslides_client.set_credentials(google_creds)

    def get_credentials(self):
        return getattr(self, "_abstract_credentials", None)

    def initialize_credentials(self, credential_location: str) -> None:
        gslides_api.initialize_credentials(credential_location)

    def replace_text(
        self, slide_ids: list[str], match_text: str, replace_text: str, presentation_id: str
    ):
        requests = [
            ReplaceAllTextRequest(
                pageObjectIds=slide_ids,
                containsText=SubstringMatchCriteria(text=match_text),
                replaceText=replace_text,
            )
        ]
        self.gslides_client.batch_update(requests, presentation_id)

    # Factory functions
    @classmethod
    def get_default_api_client(cls) -> "GSlidesAPIClient":
        """Get the default API client wrapped in concrete implementation."""
        return cls(gslides_api.client.api_client)

    @log_time
    async def get_presentation_as_pdf(self, presentation_id: str) -> io.BytesIO:
        api_client = self.gslides_client
        request = api_client.drive_service.files().export_media(
            fileId=presentation_id, mimeType="application/pdf"
        )

        async with httpx.AsyncClient() as client:
            response = await retry(
                client.get,
                args=(request.uri,),
                kwargs=dict(headers={"Authorization": f"Bearer {api_client.crdtls.token}"}),
                max_attempts=3,
                initial_delay=1.0,
                max_delay=30.0,
            )
            response.raise_for_status()
            return io.BytesIO(response.content)


class GSlidesSpeakerNotes(AbstractSpeakerNotes):
    def __init__(self, gslides_speaker_notes: Notes):
        super().__init__()
        self._gslides_speaker_notes = gslides_speaker_notes

    def read_text(self, as_markdown: bool = True) -> str:
        return self._gslides_speaker_notes.read_text(as_markdown=as_markdown)

    def write_text(self, api_client: GSlidesAPIClient, content: str):
        self._gslides_speaker_notes.write_text(content, api_client=api_client.gslides_client)


class GSlidesElementParent(AbstractElement):
    """Generic concrete element for unsupported element types (LINE, VIDEO, etc.)."""

    gslides_element: Any = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_page_element(cls, data: Any) -> dict:
        # Accept any PageElement
        gslides_element = data

        return {
            "objectId": gslides_element.objectId,
            "presentation_id": gslides_element.presentation_id,
            "slide_id": gslides_element.slide_id,
            "alt_text": AbstractAltText(
                title=gslides_element.alt_text.title,
                description=gslides_element.alt_text.description,
            ),
            "type": (
                str(gslides_element.type.value)
                if hasattr(gslides_element.type, "value")
                else str(gslides_element.type)
            ),
            "gslides_element": gslides_element,
        }

    def absolute_size(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        return self.gslides_element.absolute_size(units=units)

    # def element_properties(self) -> dict:
    #     return self.gslides_element.element_properties()

    def absolute_position(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        return self.gslides_element.absolute_position(units=units)

    def to_markdown_element(self, name: str | None = None) -> Any:
        if name is None:
            # Different element types have different default names
            return self.gslides_element.to_markdown_element()
        else:
            return self.gslides_element.to_markdown_element(name=name)

    def create_image_element_like(self, api_client: GSlidesAPIClient) -> "GSlidesImageElement":
        gslides_element = self.gslides_element.create_image_element_like(
            api_client=api_client.gslides_client
        )
        return GSlidesImageElement(gslides_element=gslides_element)

    def set_alt_text(
        self,
        api_client: GSlidesAPIClient,
        title: str | None = None,
        description: str | None = None,
    ):
        self.gslides_element.set_alt_text(
            title=title, description=description, api_client=api_client.gslides_client
        )


class GSlidesShapeElement(AbstractShapeElement, GSlidesElementParent):
    gslides_element: ShapeElement = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_page_element(cls, data: Any) -> dict:
        if isinstance(data, ShapeElement):
            gslides_shape = data
        elif hasattr(data, "shape"):  # It's a PageElement with shape
            gslides_shape = data
        else:
            raise ValueError(f"Expected ShapeElement or PageElement with shape, got {type(data)}")

        return {
            "objectId": gslides_shape.objectId,
            "presentation_id": gslides_shape.presentation_id,
            "slide_id": gslides_shape.slide_id,
            "alt_text": AbstractAltText(
                title=gslides_shape.alt_text.title,
                description=gslides_shape.alt_text.description,
            ),
            "gslides_element": gslides_shape,
        }

    @property
    def has_text(self) -> bool:
        return self.gslides_element.has_text

    def write_text(self, api_client: GSlidesAPIClient, content: str, autoscale: bool = False):
        # Extract styles BEFORE writing to preserve original element styling.
        # skip_whitespace=True avoids picking up invisible spacer styles (e.g. white theme colors).
        # If no non-whitespace styles exist, TextContent.styles() falls back to including whitespace.
        styles = self.gslides_element.styles(skip_whitespace=True)
        logger.debug(
            f"GSlidesShapeElement.write_text: objectId={self.objectId}, "
            f"content={repr(content[:100] if content else None)}, autoscale={autoscale}, "
            f"styles_count={len(styles) if styles else 0}"
        )
        result = self.gslides_element.write_text(
            content,
            autoscale=autoscale,
            styles=styles,
            api_client=api_client.gslides_client,
        )
        logger.debug(f"GSlidesShapeElement.write_text: result={result}")

    def read_text(self, as_markdown: bool = True) -> str:
        return self.gslides_element.read_text(as_markdown=as_markdown)

    def styles(self, skip_whitespace: bool = True) -> list[TextStyle] | None:
        return self.gslides_element.styles(skip_whitespace=skip_whitespace)


class GSlidesImageElement(AbstractImageElement, GSlidesElementParent):
    gslides_element: ImageElement = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_page_element(cls, data: Any) -> dict:
        if isinstance(data, dict):
            if "gslides_element" in data:
                data = data["gslides_element"]
            else:
                raise ValueError(f"Need to supply gslides_element in dict, got {data}")

        if isinstance(data, ImageElement):
            gslides_image = data
        elif hasattr(data, "image"):  # It's a PageElement with image
            gslides_image = data
        else:
            raise ValueError(f"Expected ImageElement or PageElement with image, got {type(data)}")

        return {
            "objectId": gslides_image.objectId,
            "presentation_id": gslides_image.presentation_id,
            "slide_id": gslides_image.slide_id,
            "alt_text": AbstractAltText(
                title=gslides_image.alt_text.title,
                description=gslides_image.alt_text.description,
            ),
            "gslides_element": gslides_image,
        }

    def replace_image(
        self,
        api_client: GSlidesAPIClient,
        file: str | None = None,
        url: str | None = None,
    ):
        # Clear cropProperties to avoid Google Slides API error:
        # "CropProperties offsets cannot be updated individually"
        # This happens when recreating elements that have partial crop settings.
        if (
            self.gslides_element.image
            and self.gslides_element.image.imageProperties
            and hasattr(self.gslides_element.image.imageProperties, "cropProperties")
        ):
            self.gslides_element.image.imageProperties.cropProperties = None

        new_element = self.gslides_element.replace_image(
            file=file,
            url=url,
            api_client=api_client.gslides_client,
            enforce_size="auto",
            recreate_element=True,
        )
        if new_element is not None:
            self.gslides_element = new_element

    def get_image_data(self):
        # Get image data from the gslides image and return it directly
        return self.gslides_element.get_image_data()


class GSlidesTableElement(AbstractTableElement, GSlidesElementParent):
    gslides_element: TableElement = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_page_element(cls, data: Any) -> dict:
        if isinstance(data, TableElement):
            gslides_table = data
        elif hasattr(data, "table"):  # It's a PageElement with table
            gslides_table = data
        else:
            raise ValueError(f"Expected TableElement or PageElement with table, got {type(data)}")

        return {
            "objectId": gslides_table.objectId,
            "presentation_id": gslides_table.presentation_id,
            "slide_id": gslides_table.slide_id,
            "alt_text": AbstractAltText(
                title=gslides_table.alt_text.title,
                description=gslides_table.alt_text.description,
            ),
            "gslides_element": gslides_table,
        }

    def resize(
        self,
        api_client: GSlidesAPIClient,
        rows: int,
        cols: int,
        fix_width: bool = True,
        fix_height: bool = True,
        target_height_in: float | None = None,
    ) -> float:
        """Resize the table.

        Args:
            target_height_in: If provided, constrain total table height (rows + borders)
                                    to this value in inches. Scales both row heights and border
                                    weights proportionally.

        Returns:
            Font scale factor (1.0 if no scaling, < 1.0 if rows were added with fix_height)
        """
        target_height_emu = None
        if target_height_in is not None:
            target_height_emu = target_height_in * EMU_PER_INCH

        requests, font_scale_factor = self.gslides_element.resize_requests(
            rows,
            cols,
            fix_width=fix_width,
            fix_height=fix_height,
            target_height_emu=target_height_emu,
        )
        api_client.gslides_client.batch_update(requests, self.presentation_id)
        return font_scale_factor

    def get_horizontal_border_weight(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Get weight of horizontal borders in specified units."""
        return self.gslides_element.get_horizontal_border_weight(units=units)

    def get_row_count(self) -> int:
        """Get current number of rows."""
        return self.gslides_element.table.rows

    def get_column_count(self) -> int:
        """Get current number of columns."""
        return self.gslides_element.table.columns

    def update_content(
        self,
        api_client: GSlidesAPIClient,
        markdown_content: MarkdownTableElement,
        check_shape: bool = True,
        font_scale_factor: float = 1.0,
    ):
        requests = self.gslides_element.content_update_requests(
            markdown_content, check_shape=check_shape, font_scale_factor=font_scale_factor
        )
        api_client.gslides_client.batch_update(requests, self.presentation_id)

    def to_markdown_element(self, name: str | None = None) -> Any:
        return self.gslides_element.to_markdown_element(name=name)


# Discriminated union type for concrete elements
GSlidesElement = Annotated[
    Union[
        Annotated[GSlidesShapeElement, Tag("shape")],
        Annotated[GSlidesImageElement, Tag("image")],
        Annotated[GSlidesTableElement, Tag("table")],
        Annotated[GSlidesElementParent, Tag("generic")],
    ],
    Discriminator(concrete_element_discriminator),
]

# TypeAdapter for validating the discriminated union
_concrete_element_adapter = TypeAdapter(GSlidesElement)


def validate_concrete_element(page_element: PageElement) -> GSlidesElement:
    """Create the appropriate concrete element from a PageElement."""
    return _concrete_element_adapter.validate_python(page_element)


class GSlidesSlide(AbstractSlide):
    def __init__(self, gslides_slide: GSlide):
        # Convert gslides elements to abstract elements, skipping group containers.
        # unroll_group_elements includes both the group wrapper and its children;
        # children are real elements while the group container has no renderable
        # content and may lack size/transform, causing downstream crashes.
        elements = []
        for element in gslides_slide.page_elements_flat:
            if element.type == ElementKind.GROUP:
                continue
            concrete_element = validate_concrete_element(element)
            elements.append(concrete_element)

        super().__init__(
            elements=elements,
            objectId=gslides_slide.objectId,
            slideProperties=AbstractSlideProperties(
                isSkipped=gslides_slide.slideProperties.isSkipped or False
            ),
            speaker_notes=GSlidesSpeakerNotes(gslides_slide.speaker_notes),
        )
        self._gslides_slide = gslides_slide

    def thumbnail(
        self, api_client: GSlidesAPIClient, size: str, include_data: bool = False
    ) -> AbstractThumbnail:
        """Get thumbnail for a Google Slides slide.

        Args:
            api_client: The Google Slides API client
            size: The thumbnail size (e.g., "MEDIUM")
            include_data: If True, downloads the thumbnail image data with retry logic

        Returns:
            AbstractThumbnail with metadata and optionally the image content
        """
        import http
        import ssl

        from gslides_api.common.google_errors import detect_file_access_denied_error
        from gslides_api.common.download import download_binary_file

        # Map size string to ThumbnailSize enum
        thumbnail_size = getattr(ThumbnailSize, size, ThumbnailSize.MEDIUM)

        # Fetch thumbnail metadata with retry
        @retry(
            max_attempts=6,
            initial_delay=1.0,
            max_delay=15.0,
            exceptions=(
                TimeoutError,
                httpx.TimeoutException,
                httpx.RequestError,
                ConnectionError,
                ssl.SSLError,
                http.client.ResponseNotReady,
                http.client.IncompleteRead,
            ),
        )
        def fetch_thumbnail():
            return self._gslides_slide.thumbnail(
                size=thumbnail_size, api_client=api_client.gslides_client
            )

        try:
            gslides_thumbnail = fetch_thumbnail()
        except Exception as e:
            # Check if this is a file access denied error (drive.file scope)
            detect_file_access_denied_error(error=e, file_id=self.presentation_id)
            raise

        # Handle mime_type format: Google Slides returns "png", we need "image/png"
        mime_type = (
            gslides_thumbnail.mime_type
            if gslides_thumbnail.mime_type.startswith("image/")
            else f"image/{gslides_thumbnail.mime_type}"
        )

        content = gslides_thumbnail.payload if include_data else None
        file_size = len(content) if content else None

        return AbstractThumbnail(
            contentUrl=gslides_thumbnail.contentUrl,
            width=gslides_thumbnail.width,
            height=gslides_thumbnail.height,
            mime_type=mime_type,
            content=content,
            file_size=file_size,
        )


class GSlidesPresentation(AbstractPresentation):
    def __init__(self, gslides_presentation: Presentation):
        # Convert gslides slides to abstract slides
        slides = [GSlidesSlide(slide) for slide in gslides_presentation.slides]

        super().__init__(
            slides=slides,
            url=gslides_presentation.url,
            presentationId=getattr(gslides_presentation, "presentationId", ""),
            revisionId=getattr(gslides_presentation, "revisionId", ""),
            title=gslides_presentation.title,
        )
        self._gslides_presentation = gslides_presentation

    @property
    def url(self) -> str:
        return self._gslides_presentation.url

    def slide_height(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Return slide height in specified units."""
        height_dim = self._gslides_presentation.pageSize.height
        height_emu = height_dim.magnitude if hasattr(height_dim, "magnitude") else float(height_dim)
        return from_emu(height_emu, units)

    @classmethod
    def from_id(cls, api_client: GSlidesAPIClient, presentation_id: str) -> "GSlidesPresentation":
        from gslides_api.common.google_errors import detect_file_access_denied_error

        try:
            gslides_presentation = Presentation.from_id(
                presentation_id, api_client=api_client.gslides_client
            )
            return cls(gslides_presentation)
        except Exception as e:
            # Check if this is a file access denied error (drive.file scope)
            detect_file_access_denied_error(error=e, file_id=presentation_id)
            # If not a file access denied error, re-raise the original exception
            raise

    def copy_via_drive(
        self,
        api_client: GSlidesAPIClient,
        copy_title: str,
        folder_id: Optional[str] = None,
    ) -> "GSlidesPresentation":
        from gslides_api.common.google_errors import detect_file_access_denied_error

        try:
            copied_presentation = self._gslides_presentation.copy_via_drive(
                copy_title=copy_title, api_client=api_client.gslides_client, folder_id=folder_id
            )
            return GSlidesPresentation(copied_presentation)
        except Exception as e:
            # Check if this is a file access denied error (drive.file scope)
            detect_file_access_denied_error(error=e, file_id=self.presentationId)
            raise

    def insert_copy(
        self,
        source_slide: GSlidesSlide,
        api_client: GSlidesAPIClient,
        insertion_index: int | None = None,
    ):
        # Use the new duplicate_slide method
        new_gslide = self.duplicate_slide(source_slide, api_client)
        if insertion_index is not None:
            self.move_slide(new_gslide, insertion_index, api_client)
        return new_gslide

    def sync_from_cloud(self, api_client: GSlidesAPIClient):
        self._gslides_presentation.sync_from_cloud(api_client=api_client.gslides_client)
        # Rebuild the GSlidesSlide wrappers so they reflect the refreshed state
        # (e.g. new objectIds for speaker notes elements after slide duplication).
        self.slides = [GSlidesSlide(slide) for slide in self._gslides_presentation.slides]
        self.presentationId = getattr(self._gslides_presentation, "presentationId", "")
        self.revisionId = getattr(self._gslides_presentation, "revisionId", "")

    def save(self, api_client: GSlidesAPIClient) -> None:
        """Save/persist all changes made to this presentation."""
        api_client.flush_batch_update()

    def delete_slide(self, slide: Union[GSlidesSlide, int], api_client: GSlidesAPIClient):
        """Delete a slide from the presentation by reference or index."""
        if isinstance(slide, int):
            slide = self.slides[slide]
        if isinstance(slide, GSlidesSlide):
            # Use the existing delete logic from GSlidesSlide
            slide._gslides_slide.delete(api_client=api_client.gslides_client)
            # Remove from our slides list
            self.slides.remove(slide)

    def delete_slides(self, slides: list[Union[GSlidesSlide, int]], api_client: GSlidesAPIClient):
        for slide in slides:
            self.delete_slide(slide, api_client)

    def move_slide(
        self, slide: Union[GSlidesSlide, int], insertion_index: int, api_client: GSlidesAPIClient
    ):
        """Move a slide to a new position within the presentation."""
        if isinstance(slide, int):
            slide = self.slides[slide]
        if isinstance(slide, GSlidesSlide):
            # Use the existing move logic from GSlidesSlide
            slide._gslides_slide.move(
                insertion_index=insertion_index, api_client=api_client.gslides_client
            )
            # Update local slides list order
            self.slides.remove(slide)
            self.slides.insert(insertion_index, slide)

    def duplicate_slide(
        self, slide: Union[GSlidesSlide, int], api_client: GSlidesAPIClient
    ) -> GSlidesSlide:
        """Duplicate a slide within the presentation."""
        if isinstance(slide, int):
            slide = self.slides[slide]
        if isinstance(slide, GSlidesSlide):
            # Use the existing duplicate logic from GSlidesSlide
            duplicated = slide._gslides_slide.duplicate(api_client=api_client.gslides_client)
            new_slide = GSlidesSlide(duplicated)
            self.slides.append(new_slide)
            # Manually set parent refs since validator only runs at construction
            new_slide._parent_presentation = self
            for element in new_slide.elements:
                element._parent_presentation = self
            return new_slide
        else:
            raise ValueError("slide must be a GSlidesSlide or int")
