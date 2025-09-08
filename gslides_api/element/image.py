import logging
import mimetypes
import uuid
from typing import List, Optional
from urllib.parse import urlparse

import requests
from pydantic import Field, field_validator

from gslides_api.client import GoogleAPIClient
from gslides_api.domain.domain import Image, ImageData, ImageReplaceMethod, PageElementProperties
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.markdown.element import MarkdownImageElement as MarkdownImageElement
from gslides_api.request.request import (
    CreateImageRequest,
    ReplaceImageRequest,
    UpdateImagePropertiesRequest,
)
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.utils import dict_to_dot_separated_field_list, image_url_is_valid


class ImageElement(PageElementBase):
    """Represents an image element on a slide."""

    image: Image
    type: ElementKind = Field(
        default=ElementKind.IMAGE, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.IMAGE

    @staticmethod
    def create_image_request_like(
        e: PageElementBase,
        image_id: str | None = None,
        url: str | None = None,
        parent_id: str | None = None,
    ) -> List[GSlidesAPIRequest]:
        """Create a request to create an image element like the given element."""
        url = url or "https://upload.wikimedia.org/wikipedia/commons/2/2d/Logo_Google_blanco.png"
        element_properties = e.element_properties(parent_id or e.slide_id)
        request = CreateImageRequest(
            objectId=image_id,
            elementProperties=element_properties,
            url=url,
        )
        return [request]

    @staticmethod
    def create_image_element_like(
        e: PageElementBase,
        api_client: GoogleAPIClient | None = None,
        parent_id: str | None = None,
        url: str | None = None,
    ) -> str:

        api_client = api_client or globals()["api_client"]
        parent_id = parent_id or e.slide_id

        # Create the image element
        image_id = uuid.uuid4().hex
        requests = ImageElement.create_image_request_like(
            e,
            parent_id=parent_id,
            url=url,
            image_id=image_id,
        )
        api_client.batch_update(requests, e.presentation_id)
        return image_id

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert an ImageElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        request = CreateImageRequest(
            elementProperties=element_properties,
            url=self.image.contentUrl,
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert an ImageElement to an update request for the Google Slides API."""
        requests = self.alt_text_update_request(element_id)

        if hasattr(self.image, "imageProperties") and self.image.imageProperties is not None:
            image_properties = self.image.imageProperties.to_api_format()
            # "fields": "*" causes an error
            request = UpdateImagePropertiesRequest(
                objectId=element_id,
                imageProperties=self.image.imageProperties,
                fields=",".join(dict_to_dot_separated_field_list(image_properties)),
            )
            requests.append(request)

        return requests

    def to_markdown(self) -> str | None:
        url = self.image.sourceUrl
        if url is None:
            return None
        description = self.title or "Image"
        return f"![{description}]({url})"

    @staticmethod
    def _replace_image_requests(
        objectId: str, new_url: str, method: ImageReplaceMethod | None = None
    ):
        """
        Replace image by URL with validation.

        Args:
            new_url: New image URL
            method: Optional image replacement method

        Returns:
            List of requests to replace the image
        """
        if not new_url.startswith(("http://", "https://")):
            raise ValueError("Image URL must start with http:// or https://")

        # Validate URL before attempting replacement
        if not image_url_is_valid(new_url):
            raise ValueError(f"Image URL is not accessible or invalid: {new_url}")

        request = ReplaceImageRequest(
            imageObjectId=objectId,
            url=new_url,
            imageReplaceMethod=method.value if method is not None else None,
        )
        return [request]

    def replace_image(
        self,
        url: str | None = None,
        file: str | None = None,
        method: ImageReplaceMethod | None = None,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        # if url is None and file is None:
        #     raise ValueError("Must specify either url or file")
        # if url is not None and file is not None:
        #     raise ValueError("Must specify either url or file, not both")
        #
        # client = api_client or globals()["api_client"]
        # if file is not None:
        #     url = client.upload_image_to_drive(file)
        #
        # requests = self._replace_image_requests(url, method)
        # return client.batch_update(requests, self.presentation_id)
        return ImageElement.replace_image_from_id(
            self.objectId,
            self.presentation_id,
            url=url,
            file=file,
            method=method,
            api_client=api_client,
        )

    @staticmethod
    def replace_image_from_id(
        image_id: str,
        presentation_id: str,
        url: str | None = None,
        file: str | None = None,
        method: ImageReplaceMethod | None = None,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        if url is None and file is None:
            raise ValueError("Must specify either url or file")
        if url is not None and file is not None:
            raise ValueError("Must specify either url or file, not both")

        client = api_client or globals()["api_client"]
        if file is not None:
            url = client.upload_image_to_drive(file)

        requests = ImageElement._replace_image_requests(image_id, url, method)
        return client.batch_update(requests, presentation_id)

    def get_image_data(self) -> ImageData:
        """Retrieve the actual image data from Google Slides.

        Returns:
            ImageData: Container with image bytes, MIME type, and optional filename.

        Raises:
            ValueError: If no image URL is available.
            requests.RequestException: If the image download fails.
        """
        logger = logging.getLogger(__name__)

        # Prefer contentUrl over sourceUrl as it's Google's cached version
        url = self.image.contentUrl or self.image.sourceUrl

        if not url:
            logger.error("No image URL available for element %s", self.objectId)
            raise ValueError("No image URL available (neither contentUrl nor sourceUrl)")

        logger.info("Downloading image from URL: %s", url)

        try:
            # Download the image with retries for common network issues
            response = requests.get(url, timeout=30)
            response.raise_for_status()

            content_length = len(response.content)
            logger.debug("Downloaded %d bytes from %s", content_length, url)

            if content_length == 0:
                logger.warning("Downloaded empty image content from %s", url)
                raise ValueError("Downloaded image is empty")

        except requests.exceptions.Timeout as e:
            logger.error("Timeout downloading image from %s: %s", url, e)
            raise requests.RequestException(f"Timeout downloading image: {e}") from e
        except requests.exceptions.RequestException as e:
            logger.error("Failed to download image from %s: %s", url, e)
            raise

        # Determine MIME type
        mime_type = response.headers.get("content-type", "application/octet-stream")
        logger.debug("Content-Type header: %s", mime_type)

        # If MIME type is not image-specific, try to guess from URL
        if not mime_type.startswith("image/"):
            parsed_url = urlparse(url)
            path = parsed_url.path
            if path:
                guessed_type, _ = mimetypes.guess_type(path)
                if guessed_type and guessed_type.startswith("image/"):
                    logger.debug("Guessed MIME type from URL: %s -> %s", path, guessed_type)
                    mime_type = guessed_type
                else:
                    logger.warning(
                        "Could not determine image MIME type, using default: %s",
                        mime_type,
                    )

        # Extract filename from URL if possible
        filename = None
        parsed_url = urlparse(url)
        if parsed_url.path:
            filename = parsed_url.path.split("/")[-1]
            # Only keep if it looks like a filename with extension
            if "." not in filename:
                filename = None
            else:
                logger.debug("Extracted filename from URL: %s", filename)

        logger.info(
            "Successfully retrieved image: %d bytes, MIME type: %s",
            content_length,
            mime_type,
        )

        return ImageData(content=response.content, mime_type=mime_type, filename=filename)

    def to_markdown_element(self, name: str = "Image") -> MarkdownImageElement:
        """Convert ImageElement to MarkdownImageElement for round-trip conversion."""

        # Use sourceUrl preferentially, fallback to contentUrl
        url = self.image.sourceUrl or self.image.contentUrl or ""
        alt_text = self.title or self.description or ""

        # Create the markdown image content
        markdown_content = f"![{alt_text}]({url})"

        # Store all necessary metadata for perfect reconstruction
        metadata = {
            "objectId": self.objectId,
            "sourceUrl": self.image.sourceUrl,
            "contentUrl": self.image.contentUrl,
            "alt_text": alt_text,
            "original_markdown": markdown_content,
        }

        # Store element properties (position, size, etc.) if available
        if hasattr(self, "size") and self.size:
            metadata["size"] = {
                "width": self.size.width.magnitude,
                "height": self.size.height.magnitude,
                "unit": self.size.width.unit.value,
            }

        if hasattr(self, "transform") and self.transform:
            metadata["transform"] = (
                self.transform.to_api_format() if hasattr(self.transform, "to_api_format") else None
            )

        # Store title and description if available
        if hasattr(self, "title") and self.title:
            metadata["title"] = self.title
        if hasattr(self, "description") and self.description:
            metadata["description"] = self.description

        # Store image properties if available
        if hasattr(self.image, "imageProperties") and self.image.imageProperties:
            metadata["imageProperties"] = (
                self.image.imageProperties.to_api_format()
                if hasattr(self.image.imageProperties, "to_api_format")
                else None
            )

        return MarkdownImageElement(name=name, content=markdown_content, metadata=metadata)

    @classmethod
    def from_markdown_element(
        cls,
        markdown_elem: MarkdownImageElement,
        parent_id: str,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "ImageElement":
        """Create ImageElement from MarkdownImageElement with preserved metadata."""

        # Extract metadata
        metadata = markdown_elem.metadata or {}
        object_id = metadata.get("objectId")

        # Get the URL from the content (it's stored as URL in MarkdownImageElement.content)
        url = markdown_elem.content

        # Create the Image domain object
        image = Image(
            contentUrl=metadata.get("contentUrl"),
            sourceUrl=metadata.get("sourceUrl") or url,  # Fallback to content URL
        )

        # Restore image properties if available
        if "imageProperties" in metadata and metadata["imageProperties"]:
            from gslides_api.domain.domain import ImageProperties

            image.imageProperties = ImageProperties(**metadata["imageProperties"])

        # Create element properties from metadata

        element_props = PageElementProperties(pageObjectId=parent_id)

        # Restore size if available, otherwise provide default
        if "size" in metadata:
            size_data = metadata["size"]
            from gslides_api.domain.domain import Size
            from gslides_api.domain.domain import Dimension
            from gslides_api.domain.domain import Unit

            element_props.size = Size(
                width=Dimension(magnitude=size_data["width"], unit=Unit(size_data["unit"])),
                height=Dimension(magnitude=size_data["height"], unit=Unit(size_data["unit"])),
            )
        else:
            # Provide default size for images
            from gslides_api.domain.domain import Size
            from gslides_api.domain.domain import Dimension
            from gslides_api.domain.domain import Unit

            element_props.size = Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=150, unit=Unit.PT),
            )

        # Restore transform if available, otherwise create default
        if "transform" in metadata and metadata["transform"]:
            from gslides_api.domain.domain import Transform

            transform_data = metadata["transform"]
            element_props.transform = Transform(**transform_data)
        else:
            # Create a default identity transform
            from gslides_api.domain.domain import Transform

            element_props.transform = Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            )

        # Create the image element
        image_element = cls(
            objectId=object_id or "image_" + str(hash(markdown_elem.content))[:8],
            size=element_props.size,
            transform=element_props.transform,
            title=metadata.get("title"),
            description=metadata.get("description"),
            image=image,
            slide_id=parent_id,
            presentation_id="",  # Will need to be set by caller
        )

        return image_element
