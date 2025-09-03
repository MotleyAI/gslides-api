import logging
import mimetypes
import uuid
from typing import Annotated, Any, Dict, List, Optional, Union
from urllib.parse import urlparse

import requests
from pydantic import Discriminator, Field, Tag, field_validator

from gslides_api.client import GoogleAPIClient, api_client
from gslides_api.domain import (Group, Image, ImageData, ImageReplaceMethod,
                                Line, SheetsChart, SpeakerSpotlight, Table,
                                Video, WordArt)
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.shape import ShapeElement
from gslides_api.markdown.element import ImageElement as MarkdownImageElement
from gslides_api.markdown.element import TableData
from gslides_api.markdown.element import TableElement as MarkdownTableElement
from gslides_api.request.request import (  # UpdateSheetsChartPropertiesRequest,; CreateWordArtRequest,
    CreateImageRequest, CreateLineRequest, CreateSheetsChartRequest,
    CreateVideoRequest, GSlidesAPIRequest, ReplaceImageRequest,
    UpdateImagePropertiesRequest, UpdateLinePropertiesRequest,
    UpdateVideoPropertiesRequest)
from gslides_api.request.table import CreateTableRequest
from gslides_api.utils import (dict_to_dot_separated_field_list,
                               image_url_is_valid)


def element_discriminator(v: Any) -> str:
    """Discriminator function to determine which PageElement subclass to use based on which field is present."""
    if isinstance(v, dict):
        if v.get("shape") is not None:
            return "shape"
        elif v.get("table") is not None:
            return "table"
        elif v.get("image") is not None:
            return "image"
        elif v.get("video") is not None:
            return "video"
        elif v.get("line") is not None:
            return "line"
        elif v.get("wordArt") is not None:
            return "wordArt"
        elif v.get("sheetsChart") is not None:
            return "sheetsChart"
        elif v.get("speakerSpotlight") is not None:
            return "speakerSpotlight"
        elif v.get("elementGroup") is not None:
            return "group"
    else:
        # Handle model instances
        if hasattr(v, "shape") and v.shape is not None:
            return "shape"
        elif hasattr(v, "table") and v.table is not None:
            return "table"
        elif hasattr(v, "image") and v.image is not None:
            return "image"
        elif hasattr(v, "video") and v.video is not None:
            return "video"
        elif hasattr(v, "line") and v.line is not None:
            return "line"
        elif hasattr(v, "wordArt") and v.wordArt is not None:
            return "wordArt"
        elif hasattr(v, "sheetsChart") and v.sheetsChart is not None:
            return "sheetsChart"
        elif hasattr(v, "speakerSpotlight") and v.speakerSpotlight is not None:
            return "speakerSpotlight"
        elif hasattr(v, "elementGroup") and v.elementGroup is not None:
            return "group"

    # Return None if no discriminator found - this will raise an error
    return None


class TableElement(PageElementBase):
    """Represents a table element on a slide."""

    table: Table
    type: ElementKind = Field(
        default=ElementKind.TABLE, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.TABLE

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        request = CreateTableRequest(
            elementProperties=element_properties.to_api_format(),
            rows=self.table.rows,
            columns=self.table.columns,
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a TableElement to an update request for the Google Slides API."""
        # Tables don't have specific properties to update beyond base properties
        return self.alt_text_update_request(element_id)

    def extract_table_data(self) -> TableData:
        """Extract table data from Google Slides Table structure into simple TableData format."""
        if not self.table.tableRows or not self.table.rows or not self.table.columns:
            raise ValueError("Table has no data to extract")

        headers = []
        rows = []

        # Extract data from tableRows
        for row_idx, table_row in enumerate(self.table.tableRows):
            row_cells = []

            # Each table row contains tableCells
            if "tableCells" in table_row:
                for cell in table_row["tableCells"]:
                    # Extract text content from cell
                    cell_text = ""
                    if "text" in cell and "textElements" in cell["text"]:
                        text_parts = []
                        for text_element in cell["text"]["textElements"]:
                            if (
                                "textRun" in text_element
                                and "content" in text_element["textRun"]
                            ):
                                text_parts.append(text_element["textRun"]["content"])
                        cell_text = "".join(text_parts).strip()
                    row_cells.append(cell_text)

            # First row is typically headers
            if row_idx == 0:
                headers = row_cells
            else:
                # Pad row with empty strings if it's shorter than headers
                while len(row_cells) < len(headers):
                    row_cells.append("")
                rows.append(row_cells[: len(headers)])  # Trim if longer than headers

        if not headers:
            raise ValueError("No headers found in table")

        return TableData(headers=headers, rows=rows)

    def to_markdown_element(self, name: str = "Table") -> MarkdownTableElement:
        """Convert TableElement to MarkdownTableElement for round-trip conversion."""

        # Check if we have stored table data from markdown conversion
        if hasattr(self, "_markdown_table_data") and self._markdown_table_data:
            table_data = self._markdown_table_data
        else:
            # Extract table data from Google Slides structure
            try:
                table_data = self.extract_table_data()
            except ValueError as e:
                # If we can't extract data, create an empty table
                table_data = TableData(headers=["Column 1"], rows=[])

        # Store all necessary metadata for perfect reconstruction
        metadata = {
            "objectId": self.objectId,
            "rows": self.table.rows,
            "columns": self.table.columns,
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
                self.transform.to_api_format()
                if hasattr(self.transform, "to_api_format")
                else None
            )

        # Store title and description if available
        if hasattr(self, "title") and self.title:
            metadata["title"] = self.title
        if hasattr(self, "description") and self.description:
            metadata["description"] = self.description

        # Store raw table structure for perfect reconstruction
        if self.table.tableRows:
            metadata["tableRows"] = self.table.tableRows
        if self.table.tableColumns:
            metadata["tableColumns"] = self.table.tableColumns
        if self.table.horizontalBorderRows:
            metadata["horizontalBorderRows"] = self.table.horizontalBorderRows
        if self.table.verticalBorderRows:
            metadata["verticalBorderRows"] = self.table.verticalBorderRows

        return MarkdownTableElement(name=name, content=table_data, metadata=metadata)

    @classmethod
    def from_markdown_element(
        cls,
        markdown_elem: MarkdownTableElement,
        parent_id: str,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "TableElement":
        """Create TableElement from MarkdownTableElement with preserved metadata."""

        # Extract metadata
        metadata = markdown_elem.metadata or {}
        object_id = metadata.get("objectId")

        # Get table data
        table_data = markdown_elem.content

        # Create the Table domain object
        table = Table(
            rows=metadata.get("rows", len(table_data.rows) + 1),  # +1 for header
            columns=metadata.get("columns", len(table_data.headers)),
        )

        # Restore full table structure if available in metadata
        if "tableRows" in metadata:
            table.tableRows = metadata["tableRows"]
        if "tableColumns" in metadata:
            table.tableColumns = metadata["tableColumns"]
        if "horizontalBorderRows" in metadata:
            table.horizontalBorderRows = metadata["horizontalBorderRows"]
        if "verticalBorderRows" in metadata:
            table.verticalBorderRows = metadata["verticalBorderRows"]

        # Create element properties from metadata
        from gslides_api.domain import PageElementProperties

        element_props = PageElementProperties(pageObjectId=parent_id)

        # Restore size if available, otherwise provide default
        if "size" in metadata:
            size_data = metadata["size"]
            from gslides_api.domain import Dimension, Size, Unit

            element_props.size = Size(
                width=Dimension(
                    magnitude=size_data["width"], unit=Unit(size_data["unit"])
                ),
                height=Dimension(
                    magnitude=size_data["height"], unit=Unit(size_data["unit"])
                ),
            )
        else:
            # Provide default size for tables
            from gslides_api.domain import Dimension, Size, Unit

            # Calculate size based on table dimensions
            num_rows = len(table_data.rows) + 1  # +1 for header
            num_cols = len(table_data.headers)

            # Basic sizing: 100pt per column, 30pt per row
            default_width = max(300, num_cols * 100)
            default_height = max(150, num_rows * 30)

            element_props.size = Size(
                width=Dimension(magnitude=default_width, unit=Unit.PT),
                height=Dimension(magnitude=default_height, unit=Unit.PT),
            )

        # Restore transform if available, otherwise create default
        if "transform" in metadata and metadata["transform"]:
            from gslides_api.domain import Transform

            transform_data = metadata["transform"]
            element_props.transform = Transform(**transform_data)
        else:
            # Create a default identity transform
            from gslides_api.domain import Transform

            element_props.transform = Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            )

        # Create the table element
        table_element = cls(
            objectId=object_id
            or "table_" + str(hash(str(markdown_elem.content.headers)))[:8],
            size=element_props.size,
            transform=element_props.transform,
            title=metadata.get("title"),
            description=metadata.get("description"),
            table=table,
            slide_id=parent_id,
            presentation_id="",  # Will need to be set by caller
        )

        # Store the markdown table data for round-trip conversion
        table_element._markdown_table_data = table_data

        return table_element


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
        url = (
            url
            or "https://upload.wikimedia.org/wikipedia/commons/2/2d/Logo_Google_blanco.png"
        )
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
        from gslides_api.page.slide import Slide

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

        if (
            hasattr(self.image, "imageProperties")
            and self.image.imageProperties is not None
        ):
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
            raise ValueError(
                "No image URL available (neither contentUrl nor sourceUrl)"
            )

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
                    logger.debug(
                        "Guessed MIME type from URL: %s -> %s", path, guessed_type
                    )
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

        return ImageData(
            content=response.content, mime_type=mime_type, filename=filename
        )

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
                self.transform.to_api_format()
                if hasattr(self.transform, "to_api_format")
                else None
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

        return MarkdownImageElement(
            name=name, content=markdown_content, metadata=metadata
        )

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
            from gslides_api.domain import ImageProperties

            image.imageProperties = ImageProperties(**metadata["imageProperties"])

        # Create element properties from metadata
        from gslides_api.domain import PageElementProperties

        element_props = PageElementProperties(pageObjectId=parent_id)

        # Restore size if available, otherwise provide default
        if "size" in metadata:
            size_data = metadata["size"]
            from gslides_api.domain import Dimension, Size, Unit

            element_props.size = Size(
                width=Dimension(
                    magnitude=size_data["width"], unit=Unit(size_data["unit"])
                ),
                height=Dimension(
                    magnitude=size_data["height"], unit=Unit(size_data["unit"])
                ),
            )
        else:
            # Provide default size for images
            from gslides_api.domain import Dimension, Size, Unit

            element_props.size = Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=150, unit=Unit.PT),
            )

        # Restore transform if available, otherwise create default
        if "transform" in metadata and metadata["transform"]:
            from gslides_api.domain import Transform

            transform_data = metadata["transform"]
            element_props.transform = Transform(**transform_data)
        else:
            # Create a default identity transform
            from gslides_api.domain import Transform

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


class VideoElement(PageElementBase):
    """Represents a video element on a slide."""

    video: Video
    type: ElementKind = Field(
        default=ElementKind.VIDEO, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.VIDEO

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a VideoElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)

        if self.video.source is None:
            raise ValueError("Video source type is required")

        request = CreateVideoRequest(
            elementProperties=element_properties,
            source=self.video.source.value,
            id=self.video.id,
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a VideoElement to an update request for the Google Slides API."""
        requests = self.alt_text_update_request(element_id)

        if (
            hasattr(self.video, "videoProperties")
            and self.video.videoProperties is not None
        ):
            video_properties = self.video.videoProperties.to_api_format()
            video_request = UpdateVideoPropertiesRequest(
                objectId=element_id,
                videoProperties=video_properties,
                fields=",".join(dict_to_dot_separated_field_list(video_properties)),
            )
            requests.append(video_request)

        return requests


class LineElement(PageElementBase):
    """Represents a line element on a slide."""

    line: Line
    type: ElementKind = Field(
        default=ElementKind.LINE, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.LINE

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a LineElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        request = CreateLineRequest(
            elementProperties=element_properties,
            lineCategory=self.line.lineType if self.line.lineType else "STRAIGHT",
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a LineElement to an update request for the Google Slides API."""
        requests = self.alt_text_update_request(element_id)

        if (
            hasattr(self.line, "lineProperties")
            and self.line.lineProperties is not None
        ):
            line_properties = self.line.lineProperties.to_api_format()
            line_request = UpdateLinePropertiesRequest(
                objectId=element_id,
                lineProperties=line_properties,
                fields=",".join(dict_to_dot_separated_field_list(line_properties)),
            )
            requests.append(line_request)

        return requests


class WordArtElement(PageElementBase):
    """Represents a word art element on a slide."""

    wordArt: WordArt
    type: ElementKind = Field(
        default=ElementKind.WORD_ART,
        description="The type of page element",
        exclude=True,
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.WORD_ART

    # def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
    #     """Convert a WordArtElement to a create request for the Google Slides API."""
    #     element_properties = self.element_properties(parent_id)
    #
    #     if not self.wordArt.renderedText:
    #         raise ValueError("Rendered text is required for Word Art")
    #
    #     request = CreateWordArtRequest(
    #         elementProperties=element_properties,
    #         renderedText=self.wordArt.renderedText,
    #     )
    #     return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a WordArtElement to an update request for the Google Slides API."""
        # WordArt doesn't have specific properties to update beyond base properties
        return self.alt_text_update_request(element_id)


class SheetsChartElement(PageElementBase):
    """Represents a sheets chart element on a slide."""

    sheetsChart: SheetsChart
    type: ElementKind = Field(
        default=ElementKind.SHEETS_CHART,
        description="The type of page element",
        exclude=True,
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.SHEETS_CHART

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a SheetsChartElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)

        if not self.sheetsChart.spreadsheetId or not self.sheetsChart.chartId:
            raise ValueError(
                "Spreadsheet ID and Chart ID are required for Sheets Chart"
            )

        request = CreateSheetsChartRequest(
            elementProperties=element_properties,
            spreadsheetId=self.sheetsChart.spreadsheetId,
            chartId=self.sheetsChart.chartId,
        )
        return [request]

    # def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
    #     """Convert a SheetsChartElement to an update request for the Google Slides API."""
    #     requests = self.alt_text_update_request(element_id)
    #
    #     if (
    #         hasattr(self.sheetsChart, "sheetsChartProperties")
    #         and self.sheetsChart.sheetsChartProperties is not None
    #     ):
    #         chart_properties = self.sheetsChart.sheetsChartProperties.to_api_format()
    #         chart_request = UpdateSheetsChartPropertiesRequest(
    #             objectId=element_id,
    #             sheetsChartProperties=chart_properties,
    #             fields=",".join(dict_to_dot_separated_field_list(chart_properties)),
    #         )
    #         requests.append(chart_request)
    #
    #     return requests


class SpeakerSpotlightElement(PageElementBase):
    """Represents a speaker spotlight element on a slide."""

    speakerSpotlight: SpeakerSpotlight
    type: ElementKind = Field(
        default=ElementKind.SPEAKER_SPOTLIGHT,
        description="The type of page element",
        exclude=True,
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.SPEAKER_SPOTLIGHT

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a SpeakerSpotlightElement to a create request for the Google Slides API."""
        # Note: Speaker spotlight creation is not directly supported in the API
        # This is a placeholder implementation
        raise NotImplementedError(
            "Speaker spotlight creation is not supported by the Google Slides API"
        )

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a SpeakerSpotlightElement to an update request for the Google Slides API."""
        # Speaker spotlight updates are not directly supported
        return self.alt_text_update_request(element_id)


class GroupElement(PageElementBase):
    """Represents a group element on a slide."""

    elementGroup: Group
    type: ElementKind = Field(
        default=ElementKind.GROUP, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.GROUP

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a GroupElement to a create request for the Google Slides API."""
        # Note: Group creation is typically done by grouping existing elements
        # This is a placeholder implementation
        raise NotImplementedError(
            "Group creation should be done by grouping existing elements"
        )

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a GroupElement to an update request for the Google Slides API."""
        # Groups don't have specific properties to update beyond base properties
        return self.alt_text_update_request(element_id)


# Create the discriminated union type
PageElement = Annotated[
    Union[
        Annotated[ShapeElement, Tag("shape")],
        Annotated[TableElement, Tag("table")],
        Annotated[ImageElement, Tag("image")],
        Annotated[VideoElement, Tag("video")],
        Annotated[LineElement, Tag("line")],
        Annotated[WordArtElement, Tag("wordArt")],
        Annotated[SheetsChartElement, Tag("sheetsChart")],
        Annotated[SpeakerSpotlightElement, Tag("speakerSpotlight")],
        Annotated[GroupElement, Tag("group")],
    ],
    Discriminator(element_discriminator),
]

# Rebuild models to resolve forward references
Group.model_rebuild()
GroupElement.model_rebuild()
