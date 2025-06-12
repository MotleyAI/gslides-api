from typing import Optional, Dict, Any, List, Union, Annotated
from enum import Enum

from pydantic import Field, Discriminator, Tag

from gslides_api.domain import (
    GSlidesBaseModel,
    Transform,
    Shape,
    Table,
    Image,
    Size,
    Video,
    Line,
    WordArt,
    SheetsChart,
    SpeakerSpotlight,
    Group,
    ImageReplaceMethod,
)
from gslides_api.execute import batch_update
from gslides_api.utils import dict_to_dot_separated_field_list, image_url_is_valid
from gslides_templater import MarkdownProcessor

markdown_processor = MarkdownProcessor()


class ElementKind(Enum):
    """Enumeration of possible page element kinds based on the Google Slides API.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages#pageelement
    """

    GROUP = "elementGroup"
    SHAPE = "shape"
    IMAGE = "image"
    VIDEO = "video"
    LINE = "line"
    TABLE = "table"
    WORD_ART = "wordArt"
    SHEETS_CHART = "sheetsChart"
    SPEAKER_SPOTLIGHT = "speakerSpotlight"


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


class PageElementBase(GSlidesBaseModel):
    """Base class for all page elements."""

    objectId: str
    size: Size
    transform: Transform
    title: Optional[str] = None
    description: Optional[str] = None
    # Store the presentation ID for reference but exclude from model_dump
    presentation_id: Optional[str] = Field(default=None, exclude=True)

    def create_copy(self, parent_id: str, presentation_id: str):
        request = self.create_request(parent_id)
        out = batch_update(request, presentation_id)
        try:
            request_type = list(out["replies"][0].keys())[0]
            new_element_id = out["replies"][0][request_type]["objectId"]
            return new_element_id
        except:
            return None

    def element_properties(self, parent_id: str) -> Dict[str, Any]:
        """Get common element properties for API requests."""
        # Common element properties
        element_properties = {
            "pageObjectId": parent_id,
            "size": self.size.to_api_format(),
            "transform": self.transform.to_api_format(),
        }

        # Add title and description if provided
        if self.title is not None:
            element_properties["title"] = self.title
        if self.description is not None:
            element_properties["description"] = self.description

        return element_properties

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to a create request for the Google Slides API.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement create_request method")

    def update(self, presentation_id: str, element_id: Optional[str] = None) -> Dict[str, Any]:
        if element_id is None:
            element_id = self.objectId
        requests = self.element_to_update_request(element_id)
        if len(requests):
            out = batch_update(requests, presentation_id)
            return out
        else:
            return {}

    def get_base_update_requests(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """

        # Update title and description if provided
        requests = []
        if self.title is not None or self.description is not None:
            update_fields = []
            properties = {}

            if self.title is not None:
                properties["title"] = self.title
                update_fields.append("title")

            if self.description is not None:
                properties["description"] = self.description
                update_fields.append("description")

            if update_fields:
                requests.append(
                    {
                        "updatePageElementProperties": {
                            "objectId": element_id,
                            "pageElementProperties": properties,
                            "fields": ",".join(update_fields),
                        }
                    }
                )
        return requests

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to an update request for the Google Slides API.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement element_to_update_request method")

    def to_markdown(self) -> str | None:
        """Convert a PageElement to markdown.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement to_markdown method")


class ShapeElement(PageElementBase):
    """Represents a shape element on a slide."""

    shape: Shape

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        return [
            {
                "createShape": {
                    "elementProperties": element_properties,
                    "shapeType": self.shape.shapeType.value,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """

        # Update title and description if provided
        requests = self.get_base_update_requests(element_id)

        # shape_properties = self.shape.shapeProperties.to_api_format()
        ## TODO: fix the below, now causes error
        # b'{\n  "error": {\n    "code": 400,\n    "message": "Invalid requests[0].updateShapeProperties: Updating shapeBackgroundFill propertyState to INHERIT is not supported for shape with no placeholder parent shape",\n    "status": "INVALID_ARGUMENT"\n  }\n}\n'
        # out = [
        #     {
        #         "updateShapeProperties": {
        #             "objectId": element_id,
        #             "shapeProperties": shape_properties,
        #             "fields": "*",
        #         }
        #     }
        # ]
        shape_requests = []
        if self.shape.text is None:
            return requests
        for te in self.shape.text.textElements:
            if te.textRun is None:
                # TODO: What is the role of empty ParagraphMarkers?
                continue

            style = te.textRun.style.to_api_format()
            shape_requests += [
                {
                    "insertText": {
                        "objectId": element_id,
                        "text": te.textRun.content,
                        "insertionIndex": te.startIndex,
                    }
                },
                {
                    "updateTextStyle": {
                        "objectId": element_id,
                        "textRange": {
                            "startIndex": te.startIndex or 0,
                            "endIndex": te.endIndex,
                            "type": "FIXED_RANGE",
                        },
                        "style": style,
                        "fields": "*",
                    }
                },
            ]

        return requests + shape_requests

    def to_markdown(self) -> str | None:
        # TODO: make an implementation that doesn't suck
        if hasattr(self.shape, "text") and self.shape.text is not None:
            if self.shape.text.textElements:
                out = []
                for te in self.shape.text.textElements:
                    if te.textRun is not None:
                        out.append(te.textRun.content)
                return "".join(out)
            else:
                return None

    def _write_plain_text_requests(self, text: str):
        raise NotImplementedError("Writing plain text to shape elements is not supported yet")

    def _write_markdown_requests(self, markdown: str):
        requests = markdown_processor.create_slides_requests(self.objectId, markdown)
        return requests

    def write_text(self, text: str, as_markdown: bool = True):
        if as_markdown:
            requests = self._write_markdown_requests(text)
        else:
            requests = self._write_plain_text_requests(text)
        return batch_update(requests, self.presentation_id)


class TableElement(PageElementBase):
    """Represents a table element on a slide."""

    table: Table

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a TableElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        return [
            {
                "createTable": {
                    "elementProperties": element_properties,
                    "rows": self.table.rows,
                    "columns": self.table.columns,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a TableElement to an update request for the Google Slides API."""
        # Tables don't have specific properties to update beyond base properties
        return self.get_base_update_requests(element_id)


class ImageElement(PageElementBase):
    """Represents an image element on a slide."""

    image: Image

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert an ImageElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        return [
            {
                "createImage": {
                    "elementProperties": element_properties,
                    "url": self.image.contentUrl,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert an ImageElement to an update request for the Google Slides API."""
        requests = self.get_base_update_requests(element_id)

        if hasattr(self.image, "imageProperties") and self.image.imageProperties is not None:
            image_properties = self.image.imageProperties.to_api_format()
            # "fields": "*" causes an error
            image_requests = [
                {
                    "updateImageProperties": {
                        "objectId": element_id,
                        "imageProperties": image_properties,
                        "fields": ",".join(dict_to_dot_separated_field_list(image_properties)),
                    }
                }
            ]
            return requests + image_requests

        return requests

    def to_markdown(self) -> str | None:
        url = self.image.sourceUrl
        if url is None:
            return None
        description = self.title or "Image"
        return f"![{description}]({url})"

    def _replace_image_requests(self, new_url: str, method: ImageReplaceMethod | None = None):
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

        request = {"replaceImage": {"imageObjectId": self.objectId, "url": new_url}}
        if method is not None:
            request["replaceImage"]["imageReplaceMethod"] = method.value
        return [request]

    def replace_image(self, new_url: str, method: ImageReplaceMethod | None = None):
        requests = self._replace_image_requests(new_url, method)
        return batch_update(requests, self.presentation_id)


class VideoElement(PageElementBase):
    """Represents a video element on a slide."""

    video: Video

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a VideoElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)

        if self.video.source is None:
            raise ValueError("Video source type is required")

        return [
            {
                "createVideo": {
                    "elementProperties": element_properties,
                    "source": self.video.source.value,
                    "id": self.video.id,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a VideoElement to an update request for the Google Slides API."""
        requests = self.get_base_update_requests(element_id)

        if hasattr(self.video, "videoProperties") and self.video.videoProperties is not None:
            video_properties = self.video.videoProperties.to_api_format()
            video_requests = [
                {
                    "updateVideoProperties": {
                        "objectId": element_id,
                        "videoProperties": video_properties,
                        "fields": ",".join(dict_to_dot_separated_field_list(video_properties)),
                    }
                }
            ]
            return requests + video_requests

        return requests


class LineElement(PageElementBase):
    """Represents a line element on a slide."""

    line: Line

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a LineElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)
        return [
            {
                "createLine": {
                    "elementProperties": element_properties,
                    "lineCategory": self.line.lineType if self.line.lineType else "STRAIGHT",
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a LineElement to an update request for the Google Slides API."""
        requests = self.get_base_update_requests(element_id)

        if hasattr(self.line, "lineProperties") and self.line.lineProperties is not None:
            line_properties = self.line.lineProperties.to_api_format()
            line_requests = [
                {
                    "updateLineProperties": {
                        "objectId": element_id,
                        "lineProperties": line_properties,
                        "fields": ",".join(dict_to_dot_separated_field_list(line_properties)),
                    }
                }
            ]
            return requests + line_requests

        return requests


class WordArtElement(PageElementBase):
    """Represents a word art element on a slide."""

    wordArt: WordArt

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a WordArtElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)

        if not self.wordArt.renderedText:
            raise ValueError("Rendered text is required for Word Art")

        return [
            {
                "createWordArt": {
                    "elementProperties": element_properties,
                    "renderedText": self.wordArt.renderedText,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a WordArtElement to an update request for the Google Slides API."""
        # WordArt doesn't have specific properties to update beyond base properties
        return self.get_base_update_requests(element_id)


class SheetsChartElement(PageElementBase):
    """Represents a sheets chart element on a slide."""

    sheetsChart: SheetsChart

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a SheetsChartElement to a create request for the Google Slides API."""
        element_properties = self.element_properties(parent_id)

        if not self.sheetsChart.spreadsheetId or not self.sheetsChart.chartId:
            raise ValueError("Spreadsheet ID and Chart ID are required for Sheets Chart")

        return [
            {
                "createSheetsChart": {
                    "elementProperties": element_properties,
                    "spreadsheetId": self.sheetsChart.spreadsheetId,
                    "chartId": self.sheetsChart.chartId,
                }
            }
        ]

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a SheetsChartElement to an update request for the Google Slides API."""
        requests = self.get_base_update_requests(element_id)

        if (
            hasattr(self.sheetsChart, "sheetsChartProperties")
            and self.sheetsChart.sheetsChartProperties is not None
        ):
            chart_properties = self.sheetsChart.sheetsChartProperties.to_api_format()
            chart_requests = [
                {
                    "updateSheetsChartProperties": {
                        "objectId": element_id,
                        "sheetsChartProperties": chart_properties,
                        "fields": ",".join(dict_to_dot_separated_field_list(chart_properties)),
                    }
                }
            ]
            return requests + chart_requests

        return requests


class SpeakerSpotlightElement(PageElementBase):
    """Represents a speaker spotlight element on a slide."""

    speakerSpotlight: SpeakerSpotlight

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a SpeakerSpotlightElement to a create request for the Google Slides API."""
        # Note: Speaker spotlight creation is not directly supported in the API
        # This is a placeholder implementation
        raise NotImplementedError(
            "Speaker spotlight creation is not supported by the Google Slides API"
        )

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a SpeakerSpotlightElement to an update request for the Google Slides API."""
        # Speaker spotlight updates are not directly supported
        return self.get_base_update_requests(element_id)


class GroupElement(PageElementBase):
    """Represents a group element on a slide."""

    elementGroup: Group

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a GroupElement to a create request for the Google Slides API."""
        # Note: Group creation is typically done by grouping existing elements
        # This is a placeholder implementation
        raise NotImplementedError("Group creation should be done by grouping existing elements")

    def element_to_update_request(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a GroupElement to an update request for the Google Slides API."""
        # Groups don't have specific properties to update beyond base properties
        return self.get_base_update_requests(element_id)


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
