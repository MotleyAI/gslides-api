from typing import Annotated, Any, List, Union

from pydantic import Discriminator, Field, Tag, field_validator

from gslides_api.domain.domain import Group, Line, SheetsChart, SpeakerSpotlight, Video, WordArt
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.image import ImageElement
from gslides_api.element.shape import ShapeElement
from gslides_api.element.table import TableElement
from gslides_api.request.request import (  # UpdateSheetsChartPropertiesRequest,; CreateWordArtRequest,
    CreateLineRequest,
    CreateSheetsChartRequest,
    CreateVideoRequest,
    UpdateLinePropertiesRequest,
    UpdateVideoPropertiesRequest,
)
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.utils import dict_to_dot_separated_field_list


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

        if hasattr(self.video, "videoProperties") and self.video.videoProperties is not None:
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

        if hasattr(self.line, "lineProperties") and self.line.lineProperties is not None:
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
            raise ValueError("Spreadsheet ID and Chart ID are required for Sheets Chart")

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
        raise NotImplementedError("Group creation should be done by grouping existing elements")

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
