from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api.domain import GSlidesBaseModel, Transform, Size
from gslides_api.execute import api_client


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


class PageElementBase(GSlidesBaseModel):
    """Base class for all page elements."""

    objectId: str
    size: Size
    transform: Transform
    title: Optional[str] = None
    description: Optional[str] = None
    type: ElementKind = Field(description="The type of page element", exclude=True)
    # Store the presentation ID for reference but exclude from model_dump
    presentation_id: Optional[str] = Field(default=None, exclude=True)

    def create_copy(self, parent_id: str, presentation_id: str):
        request = self.create_request(parent_id)
        out = api_client.batch_update(request, presentation_id)
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

    def get_base_update_requests(self, element_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """

        # Update title and description if provided, they're both alt text properties
        requests = []
        if self.title is not None or self.description is not None:
            properties = {}

            if self.title is not None:
                properties["title"] = self.title

            if self.description is not None:
                properties["description"] = self.description

            # TODO: add all the allowed fields!
            if properties:
                requests.append(
                    {
                        "updatePageElementAltText": {
                            "objectId": element_id,
                            **properties,
                        }
                    }
                )
        return requests

    def create_request(self, parent_id: str) -> List[Dict[str, Any]]:
        """Convert a PageElement to a create request for the Google Slides API.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement create_request method")

    def update(
        self, element_id: Optional[str] = None, presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        if element_id is None:
            element_id = self.objectId

        if presentation_id is None:
            presentation_id = self.presentation_id

        requests = self.element_to_update_request(element_id)
        if len(requests):
            out = api_client.batch_update(requests, presentation_id)
            return out
        else:
            return {}

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
