from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

from pydantic import Field

from gslides_api.domain import GSlidesBaseModel, Transform, Size
from gslides_api.request.request import GSlidesAPIRequest, UpdatePageElementAltTextRequest
from gslides_api.client import api_client, GoogleAPIClient


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
    size: Optional[Size] = None
    transform: Transform
    title: Optional[str] = None
    description: Optional[str] = None
    type: ElementKind = Field(description="The type of page element", exclude=True)
    # Store the presentation ID for reference but exclude from model_dump
    presentation_id: Optional[str] = Field(default=None, exclude=True)

    def create_copy(
        self, parent_id: str, presentation_id: str, api_client: Optional[GoogleAPIClient] = None
    ):
        client = api_client or globals()["api_client"]
        request = self.create_request(parent_id)
        out = client.batch_update(request, presentation_id)
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

    def alt_text_update_request(
        self, element_id: str, title: str | None = None, description: str | None = None
    ) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """

        if self.title is not None or self.description is not None:
            return [
                UpdatePageElementAltTextRequest(
                    objectId=element_id,
                    title=title or self.title,
                    description=description or self.description,
                )
            ]
        else:
            return []

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to a create request for the Google Slides API.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement create_request method")

    def update(
        self,
        element_id: Optional[str] = None,
        presentation_id: Optional[str] = None,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> Dict[str, Any]:
        if element_id is None:
            element_id = self.objectId

        if presentation_id is None:
            presentation_id = self.presentation_id

        client = api_client or globals()["api_client"]
        request_objects = self.element_to_update_request(element_id)
        if len(request_objects):
            out = client.batch_update(request_objects, presentation_id)
            return out
        else:
            return {}

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to an update request for the Google Slides API.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement element_to_update_request method")

    def to_markdown(self) -> str | None:
        """Convert a PageElement to markdown.

        This method should be overridden by subclasses.
        """
        raise NotImplementedError("Subclasses must implement to_markdown method")

    def set_alt_text(
        self, title: str, description: str, api_client: Optional[GoogleAPIClient] = None
    ):
        client = api_client or globals()["api_client"]
        client.batch_update(
            self.alt_text_update_request(
                title=title, description=description, element_id=self.objectId
            ),
            self.presentation_id,
        )

    def absolute_size(self, units: str = "in") -> Tuple[float, float]:
        """Calculate the absolute size of the element in the specified units.

        Args:
            units: The units to return the size in. Can be "cm" or "in".

        Returns:
            A tuple of (width, height) in the specified units.

        Raises:
            ValueError: If units is not "cm" or "in".
            ValueError: If size is None.
        """
        if units not in ["cm", "in"]:
            raise ValueError("Units must be 'cm' or 'in'")

        if self.size is None:
            raise ValueError("Element size is not available")

        # Extract width and height from size
        # Size can have width/height as either float or Dimension objects
        if hasattr(self.size.width, "magnitude"):
            width_emu = self.size.width.magnitude
        else:
            width_emu = self.size.width

        if hasattr(self.size.height, "magnitude"):
            height_emu = self.size.height.magnitude
        else:
            height_emu = self.size.height

        # Apply transform scaling
        actual_width_emu = width_emu * self.transform.scaleX
        actual_height_emu = height_emu * self.transform.scaleY

        # Convert from EMUs to the requested units
        if units == "cm":
            # 1 EMU = 1/360000 cm (as per the instructions)
            width_result = actual_width_emu / 360000
            height_result = actual_height_emu / 360000
        else:  # units == "in"
            # 1 inch = 914400 EMUs (as per the instructions)
            width_result = actual_width_emu / 914400
            height_result = actual_height_emu / 914400

        return (width_result, height_result)
