import logging
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

from pydantic import Field

from gslides_api.client import GoogleAPIClient, api_client
from gslides_api.domain.domain import (
    GSlidesBaseModel,
    OutputUnit,
    PageElementProperties,
    Size,
    Transform,
)
from gslides_api.request.request import UpdatePageElementAltTextRequest
from gslides_api.request.parent import GSlidesAPIRequest

logger = logging.getLogger(__name__)


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


class AltText(GSlidesBaseModel):
    title: Optional[str] = None
    description: Optional[str] = None


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
    slide_id: Optional[str] = Field(default=None, exclude=True)

    # EMU conversion constants
    _EMU_PER_CM = 360000  # 1 EMU = 1/360,000 cm
    _EMU_PER_INCH = 914400  # 1 inch = 914,400 EMUs

    def _convert_emu_to_units(self, value_emu: float, units: OutputUnit) -> float:
        """Convert a value from EMUs to the specified units.

        Args:
            value_emu: The value in EMUs to convert.
            units: The target units (OutputUnit.CM or OutputUnit.IN).

        Returns:
            The converted value in the specified units.

        Raises:
            TypeError: If units is not an OutputUnit enum value.
        """
        try:
            units = OutputUnit(units)
        except Exception as e:
            raise TypeError(f"units must be an OutputUnit enum value, got {units}") from e

        if not isinstance(units, OutputUnit):
            raise TypeError(f"units must be an OutputUnit enum value, got {type(units)}")

        if units == OutputUnit.CM:
            return value_emu / self._EMU_PER_CM
        elif units == OutputUnit.IN:
            return value_emu / self._EMU_PER_INCH
        else:
            raise ValueError(f"Unsupported OutputUnit: {units}")

    def create_copy(
        self,
        parent_id: str,
        presentation_id: str,
        api_client: Optional[GoogleAPIClient] = None,
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

    def delete(self, api_client: Optional[GoogleAPIClient] = None) -> None:
        assert (
            self.presentation_id is not None
        ), "self.presentation_id must be set when calling delete()"
        client = api_client or globals()["api_client"]
        client.delete_object(self.objectId, self.presentation_id)

    def element_properties(self, parent_id: str) -> PageElementProperties:
        """Get common element properties for API requests."""
        # Common element properties
        element_properties = {
            "pageObjectId": parent_id,
            "size": self.size.to_api_format(),
            "transform": self.transform.to_api_format(),
        }

        # TODO: this will be ignored - where are they set?
        # # Add title and description if provided
        # if self.title is not None:
        #     element_properties["title"] = self.title
        # if self.description is not None:
        #     element_properties["description"] = self.description

        return PageElementProperties.model_validate(element_properties)

    @classmethod
    def from_ids(
        cls,
        presentation_id: str,
        slide_id: str,
        element_id: str,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "PageElementBase":
        from gslides_api.page.slide import Slide

        slide = Slide.from_ids(presentation_id, slide_id, api_client=api_client)
        return slide.get_element_by_id(element_id)

    def sync_from_cloud(self, api_client: Optional[GoogleAPIClient] = None) -> None:
        new_state = PageElementBase.from_ids(
            self.presentation_id, self.slide_id, self.objectId, api_client=api_client
        )
        self.__dict__ = new_state.__dict__

    def alt_text_update_request(
        self, element_id: str, title: str | None = None, description: str | None = None
    ) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """
        if (
            title is not None
            or description is not None
            or self.title is not None
            or self.description is not None
        ):
            return [
                UpdatePageElementAltTextRequest(
                    objectId=element_id,
                    title=title if title is not None else self.title,
                    description=(description if description is not None else self.description),
                )
            ]
        else:
            return []

    def set_alt_text(
        self,
        title: str | None = None,
        description: str | None = None,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        client = api_client or globals()["api_client"]
        if not title and not description:
            logger.warning(
                "No alt text provided, skipping update. \n "
                "Remember that Google Slides API won't allow to write empty strings."
            )
            return
        client.batch_update(
            self.alt_text_update_request(
                title=title, description=description, element_id=self.objectId
            ),
            self.presentation_id,
        )

    @property
    def alt_text(self):
        # Don't provide a setter as want to also pass api_client to the setter
        return AltText(title=self.title, description=self.description)

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

    def absolute_size(self, units: OutputUnit) -> Tuple[float, float]:
        """Calculate the absolute size of the element in the specified units.

        This method calculates the actual rendered size of the element, taking into
        account any scaling applied via the transform. The size represents the
        width and height of the element as it appears on the slide.

        Args:
            units: The units to return the size in. Can be "cm" or "in".

        Returns:
            A tuple of (width, height) representing the element's dimensions
            in the specified units.

        Raises:
            ValueError: If units is not "cm" or "in".
            ValueError: If element size is not available.
        """

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
        width_result = self._convert_emu_to_units(actual_width_emu, units)
        height_result = self._convert_emu_to_units(actual_height_emu, units)

        return width_result, height_result

    def absolute_position(self, units: OutputUnit = OutputUnit.CM) -> Tuple[float, float]:
        """Calculate the absolute position of the element on the page in the specified units.

        Position represents the distance of the top-left corner of the element
        from the top-left corner of the slide.

        Args:
            units: The units to return the position in. Can be "cm" or "in".

        Returns:
            A tuple of (x, y) representing the position in the specified units,
            where x is the horizontal distance from the left edge and y is the
            vertical distance from the top edge of the slide.
        """

        # Extract position from transform (translateX, translateY are in EMUs)
        x_emu = self.transform.translateX
        y_emu = self.transform.translateY

        # Convert from EMUs to the requested units
        x_result = self._convert_emu_to_units(x_emu, units)
        y_result = self._convert_emu_to_units(y_emu, units)

        return x_result, y_result
