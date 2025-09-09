from typing import List, Optional

from pydantic import Field, field_validator

from gslides_api.client import GoogleAPIClient
from gslides_api.client import api_client as default_api_client
from gslides_api.domain.domain import (
    Dimension,
    GSlidesBaseModel,
    OutputUnit,
    PageElementProperties,
    Size,
    Transform,
    Unit,
)
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.text_content import TextContent
from gslides_api.markdown.element import MarkdownTextElement as MarkdownTextElement
from gslides_api.markdown.from_markdown import markdown_to_text_elements, text_elements_to_requests
from gslides_api.markdown.to_markdown import text_elements_to_markdown
from gslides_api.domain.request import Range, RangeType
from gslides_api.request.request import (
    CreateShapeRequest,
    DeleteParagraphBulletsRequest,
    DeleteTextRequest,
    UpdateTextStyleRequest,
)
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.domain.text import Placeholder, ShapeProperties, TextStyle, Type, Type as ShapeType


class Shape(GSlidesBaseModel):
    """Represents a shape in a slide."""

    shapeProperties: ShapeProperties
    shapeType: Optional[Type] = None  # Make optional to preserve original JSON exactly
    text: Optional[TextContent] = None
    placeholder: Optional[Placeholder] = None


class ShapeElement(PageElementBase):
    """Represents a shape element on a slide."""

    shape: Shape
    type: ElementKind = Field(
        default=ElementKind.SHAPE, description="The type of page element", exclude=True
    )

    @field_validator("type")
    @classmethod
    def validate_type(cls, v):
        return ElementKind.SHAPE

    def create_request(self, parent_id: str) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to a create request for the Google Slides API."""
        element_props = self.element_properties(parent_id)

        request = CreateShapeRequest(
            elementProperties=element_props, shapeType=self.shape.shapeType
        )
        return [request]

    def element_to_update_request(self, element_id: str) -> List[GSlidesAPIRequest]:
        """Convert a PageElement to an update request for the Google Slides API.
        :param element_id: The id of the element to update, if not the same as e objectId
        :type element_id: str, optional
        :return: The update request
        :rtype: list

        """

        # Update title and description if provided
        requests: list[GSlidesAPIRequest] = self.alt_text_update_request(element_id)

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
        if self.shape.text is not None:
            text_requests = text_elements_to_requests(
                self.shape.text.textElements, [], objectId=element_id
            )
            requests.extend(text_requests[0])

        return requests

    def delete_text_request(self) -> List[GSlidesAPIRequest]:
        return self.shape.text.delete_text_request(self.objectId)

    def delete_text(self, api_client: Optional[GoogleAPIClient] = None):
        client = api_client or default_api_client
        return client.batch_update(self.delete_text_request(), self.presentation_id)

    @property
    def styles(self) -> List[TextStyle] | None:
        if not hasattr(self.shape, "text") or self.shape.text is None:
            return None

        return self.shape.text.styles

    def to_markdown(self) -> str | None:
        """Convert the shape's text content back to markdown format.

        This method reconstructs markdown from the Google Slides API response,
        handling formatting like bold, italic, bullet points, nested lists, and code spans.
        """
        if self.shape.text is None:
            return None

        return self.shape.text.to_markdown()

    @property
    def has_text(self):
        return self.shape.text is not None and self.shape.text.has_text

    def write_text(
        self,
        text: str,
        as_markdown: bool = True,
        styles: List[TextStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        size_inches = self.absolute_size(OutputUnit.IN)
        requests = self.shape.text.write_text_requests(
            text=text,
            as_markdown=as_markdown,
            styles=styles,
            overwrite=overwrite,
            autoscale=autoscale,
            size_inches=size_inches,
        )

        for r in requests:
            r.objectId = self.objectId

        if requests:
            client = api_client or default_api_client
            return client.batch_update(requests, self.presentation_id)

    def read_text(self, as_markdown: bool = True):
        if not self.has_text:
            return ""
        return self.shape.text.read_text(as_markdown)

    def to_markdown_element(self, name: str = "Text") -> MarkdownTextElement:
        """Convert ShapeElement to MarkdownTextElement for round-trip conversion."""
        content = self.to_markdown() or ""

        # Store position, size, and other properties in metadata for perfect reconstruction
        metadata = {
            "objectId": self.objectId,
            "shape_type": self.shape.shapeType.value if self.shape.shapeType else None,
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

        # Store text styles if available
        if self.styles:
            metadata["styles"] = [
                style.to_api_format() if hasattr(style, "to_api_format") else str(style)
                for style in self.styles
            ]

        return MarkdownTextElement(name=name, content=content, metadata=metadata)

    @classmethod
    def from_markdown_element(
        cls,
        markdown_elem: MarkdownTextElement,
        parent_id: str,
        shape_type: Optional[str] = None,
        api_client: Optional[GoogleAPIClient] = None,
    ) -> "ShapeElement":
        """Create ShapeElement from MarkdownTextElement with preserved metadata."""

        # Extract metadata
        metadata = markdown_elem.metadata or {}
        object_id = metadata.get("objectId")
        stored_shape_type = metadata.get("shape_type") or shape_type or "TEXT_BOX"

        # Create basic shape with text content
        from gslides_api.domain.text import ShapeProperties
        from gslides_api.element.text_content import TextContent

        # Create a minimal shape - the actual content will be written via write_text
        shape = Shape(
            shapeProperties=ShapeProperties(),
            shapeType=ShapeType(stored_shape_type),
            text=TextContent(textElements=[]) if markdown_elem.content.strip() else None,
        )

        # Create element properties from metadata
        element_props = PageElementProperties(pageObjectId=parent_id)

        # Restore size if available, otherwise provide default
        if "size" in metadata:
            size_data = metadata["size"]
            element_props.size = Size(
                width=Dimension(magnitude=size_data["width"], unit=Unit(size_data["unit"])),
                height=Dimension(magnitude=size_data["height"], unit=Unit(size_data["unit"])),
            )
        else:
            element_props.size = Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=200, unit=Unit.PT),
            )

        # Restore transform if available, otherwise create default
        if "transform" in metadata and metadata["transform"]:

            transform_data = metadata["transform"]
            element_props.transform = Transform(**transform_data)
        else:
            element_props.transform = Transform(
                scaleX=1.0, scaleY=1.0, translateX=0.0, translateY=0.0, unit="EMU"
            )

        # Create the shape element
        shape_element = cls(
            objectId=object_id or "shape_" + str(hash(markdown_elem.content))[:8],
            size=element_props.size,
            transform=element_props.transform,
            title=metadata.get("title"),
            description=metadata.get("description"),
            shape=shape,
            slide_id=parent_id,
            presentation_id="",  # Will need to be set by caller
        )

        return shape_element
