from typing import List, Optional

from pydantic import Field, field_validator

from gslides_api.agnostic.element import MarkdownTextElement as MarkdownTextElement
from gslides_api.agnostic.text import RichStyle
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
from gslides_api.domain.request import Range, RangeType
from gslides_api.domain.text import PlaceholderType, ShapeProperties, TextStyle
from gslides_api.domain.text import Type
from gslides_api.domain.text import Type as ShapeType
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.text_content import TextContent
from gslides_api.markdown.from_markdown import markdown_to_text_elements, text_elements_to_requests
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.request import (
    CreateShapeRequest,
    DeleteParagraphBulletsRequest,
    DeleteTextRequest,
    UpdateTextStyleRequest,
)


class Placeholder(GSlidesBaseModel):
    """Represents a placeholder in a slide."""

    type: PlaceholderType
    parentObjectId: Optional[str] = None
    index: Optional[int] = None
    # This is not in the API, we fetch it from elsewhere in the Presentation object using parentObjectId
    parent_object: Optional["ShapeElement"] = Field(default=None, exclude=True)


class Shape(GSlidesBaseModel):
    """Represents a shape in a slide."""

    shapeProperties: ShapeProperties
    shapeType: Optional[Type] = None  # Make optional to preserve original JSON exactly
    text: Optional[TextContent] = None
    placeholder: Optional[Placeholder] = None

    @property
    def placeholder_styles(self) -> list[RichStyle]:
        """Get styles from the placeholder's parent object.

        Returns RichStyle objects (non-markdown-renderable properties).
        """
        if self.placeholder is None or self.placeholder.parent_object is None:
            return []
        return self.placeholder.parent_object.styles(skip_whitespace=False)


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

    def styles(self, skip_whitespace: bool = True) -> List[RichStyle] | None:
        """Extract all unique RichStyle objects from the shape's text.

        Returns only RichStyle (non-markdown-renderable properties like colors, fonts).
        Text that differs only in bold/italic/strikethrough is considered ONE style,
        since those properties are stored in the markdown string itself.

        If no styles are found in the text, falls back to placeholder styles.
        """
        if not hasattr(self.shape, "text") or self.shape.text is None:
            styles = None
        else:
            styles = self.shape.text.styles(skip_whitespace)

            if styles is not None and len(styles) == 1 and styles[0].is_default():
                styles = None

        if styles is None:
            styles = self.shape.placeholder_styles

        return styles

    @property
    def has_text(self):
        return self.shape.text is not None and self.shape.text.has_text

    def write_text(
        self,
        text: str,
        as_markdown: bool = True,
        styles: List[RichStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        """Write text to the shape, optionally parsing as markdown.

        Args:
            text: The text content to write (can be markdown if as_markdown=True)
            as_markdown: If True, parse text as markdown and apply formatting
            styles: List of RichStyle objects to apply. If None, uses self.styles().
                   RichStyle contains non-markdown properties (colors, fonts, etc.)
                   Markdown formatting (bold, italic) is derived from parsing the text.
            overwrite: If True, delete existing text before writing
            autoscale: If True, scale font size to fit text in the element
            api_client: Optional client to use for the API call
        """
        size_inches = self.absolute_size(OutputUnit.IN)
        if not self.shape.text:
            self.shape.text = TextContent(textElements=[])

        if not styles:
            styles = self.styles()
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
        content = self.read_text(as_markdown=True) or ""

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
        if self.styles():
            metadata["styles"] = [
                style.to_api_format() if hasattr(style, "to_api_format") else str(style)
                for style in self.styles()
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
            text=(TextContent(textElements=[]) if markdown_elem.content.strip() else None),
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


Placeholder.model_rebuild()
