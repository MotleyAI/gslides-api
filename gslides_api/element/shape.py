from typing import Any, Dict, List, Optional, Optional

from pydantic import Field, field_validator

from gslides_api import TextElement
from gslides_api.text import Shape, TextStyle
from gslides_api.element.base import PageElementBase, ElementKind
from gslides_api.client import api_client, GoogleAPIClient
from gslides_api.markdown import markdown_to_text_elements, text_elements_to_markdown
from gslides_api.request.request import (
    GSlidesAPIRequest,
    InsertTextRequest,
    UpdateTextStyleRequest,
    CreateShapeRequest,
    DeleteTextRequest,
)
from gslides_api.request.domain import Range, RangeType, ElementProperties


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
        element_properties_dict = self.element_properties(parent_id)
        element_props = ElementProperties(
            pageObjectId=element_properties_dict["pageObjectId"],
            size=element_properties_dict.get("size"),
            transform=element_properties_dict.get("transform"),
            title=element_properties_dict.get("title"),
            description=element_properties_dict.get("description"),
        )
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
        requests = self.alt_text_update_request(element_id)

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
            text_requests = text_elements_to_requests(self.shape.text.textElements, element_id)
            requests.extend(text_requests)

        return requests

    def delete_text_request(self):
        return [DeleteTextRequest(objectId=self.objectId, textRange=Range(type=RangeType.ALL))]

    def delete_text(self, api_client: Optional[GoogleAPIClient] = None):
        client = api_client or globals()["api_client"]
        return client.batch_update(self.delete_text_request(), self.presentation_id)

    @property
    def styles(self) -> List[TextStyle] | None:
        if not hasattr(self.shape, "text") or self.shape.text is None:
            return None
        if not hasattr(self.shape.text, "textElements") or not self.shape.text.textElements:
            return None
        styles = []
        for te in self.shape.text.textElements:
            if te.textRun is None:
                continue
            if te.textRun.style not in styles:
                styles.append(te.textRun.style)
        return styles

    def to_markdown(self) -> str | None:
        """Convert the shape's text content back to markdown format.

        This method reconstructs markdown from the Google Slides API response,
        handling formatting like bold, italic, bullet points, nested lists, and code spans.
        """
        if not hasattr(self.shape, "text") or self.shape.text is None:
            return None
        if not hasattr(self.shape.text, "textElements") or not self.shape.text.textElements:
            return None

        elements = self.shape.text.textElements
        return text_elements_to_markdown(elements)

    @property
    def has_text(self):
        return (
            self.shape.text is not None
            and hasattr(self.shape.text, "textElements")
            and len(self.shape.text.textElements) > 0
            and self.shape.text.textElements[0].endIndex > 0
        )

    def write_text(
        self,
        text: str,
        as_markdown: bool = True,
        styles: List[TextStyle] | None = None,
        append: bool = False,
        api_client: Optional[GoogleAPIClient] = None,
    ):
        styles = styles or self.styles
        if self.has_text and not append:
            requests = self.delete_text_request()
        else:
            requests = []
        style_args = {}
        if styles is not None:
            if len(styles) == 1:
                style_args["base_style"] = styles[0]
            elif len(styles) > 1:
                style_args["heading_style"] = styles[0]
                style_args["base_style"] = styles[1]

        elements = markdown_to_text_elements(text, **style_args)
        requests += text_elements_to_requests(elements, self.objectId)

        if not as_markdown:
            requests = [r for r in requests if not isinstance(r, UpdateTextStyleRequest)]

        if requests:
            client = api_client or globals()["api_client"]
            return client.batch_update(requests, self.presentation_id)

    def read_text(self, as_markdown: bool = True):
        if not self.has_text:
            return ""
        if as_markdown:
            return self.to_markdown()
        else:
            out = []
            for te in self.shape.text.textElements:
                if te.textRun is not None:
                    out.append(te.textRun.content)
                elif te.paragraphMarker is not None:
                    if len(out) > 0:
                        out.append("\n")
            return "".join(out)


def text_elements_to_requests(
    text_elements: List[TextElement | GSlidesAPIRequest], objectId: str
) -> List[GSlidesAPIRequest]:
    requests = []
    for te in text_elements:
        if isinstance(te, GSlidesAPIRequest):
            te.objectId = objectId
            requests.append(te)
            continue
        else:
            assert isinstance(te, TextElement), f"Expected TextElement, got {te}"
        if te.textRun is None:
            # An empty text run will have a non-None ParagraphMarker
            # Apparently there's no direct way to insert ParagraphMarkers, instead they have to be created
            # as a side effect of inserting text or by specialized calls like createParagraphBullets
            # So we just ignore them when inserting text
            continue

        # Create InsertTextRequest
        insert_request = InsertTextRequest(
            objectId=objectId,
            text=te.textRun.content,
            insertionIndex=te.startIndex,
        )

        # Create UpdateTextStyleRequest
        text_range = Range(
            type=RangeType.FIXED_RANGE,
            startIndex=te.startIndex or 0,
            endIndex=te.endIndex,
        )
        update_style_request = UpdateTextStyleRequest(
            objectId=objectId,
            textRange=text_range,
            style=te.textRun.style,
            fields="*",
        )

        requests.extend([insert_request, update_style_request])
    return requests
