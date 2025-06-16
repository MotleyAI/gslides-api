from typing import Any, Dict, List

from pydantic import Field, field_validator


from gslides_api.domain import Shape, TextElement, TextStyle
from gslides_api.element.base import PageElementBase, ElementKind
from gslides_api.execute import batch_update
from gslides_api.markdown import markdown_to_text_elements


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
        if self.shape.text is not None:
            requests += text_elements_to_requests(self.shape.text.textElements, element_id)

        return requests

    def _delete_text_request(self):
        return [{"deleteText": {"objectId": self.objectId, "textRange": {"type": "ALL"}}}]

    def delete_text(self):
        return batch_update(self._delete_text_request(), self.presentation_id)

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

    def _write_plain_text_requests(self, text: str, style: TextStyle | None = None):
        raise NotImplementedError("Writing plain text to shape elements is not supported yet")

    def _write_markdown_requests(self, markdown: str, style: TextStyle | None = None):
        elements = markdown_to_text_elements(markdown, base_style=style)
        requests = text_elements_to_requests(elements, self.objectId)
        return requests

    @property
    def style(self):
        if not hasattr(self.shape, "text") or self.shape.text is None:
            return None
        if not hasattr(self.shape.text, "textElements") or not self.shape.text.textElements:
            return None
        for te in self.shape.text.textElements:
            if te.textRun is None:
                continue
            return te.textRun.style

    def write_text(self, text: str, as_markdown: bool = True):
        requests = self._delete_text_request()
        if as_markdown:
            requests += self._write_markdown_requests(text, style=self.style)
        else:
            requests += self._write_plain_text_requests(text, style=self.style)
        return batch_update(requests, self.presentation_id)


def text_elements_to_requests(text_elements: List[TextElement], objectId: str):
    requests = []
    for te in text_elements:
        if te.textRun is None:
            # TODO: What is the role of empty ParagraphMarkers?
            continue

        style = te.textRun.style.to_api_format()
        requests += [
            {
                "insertText": {
                    "objectId": objectId,
                    "text": te.textRun.content,
                    "insertionIndex": te.startIndex,
                }
            },
            {
                "updateTextStyle": {
                    "objectId": objectId,
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
    return requests
