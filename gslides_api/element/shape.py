from typing import Any, Dict, List

from pydantic import Field, field_validator


from gslides_api.domain import Shape, TextElement, TextStyle
from gslides_api.element.base import PageElementBase, ElementKind
from gslides_api.execute import batch_update
from gslides_api.markdown import markdown_to_text_elements
from gslides_api.requests import GslidesAPIRequest


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
        """Convert the shape's text content back to markdown format.

        This method reconstructs markdown from the Google Slides API response,
        handling formatting like bold, italic, bullet points, nested lists, and code spans.
        """
        if not hasattr(self.shape, "text") or self.shape.text is None:
            return None
        if not hasattr(self.shape.text, "textElements") or not self.shape.text.textElements:
            return None

        result = []
        current_paragraph = []

        # Track list state for proper nesting
        current_list_info = None  # Will store (listId, nestingLevel, glyph, is_numbered)
        pending_bullet_info = None  # Store bullet info until we get the text content

        for i, te in enumerate(self.shape.text.textElements):
            # Handle paragraph markers (for bullets and paragraph breaks)
            if te.paragraphMarker is not None:
                # Check if this is a bullet point
                if te.paragraphMarker.bullet is not None:
                    bullet = te.paragraphMarker.bullet
                    list_id = bullet.listId if hasattr(bullet, "listId") else None
                    nesting_level = bullet.nestingLevel if hasattr(bullet, "nestingLevel") else 0
                    glyph = bullet.glyph if hasattr(bullet, "glyph") and bullet.glyph else "●"

                    # Determine if this is a numbered list based on the glyph
                    is_numbered = self._is_numbered_list_glyph(glyph)

                    # Store the bullet info to be used when we encounter the text content
                    pending_bullet_info = (list_id, nesting_level, glyph, is_numbered)
                    continue
                else:
                    # Regular paragraph marker - clear any pending bullet info
                    pending_bullet_info = None
                    continue

            # Handle text runs
            if te.textRun is not None:
                content = te.textRun.content
                style = te.textRun.style

                # Handle bullet points - add bullet marker at start of line
                if pending_bullet_info and content.strip() and not current_paragraph:
                    list_id, nesting_level, glyph, is_numbered = pending_bullet_info

                    # Generate the appropriate indentation and bullet marker
                    indent = self._get_list_indentation(nesting_level)
                    bullet_marker = self._format_bullet_marker_with_nesting(glyph)

                    current_paragraph.append(indent + bullet_marker)
                    current_list_info = pending_bullet_info
                    pending_bullet_info = None  # Clear after use

                # Apply formatting based on style
                formatted_content = self._apply_markdown_formatting(content, style)
                current_paragraph.append(formatted_content)

                # Handle line breaks
                if "\n" in content:
                    # Join current paragraph and add to result
                    paragraph_text = "".join(current_paragraph).rstrip()
                    if paragraph_text:
                        result.append(paragraph_text)
                    current_paragraph = []

        # Add any remaining paragraph content
        if current_paragraph:
            paragraph_text = "".join(current_paragraph).rstrip()
            if paragraph_text:
                result.append(paragraph_text)

        return "\n".join(result) if result else None

    def _apply_markdown_formatting(self, content: str, style) -> str:
        """Apply markdown formatting to content based on text style."""
        if style is None:
            return content

        # Handle hyperlinks first (they take precedence)
        if hasattr(style, "link") and style.link:
            # Handle both dict and object cases
            url = None
            if isinstance(style.link, dict) and "url" in style.link:
                url = style.link["url"]
            elif hasattr(style.link, "url"):
                url = style.link.url

            if url:
                # For links, format as [text](url)
                clean_content = content.strip()
                if clean_content:
                    return f"[{clean_content}]({url})"
            return content

        # Handle code spans (different font family)
        if (
            hasattr(style, "fontFamily")
            and style.fontFamily
            and style.fontFamily.lower() in ["courier new", "courier", "monospace"]
        ):
            # For code spans, only format the non-whitespace content
            if content.strip():
                return f"`{content.strip()}`"
            return content

        # For formatting, we need to preserve leading/trailing spaces
        # but only format the actual text content
        leading_space = ""
        trailing_space = ""
        text_content = content

        # Extract leading spaces
        for char in content:
            if char in " \t":
                leading_space += char
            else:
                break

        # Extract trailing spaces (but not newlines)
        temp_content = content.rstrip("\n")
        trailing_newlines = content[len(temp_content) :]

        for char in reversed(temp_content):
            if char in " \t":
                trailing_space = char + trailing_space
            else:
                break

        # Get the actual text content without leading/trailing spaces
        text_content = content.strip(" \t").rstrip("\n")

        # Apply formatting only to the text content
        if text_content:
            # Handle strikethrough first (can combine with other formatting)
            if hasattr(style, "strikethrough") and style.strikethrough:
                text_content = f"~~{text_content}~~"

            # Handle combined bold and italic (***text***)
            if hasattr(style, "bold") and style.bold and hasattr(style, "italic") and style.italic:
                text_content = f"***{text_content}***"
            # Handle bold only
            elif hasattr(style, "bold") and style.bold:
                text_content = f"**{text_content}**"
            # Handle italic only
            elif hasattr(style, "italic") and style.italic:
                text_content = f"*{text_content}*"

        # Reconstruct with preserved spacing
        return leading_space + text_content + trailing_space + trailing_newlines

    def _is_numbered_list_glyph(self, glyph: str) -> bool:
        """Determine if a glyph represents a numbered list item."""
        if not glyph:
            return False

        # Check if the glyph contains digits or letters (indicating numbering)
        return any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph)

    def _get_list_indentation(self, nesting_level: int | None) -> str:
        """Get the appropriate indentation for a list item based on nesting level."""
        if nesting_level is None:
            nesting_level = 0

        # Use 2 spaces per nesting level for markdown compatibility
        return "  " * nesting_level

    def _format_bullet_marker_with_nesting(self, glyph: str) -> str:
        """Format the bullet marker based on the glyph and nesting level.

        According to the user's requirement, nested lists should be consistently
        either all ordered or all unordered throughout the nesting hierarchy.
        """
        if not glyph:
            return "* "

        if any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph):
            # For numbered lists, convert all levels to numbered format
            # Since markdown doesn't support nested numbering well, we'll use "1. " format for all levels
            # and rely on indentation to show the hierarchy
            return normalize_numbered_glyph(glyph)
        else:
            # This is a bullet list - use bullets for all levels
            return "* "

    def _format_bullet_marker(self, glyph: str) -> str:
        """Format the bullet marker based on the glyph from the API."""
        if not glyph:
            return "* "

        # Check if this looks like a numbered list
        if any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph):
            # This is a numbered list - use the glyph as-is if it ends with period
            if glyph.endswith("."):
                return f"{glyph} "
            else:
                return f"{glyph}. "
        else:
            # This is a bullet list
            return "* "

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


def text_elements_to_requests(text_elements: List[TextElement | GslidesAPIRequest], objectId: str):
    requests = []
    for te in text_elements:
        if isinstance(te, GslidesAPIRequest):
            te.objectId = objectId
            requests += te.to_request()
            continue
        else:
            assert isinstance(te, TextElement), f"Expected TextElement, got {te}"
        if te.textRun is None:
            # An empty text run will have a non-None ParagraphMarker
            # Apparently there's no direct way to insert ParagraphMarkers, instead they have to be created
            # as a side effect of inserting text or by specialized calls like createParagraphBullets
            # So we just ignore them when inserting text
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


def normalize_numbered_glyph(glyph: str) -> str:
    """Normalize the glyph for numbered lists."""
    if glyph.endswith("."):
        # Try to extract the number
        number_part = glyph[:-1]
    else:
        number_part = glyph
    latin = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"]
    alpha = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]

    if number_part in latin:
        out = str(latin.index(number_part) + 1)
    elif number_part in alpha:
        out = str(alpha.index(number_part) + 1)
    elif number_part.isdigit():
        out = number_part
    else:
        raise ValueError(f"Unsupported glyph format: {glyph}")

    return f"{out}. "
