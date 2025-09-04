from typing import Any, Dict, List, Optional


from gslides_api.domain import GSlidesBaseModel
from gslides_api.text import Placeholder, TextStyle, Type, TextElement, ShapeProperties
from gslides_api.markdown.to_markdown import text_elements_to_markdown


class TextContent(GSlidesBaseModel):
    """Represents text content with its elements and lists."""

    textElements: Optional[List[TextElement]] = None
    lists: Optional[Dict[str, Any]] = None

    @property
    def styles(self) -> List[TextStyle] | None:
        """Extract all unique text styles from the text elements."""
        if not self.textElements:
            return None
        styles = []
        for te in self.textElements:
            if te.textRun is None:
                continue
            if te.textRun.content.strip() == "":
                continue
            if te.textRun.style not in styles:
                styles.append(te.textRun.style)
        return styles

    def to_markdown(self) -> str | None:
        """Convert the shape's text content back to markdown format.

        This method reconstructs markdown from the Google Slides API response,
        handling formatting like bold, italic, bullet points, nested lists, and code spans.
        """
        if not self.textElements:
            return None

        return text_elements_to_markdown(self.textElements)

    @property
    def has_text(self):
        return len(self.textElements) > 0 and self.textElements[-1].endIndex > 0


class TextContainer(GSlidesBaseModel):
    """Contains shared text manipulation logic between text boxes and tables."""


class Shape(GSlidesBaseModel):
    """Represents a shape in a slide."""

    shapeProperties: ShapeProperties
    shapeType: Optional[Type] = None  # Make optional to preserve original JSON exactly
    text: Optional[TextContent] = None
    placeholder: Optional[Placeholder] = None
