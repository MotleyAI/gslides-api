from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api import Dimension, Outline, Shadow, ShapeBackgroundFill
from gslides_api.domain import GSlidesBaseModel, OptionalColor


class ShapeType(Enum):
    """Enumeration of possible shape types."""

    TEXT_BOX = "TEXT_BOX"
    RECTANGLE = "RECTANGLE"
    ROUND_RECTANGLE = "ROUND_RECTANGLE"
    ELLIPSE = "ELLIPSE"
    LINE = "LINE"
    IMAGE = "IMAGE"
    UNKNOWN = "UNKNOWN"
    CUSTOM = "CUSTOM"
    CHEVRON = "CHEVRON"
    ROUNDED_RECTANGLE = "ROUND_2_SAME_RECTANGLE"


class PlaceholderType(Enum):
    """Enumeration of possible placeholder types."""

    TITLE = "TITLE"
    BODY = "BODY"
    SUBTITLE = "SUBTITLE"
    CENTERED_TITLE = "CENTERED_TITLE"
    SLIDE_IMAGE = "SLIDE_IMAGE"
    SLIDE_NUMBER = "SLIDE_NUMBER"
    UNKNOWN = "UNKNOWN"


class BaselineOffset(Enum):
    """The ways in which text can be vertically offset from its normal position."""

    BASELINE_OFFSET_UNSPECIFIED = "BASELINE_OFFSET_UNSPECIFIED"
    NONE = "NONE"
    SUPERSCRIPT = "SUPERSCRIPT"
    SUBSCRIPT = "SUBSCRIPT"


class ParagraphStyle(GSlidesBaseModel):
    """Represents styling for paragraphs."""

    direction: str = "LEFT_TO_RIGHT"
    indentStart: Optional[Dict[str, Any]] = None
    indentFirstLine: Optional[Dict[str, Any]] = None
    indentEnd: Optional[Dict[str, Any]] = None
    spacingMode: Optional[str] = None
    lineSpacing: Optional[float] = None
    spaceAbove: Optional[Dict[str, Any]] = None
    spaceBelow: Optional[Dict[str, Any]] = None
    alignment: Optional[str] = None


class BulletStyle(GSlidesBaseModel):
    """Represents styling for bullets in lists."""

    glyph: Optional[str] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    fontFamily: Optional[str] = None


class TextRun(GSlidesBaseModel):
    """Represents a run of text with consistent styling."""

    content: str
    style: "TextStyle" = Field(default_factory=lambda: TextStyle())


class AutoTextType(Enum):
    """Enumeration of possible auto text types."""

    SLIDE_NUMBER = "SLIDE_NUMBER"
    SLIDE_COUNT = "SLIDE_COUNT"
    CURRENT_DATE = "CURRENT_DATE"
    CURRENT_TIME = "CURRENT_TIME"


class WeightedFontFamily(GSlidesBaseModel):
    """Represents a font family and weight used to style a TextRun."""

    fontFamily: str
    weight: int = 400  # Default to "normal" weight


class Link(GSlidesBaseModel):
    """Represents a hyperlink."""

    url: Optional[str] = None
    slideIndex: Optional[int] = None
    pageObjectId: Optional[str] = None
    relativeLink: Optional[str] = None


class TextStyle(GSlidesBaseModel):
    """Represents the styling that can be applied to a TextRun.

    Based on: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/text#Page.TextStyle
    """

    backgroundColor: Optional[OptionalColor] = None
    foregroundColor: Optional[OptionalColor] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    fontFamily: Optional[str] = None
    fontSize: Optional[Dimension] = None
    link: Optional[Link] = None
    baselineOffset: Optional[BaselineOffset] = None
    smallCaps: Optional[bool] = None
    strikethrough: Optional[bool] = None
    underline: Optional[bool] = None
    weightedFontFamily: Optional[WeightedFontFamily] = None


class Bullet(GSlidesBaseModel):
    """Represents a bullet point in a list.

    Based on: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/text#Page.Bullet
    """

    listId: Optional[str] = None
    nestingLevel: Optional[int] = None
    glyph: Optional[str] = None
    bulletStyle: Optional[TextStyle] = None


class ParagraphMarker(GSlidesBaseModel):
    """Represents a paragraph marker with styling."""

    style: ParagraphStyle = Field(default_factory=ParagraphStyle)
    bullet: Optional[Bullet] = None


class ShapeProperties(GSlidesBaseModel):
    """Represents properties of a shape."""

    shapeBackgroundFill: Optional[ShapeBackgroundFill] = None
    outline: Optional[Outline] = None
    shadow: Optional[Shadow] = None
    autofit: Optional[Dict[str, Any]] = None
    contentAlignment: Optional[str] = None


class Placeholder(GSlidesBaseModel):
    """Represents a placeholder in a slide."""

    type: PlaceholderType
    parentObjectId: Optional[str] = None
    index: Optional[int] = None


class AutoText(GSlidesBaseModel):
    """Represents auto text content that is generated automatically."""

    type: AutoTextType
    style: Optional[TextStyle] = Field(default_factory=TextStyle)
    content: Optional[str] = None


class TextElement(GSlidesBaseModel):
    """Represents an element within text content."""

    endIndex: int
    startIndex: Optional[int] = None
    paragraphMarker: Optional[ParagraphMarker] = None
    textRun: Optional[TextRun] = None
    autoText: Optional[AutoText] = None


class Text(GSlidesBaseModel):
    """Represents text content with its elements and lists."""

    textElements: List[TextElement]
    lists: Optional[Dict[str, Any]] = None


class Shape(GSlidesBaseModel):
    """Represents a shape in a slide."""

    shapeProperties: ShapeProperties
    shapeType: Optional[ShapeType] = None  # Make optional to preserve original JSON exactly
    text: Optional[Text] = None
    placeholder: Optional[Placeholder] = None
