from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import Field

from gslides_api import Dimension, Outline, Shadow, ShapeBackgroundFill
from gslides_api.domain import GSlidesBaseModel, OptionalColor


class Type(Enum):
    """Enumeration of possible shape types.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/shapes#Page.Type
    """

    # Basic shape types
    TYPE_UNSPECIFIED = "TYPE_UNSPECIFIED"
    TEXT_BOX = "TEXT_BOX"
    RECTANGLE = "RECTANGLE"
    ROUND_RECTANGLE = "ROUND_RECTANGLE"
    ELLIPSE = "ELLIPSE"

    # Arc and arrow shapes
    ARC = "ARC"
    BENT_ARROW = "BENT_ARROW"
    BENT_UP_ARROW = "BENT_UP_ARROW"
    CURVED_DOWN_ARROW = "CURVED_DOWN_ARROW"
    CURVED_LEFT_ARROW = "CURVED_LEFT_ARROW"
    CURVED_RIGHT_ARROW = "CURVED_RIGHT_ARROW"
    CURVED_UP_ARROW = "CURVED_UP_ARROW"
    DOWN_ARROW = "DOWN_ARROW"
    DOWN_ARROW_CALLOUT = "DOWN_ARROW_CALLOUT"
    LEFT_ARROW = "LEFT_ARROW"
    LEFT_ARROW_CALLOUT = "LEFT_ARROW_CALLOUT"
    LEFT_RIGHT_ARROW = "LEFT_RIGHT_ARROW"
    LEFT_RIGHT_ARROW_CALLOUT = "LEFT_RIGHT_ARROW_CALLOUT"
    LEFT_RIGHT_UP_ARROW = "LEFT_RIGHT_UP_ARROW"
    LEFT_UP_ARROW = "LEFT_UP_ARROW"
    NOTCHED_RIGHT_ARROW = "NOTCHED_RIGHT_ARROW"
    QUAD_ARROW = "QUAD_ARROW"
    QUAD_ARROW_CALLOUT = "QUAD_ARROW_CALLOUT"
    RIGHT_ARROW = "RIGHT_ARROW"
    RIGHT_ARROW_CALLOUT = "RIGHT_ARROW_CALLOUT"
    STRIPED_RIGHT_ARROW = "STRIPED_RIGHT_ARROW"
    UP_ARROW = "UP_ARROW"
    UP_ARROW_CALLOUT = "UP_ARROW_CALLOUT"
    UP_DOWN_ARROW = "UP_DOWN_ARROW"
    UTURN_ARROW = "UTURN_ARROW"
    ARROW_EAST = "ARROW_EAST"
    ARROW_NORTH_EAST = "ARROW_NORTH_EAST"
    ARROW_NORTH = "ARROW_NORTH"

    # Geometric shapes
    BEVEL = "BEVEL"
    BLOCK_ARC = "BLOCK_ARC"
    CAN = "CAN"
    CHEVRON = "CHEVRON"
    CHORD = "CHORD"
    CLOUD = "CLOUD"
    CORNER = "CORNER"
    CUBE = "CUBE"
    DECAGON = "DECAGON"
    DIAGONAL_STRIPE = "DIAGONAL_STRIPE"
    DIAMOND = "DIAMOND"
    DODECAGON = "DODECAGON"
    DONUT = "DONUT"
    DOUBLE_WAVE = "DOUBLE_WAVE"
    FOLDED_CORNER = "FOLDED_CORNER"
    FRAME = "FRAME"
    HALF_FRAME = "HALF_FRAME"
    HEART = "HEART"
    HEPTAGON = "HEPTAGON"
    HEXAGON = "HEXAGON"
    HOME_PLATE = "HOME_PLATE"
    HORIZONTAL_SCROLL = "HORIZONTAL_SCROLL"
    IRREGULAR_SEAL_1 = "IRREGULAR_SEAL_1"
    IRREGULAR_SEAL_2 = "IRREGULAR_SEAL_2"
    LIGHTNING_BOLT = "LIGHTNING_BOLT"
    MOON = "MOON"
    NO_SMOKING = "NO_SMOKING"
    OCTAGON = "OCTAGON"
    PARALLELOGRAM = "PARALLELOGRAM"
    PENTAGON = "PENTAGON"
    PIE = "PIE"
    PLAQUE = "PLAQUE"
    PLUS = "PLUS"
    RIGHT_TRIANGLE = "RIGHT_TRIANGLE"
    SMILEY_FACE = "SMILEY_FACE"
    SUN = "SUN"
    TRAPEZOID = "TRAPEZOID"
    TRIANGLE = "TRIANGLE"
    VERTICAL_SCROLL = "VERTICAL_SCROLL"
    WAVE = "WAVE"
    TEARDROP = "TEARDROP"

    # Bracket and brace shapes
    BRACE_PAIR = "BRACE_PAIR"
    BRACKET_PAIR = "BRACKET_PAIR"
    LEFT_BRACE = "LEFT_BRACE"
    LEFT_BRACKET = "LEFT_BRACKET"
    RIGHT_BRACE = "RIGHT_BRACE"
    RIGHT_BRACKET = "RIGHT_BRACKET"

    # Rectangle variants
    ROUND_1_RECTANGLE = "ROUND_1_RECTANGLE"
    ROUND_2_DIAGONAL_RECTANGLE = "ROUND_2_DIAGONAL_RECTANGLE"
    ROUND_2_SAME_RECTANGLE = "ROUND_2_SAME_RECTANGLE"
    SNIP_1_RECTANGLE = "SNIP_1_RECTANGLE"
    SNIP_2_DIAGONAL_RECTANGLE = "SNIP_2_DIAGONAL_RECTANGLE"
    SNIP_2_SAME_RECTANGLE = "SNIP_2_SAME_RECTANGLE"
    SNIP_ROUND_RECTANGLE = "SNIP_ROUND_RECTANGLE"

    # Star shapes
    STAR_4 = "STAR_4"
    STAR_5 = "STAR_5"
    STAR_6 = "STAR_6"
    STAR_7 = "STAR_7"
    STAR_8 = "STAR_8"
    STAR_10 = "STAR_10"
    STAR_12 = "STAR_12"
    STAR_16 = "STAR_16"
    STAR_24 = "STAR_24"
    STAR_32 = "STAR_32"
    STARBURST = "STARBURST"

    # Math symbols
    MATH_DIVIDE = "MATH_DIVIDE"
    MATH_EQUAL = "MATH_EQUAL"
    MATH_MINUS = "MATH_MINUS"
    MATH_MULTIPLY = "MATH_MULTIPLY"
    MATH_NOT_EQUAL = "MATH_NOT_EQUAL"
    MATH_PLUS = "MATH_PLUS"

    # Callout shapes
    WEDGE_ELLIPSE_CALLOUT = "WEDGE_ELLIPSE_CALLOUT"
    WEDGE_RECTANGLE_CALLOUT = "WEDGE_RECTANGLE_CALLOUT"
    WEDGE_ROUND_RECTANGLE_CALLOUT = "WEDGE_ROUND_RECTANGLE_CALLOUT"
    CLOUD_CALLOUT = "CLOUD_CALLOUT"

    # Ribbon shapes
    RIBBON = "RIBBON"
    RIBBON_2 = "RIBBON_2"
    ELLIPSE_RIBBON = "ELLIPSE_RIBBON"
    ELLIPSE_RIBBON_2 = "ELLIPSE_RIBBON_2"

    # Flowchart shapes
    FLOW_CHART_ALTERNATE_PROCESS = "FLOW_CHART_ALTERNATE_PROCESS"
    FLOW_CHART_COLLATE = "FLOW_CHART_COLLATE"
    FLOW_CHART_CONNECTOR = "FLOW_CHART_CONNECTOR"
    FLOW_CHART_DECISION = "FLOW_CHART_DECISION"
    FLOW_CHART_DELAY = "FLOW_CHART_DELAY"
    FLOW_CHART_DISPLAY = "FLOW_CHART_DISPLAY"
    FLOW_CHART_DOCUMENT = "FLOW_CHART_DOCUMENT"
    FLOW_CHART_EXTRACT = "FLOW_CHART_EXTRACT"
    FLOW_CHART_INPUT_OUTPUT = "FLOW_CHART_INPUT_OUTPUT"
    FLOW_CHART_INTERNAL_STORAGE = "FLOW_CHART_INTERNAL_STORAGE"
    FLOW_CHART_MAGNETIC_DISK = "FLOW_CHART_MAGNETIC_DISK"
    FLOW_CHART_MAGNETIC_DRUM = "FLOW_CHART_MAGNETIC_DRUM"
    FLOW_CHART_MAGNETIC_TAPE = "FLOW_CHART_MAGNETIC_TAPE"
    FLOW_CHART_MANUAL_INPUT = "FLOW_CHART_MANUAL_INPUT"
    FLOW_CHART_MANUAL_OPERATION = "FLOW_CHART_MANUAL_OPERATION"
    FLOW_CHART_MERGE = "FLOW_CHART_MERGE"
    FLOW_CHART_MULTIDOCUMENT = "FLOW_CHART_MULTIDOCUMENT"
    FLOW_CHART_OFFLINE_STORAGE = "FLOW_CHART_OFFLINE_STORAGE"
    FLOW_CHART_OFFPAGE_CONNECTOR = "FLOW_CHART_OFFPAGE_CONNECTOR"
    FLOW_CHART_ONLINE_STORAGE = "FLOW_CHART_ONLINE_STORAGE"
    FLOW_CHART_OR = "FLOW_CHART_OR"
    FLOW_CHART_PREDEFINED_PROCESS = "FLOW_CHART_PREDEFINED_PROCESS"
    FLOW_CHART_PREPARATION = "FLOW_CHART_PREPARATION"
    FLOW_CHART_PROCESS = "FLOW_CHART_PROCESS"
    FLOW_CHART_PUNCHED_CARD = "FLOW_CHART_PUNCHED_CARD"
    FLOW_CHART_PUNCHED_TAPE = "FLOW_CHART_PUNCHED_TAPE"
    FLOW_CHART_SORT = "FLOW_CHART_SORT"
    FLOW_CHART_SUMMING_JUNCTION = "FLOW_CHART_SUMMING_JUNCTION"
    FLOW_CHART_TERMINATOR = "FLOW_CHART_TERMINATOR"

    # Miscellaneous shapes
    SPEECH = "SPEECH"
    CUSTOM = "CUSTOM"

    # Legacy/compatibility values (keeping existing ones for backward compatibility)
    LINE = "LINE"  # Not in the official API but keeping for compatibility
    IMAGE = "IMAGE"  # Not in the official API but keeping for compatibility
    UNKNOWN = "UNKNOWN"  # Not in the official API but keeping for compatibility
    ROUNDED_RECTANGLE = "ROUND_2_SAME_RECTANGLE"  # Alias for backward compatibility


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
    shapeType: Optional[Type] = None  # Make optional to preserve original JSON exactly
    text: Optional[Text] = None
    placeholder: Optional[Placeholder] = None
