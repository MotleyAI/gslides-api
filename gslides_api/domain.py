from typing import List, Dict, Optional, Any, Union
from enum import Enum
from pydantic import BaseModel, model_validator
from pydantic.json import pydantic_encoder


class GSlidesBaseModel(BaseModel):
    """Base class for all models in the Google Slides API."""

    def to_api_format(self) -> Dict[str, Any]:
        """Convert to the format expected by the Google Slides API."""
        return super().model_dump(exclude_none=True, mode="json")


class Unit(Enum):
    """Enumeration of possible units of measurement.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/Unit
    """

    UNIT_UNSPECIFIED = "UNIT_UNSPECIFIED"
    """The units are unknown."""

    EMU = "EMU"
    """An English Metric Unit (EMU) is defined as 1/360,000 of a centimeter and thus there are 914,400 EMUs per inch, and 12,700 EMUs per point."""

    PT = "PT"
    """A point, 1/72 of an inch."""


class Dimension(GSlidesBaseModel):
    """Represents a size dimension with magnitude and unit."""

    magnitude: float
    unit: Unit


class Size(GSlidesBaseModel):
    """Represents a size with width and height."""

    width: Union[float, Dimension]
    height: Union[float, Dimension]


class Transform(GSlidesBaseModel):
    """Represents a transformation applied to an element."""

    translateX: float = 0.0
    translateY: float = 0.0
    scaleX: float = 1.0
    scaleY: float = 1.0
    unit: Optional[str] = None  # Make optional to preserve original JSON exactly


class AffineTransform(GSlidesBaseModel):
    """AffineTransform uses a 3x3 matrix with an implied last row of [ 0 0 1 ] to transform source coordinates (x,y) into destination coordinates (x', y').

    The transformation follows:
    [ x']   [  scaleX  shearX  translateX  ] [ x ]
    [ y'] = [  shearY  scaleY  translateY  ] [ y ]
    [ 1 ]   [      0       0         1     ] [ 1 ]

    After transformation:
    x' = scaleX * x + shearX * y + translateX;
    y' = scaleY * y + shearY * x + translateY;

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/other#Page.AffineTransform
    """

    scaleX: float
    """The X coordinate scaling element."""

    scaleY: float
    """The Y coordinate scaling element."""

    shearX: float
    """The X coordinate shearing element."""

    shearY: float
    """The Y coordinate shearing element."""

    translateX: float
    """The X coordinate translation element."""

    translateY: float
    """The Y coordinate translation element."""

    unit: Optional[Unit] = None
    """The units for translate elements."""


class RgbColor(GSlidesBaseModel):
    """Represents an RGB color."""

    red: Optional[float] = None
    green: Optional[float] = None
    blue: Optional[float] = None


class ThemeColorType(Enum):
    """Enumeration of possible theme color types."""

    THEME_COLOR_TYPE_UNSPECIFIED = "THEME_COLOR_TYPE_UNSPECIFIED"
    DARK1 = "DARK1"
    LIGHT1 = "LIGHT1"
    DARK2 = "DARK2"
    LIGHT2 = "LIGHT2"
    ACCENT1 = "ACCENT1"
    ACCENT2 = "ACCENT2"
    ACCENT3 = "ACCENT3"
    ACCENT4 = "ACCENT4"
    ACCENT5 = "ACCENT5"
    ACCENT6 = "ACCENT6"
    HYPERLINK = "HYPERLINK"
    FOLLOWED_HYPERLINK = "FOLLOWED_HYPERLINK"
    TEXT1 = "TEXT1"
    BACKGROUND1 = "BACKGROUND1"
    TEXT2 = "TEXT2"
    BACKGROUND2 = "BACKGROUND2"


class ThemeColorPair(GSlidesBaseModel):
    """Represents a mapping of a theme color type to its concrete color."""

    type: ThemeColorType
    color: RgbColor


class ColorScheme(GSlidesBaseModel):
    """Represents a predefined color palette for a page."""

    colors: List[ThemeColorPair] = []


class Color(GSlidesBaseModel):
    """Represents a color with RGB values."""

    rgbColor: Optional[RgbColor] = None
    themeColor: Optional[ThemeColorType] = None

    @classmethod
    def from_api_format(cls, data: Dict[str, Any]) -> "Color":
        """Create a Color from API format."""
        if "rgbColor" in data:
            rgb_color = (
                RgbColor(**data["rgbColor"])
                if isinstance(data["rgbColor"], dict)
                else data["rgbColor"]
            )
            theme_color = None
            if "themeColor" in data:
                try:
                    theme_color = ThemeColorType(data["themeColor"])
                except (ValueError, TypeError):
                    # Keep as is if conversion fails
                    theme_color = data["themeColor"]
            return cls(rgbColor=rgb_color, themeColor=theme_color)
        elif "themeColor" in data:
            try:
                theme_color = ThemeColorType(data["themeColor"])
            except (ValueError, TypeError):
                # Keep as is if conversion fails
                theme_color = data["themeColor"]
            return cls(themeColor=theme_color)
        return cls()


class OptionalColor(GSlidesBaseModel):
    """A color that can either be fully opaque or fully transparent."""

    opaqueColor: Optional[Color] = None


class SolidFill(GSlidesBaseModel):
    """Represents a solid fill with color and alpha."""

    color: Optional[Color] = None
    alpha: Optional[float] = None

    @classmethod
    def from_api_format(cls, data: Dict[str, Any]) -> "SolidFill":
        """Create a SolidFill from API format."""
        color = None
        if "color" in data and isinstance(data["color"], dict):
            if "rgbColor" in data["color"] or "themeColor" in data["color"]:
                color = Color.from_api_format(data["color"])
            else:
                color = Color(**data["color"])

        return cls(color=color, alpha=data.get("alpha"))


class ShapeBackgroundFill(GSlidesBaseModel):
    """Represents the background fill of a shape."""

    solidFill: Optional[SolidFill] = None
    propertyState: Optional[str] = None


class OutlineFill(GSlidesBaseModel):
    """Represents the fill of an outline."""

    solidFill: Optional[SolidFill] = None

    @classmethod
    def from_api_format(cls, data: Dict[str, Any]) -> "OutlineFill":
        """Create an OutlineFill from API format."""
        if "solidFill" in data and isinstance(data["solidFill"], dict):
            solid_fill = SolidFill(**data["solidFill"])
            return cls(solidFill=solid_fill)
        return cls()


class Weight(GSlidesBaseModel):
    """Represents the weight of an outline."""

    magnitude: Optional[float] = None
    unit: Optional[str] = None


class DashStyle(Enum):
    """Enumeration of possible dash styles for outlines."""

    DASH_STYLE_UNSPECIFIED = "DASH_STYLE_UNSPECIFIED"
    SOLID = "SOLID"
    DOT = "DOT"
    DASH = "DASH"
    DASH_DOT = "DASH_DOT"
    LONG_DASH = "LONG_DASH"
    LONG_DASH_DOT = "LONG_DASH_DOT"


class Outline(GSlidesBaseModel):
    """Represents an outline of a shape."""

    outlineFill: Optional[OutlineFill] = None
    weight: Optional[Weight] = None
    propertyState: Optional[str] = None
    dashStyle: Optional[DashStyle] = None


class ShadowTransform(GSlidesBaseModel):
    """Represents a shadow transform."""

    scaleX: Optional[float] = None
    scaleY: Optional[float] = None
    unit: Optional[str] = None


class BlurRadius(GSlidesBaseModel):
    """Represents a blur radius."""

    magnitude: Optional[float] = None
    unit: Optional[str] = None


class ShadowType(Enum):
    """Enumeration of possible shadow types."""

    SHADOW_TYPE_UNSPECIFIED = "SHADOW_TYPE_UNSPECIFIED"
    OUTER = "OUTER"


class RectanglePosition(Enum):
    """Enumeration of possible rectangle positions."""

    RECTANGLE_POSITION_UNSPECIFIED = "RECTANGLE_POSITION_UNSPECIFIED"
    TOP_LEFT = "TOP_LEFT"
    TOP_CENTER = "TOP_CENTER"
    TOP_RIGHT = "TOP_RIGHT"
    LEFT_CENTER = "LEFT_CENTER"
    CENTER = "CENTER"
    RIGHT_CENTER = "RIGHT_CENTER"
    BOTTOM_LEFT = "BOTTOM_LEFT"
    BOTTOM_CENTER = "BOTTOM_CENTER"
    BOTTOM_RIGHT = "BOTTOM_RIGHT"


class Shadow(GSlidesBaseModel):
    """Represents a shadow."""

    transform: Optional[ShadowTransform] = None
    blurRadius: Optional[BlurRadius] = None
    color: Optional[Color] = None
    alpha: Optional[float] = None
    rotateWithShape: Optional[bool] = None
    propertyState: Optional[str] = None
    type: Optional[ShadowType] = None
    alignment: Optional[RectanglePosition] = None


class Table(GSlidesBaseModel):
    """Represents a table in a slide."""

    rows: Optional[int] = None
    columns: Optional[int] = None
    tableRows: Optional[List[Dict[str, Any]]] = None
    tableColumns: Optional[List[Dict[str, Any]]] = None
    horizontalBorderRows: Optional[List[Dict[str, Any]]] = None
    verticalBorderRows: Optional[List[Dict[str, Any]]] = None


class CropProperties(GSlidesBaseModel):
    """Represents crop properties of an image."""

    leftOffset: Optional[float] = None
    rightOffset: Optional[float] = None
    topOffset: Optional[float] = None
    bottomOffset: Optional[float] = None
    angle: Optional[float] = None


class ColorStop(GSlidesBaseModel):
    """Represents a color and position in a gradient."""

    color: Optional[Color] = None
    alpha: Optional[float] = 1.0
    position: Optional[float] = None


class RecolorName(Enum):
    """Enumeration of possible recolor effect names."""

    NONE = "NONE"
    LIGHT1 = "LIGHT1"
    LIGHT2 = "LIGHT2"
    LIGHT3 = "LIGHT3"
    LIGHT4 = "LIGHT4"
    LIGHT5 = "LIGHT5"
    LIGHT6 = "LIGHT6"
    LIGHT7 = "LIGHT7"
    LIGHT8 = "LIGHT8"
    LIGHT9 = "LIGHT9"
    LIGHT10 = "LIGHT10"
    DARK1 = "DARK1"
    DARK2 = "DARK2"
    DARK3 = "DARK3"
    DARK4 = "DARK4"
    DARK5 = "DARK5"
    DARK6 = "DARK6"
    DARK7 = "DARK7"
    DARK8 = "DARK8"
    DARK9 = "DARK9"
    DARK10 = "DARK10"
    GRAYSCALE = "GRAYSCALE"
    NEGATIVE = "NEGATIVE"
    SEPIA = "SEPIA"
    CUSTOM = "CUSTOM"


class Recolor(GSlidesBaseModel):
    """Represents a recolor effect applied to an image."""

    recolorStops: Optional[List[ColorStop]] = None
    name: Optional[RecolorName] = None


class ImageProperties(GSlidesBaseModel):
    """Represents properties of an image."""

    cropProperties: Optional[CropProperties] = None
    transparency: Optional[float] = None
    brightness: Optional[float] = None
    contrast: Optional[float] = None
    recolor: Optional[Recolor] = None
    outline: Optional[Outline] = None
    shadow: Optional[Shadow] = None
    link: Optional[Dict[str, Any]] = None


class Image(GSlidesBaseModel):
    """Represents an image in a slide."""

    contentUrl: Optional[str] = None
    imageProperties: Optional[Union[Dict[str, Any], ImageProperties]] = None
    sourceUrl: Optional[str] = None

    @model_validator(mode="after")
    def convert_image_properties(self) -> "Image":
        """Convert imageProperties to ImageProperties if it's a dict."""
        if self.imageProperties is not None and not isinstance(
            self.imageProperties, ImageProperties
        ):
            # Track the original type
            if isinstance(self.imageProperties, dict):
                self._original_properties_type = "dict"
            else:
                self._original_properties_type = type(self.imageProperties).__name__

            try:
                self.imageProperties = ImageProperties.model_validate(self.imageProperties)
            except (ValueError, TypeError):
                # Keep as is if conversion fails
                pass
        return self


class VideoSourceType(Enum):
    """Enumeration of possible video source types."""

    YOUTUBE = "YOUTUBE"
    DRIVE = "DRIVE"
    UNKNOWN = "UNKNOWN"


class ImageReplaceMethod(Enum):
    """Enumeration of possible image replace methods.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#ImageReplaceMethod
    """

    IMAGE_REPLACE_METHOD_UNSPECIFIED = "IMAGE_REPLACE_METHOD_UNSPECIFIED"
    CENTER_INSIDE = "CENTER_INSIDE"
    CENTER_CROP = "CENTER_CROP"


class VideoProperties(GSlidesBaseModel):
    """Represents properties of a video.

    As defined in the Google Slides API:
    https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/videos#videoproperties
    """

    outline: Optional[Outline] = None
    autoPlay: Optional[bool] = None
    start: Optional[int] = None
    end: Optional[int] = None
    mute: Optional[bool] = None


class Video(GSlidesBaseModel):
    """Represents a video in a slide."""

    url: Optional[str] = None
    videoProperties: Optional[VideoProperties] = None
    source: Optional[VideoSourceType] = None
    id: Optional[str] = None


class LineProperties(GSlidesBaseModel):
    """Represents properties of a line."""

    outline: Optional[Outline] = None
    shadow: Optional[Shadow] = None
    link: Optional[Dict[str, Any]] = None


class Line(GSlidesBaseModel):
    """Represents a line in a slide."""

    lineProperties: Optional[LineProperties] = None
    lineType: Optional[str] = None


class WordArt(GSlidesBaseModel):
    """Represents word art in a slide."""

    renderedText: Optional[str] = None


class SheetsChartProperties(GSlidesBaseModel):
    """Represents properties of a sheets chart."""

    outline: Optional[Outline] = None
    shadow: Optional[Shadow] = None


class SheetsChart(GSlidesBaseModel):
    """Represents a sheets chart in a slide."""

    spreadsheetId: Optional[str] = None
    chartId: Optional[int] = None
    contentUrl: Optional[str] = None
    sheetsChartProperties: Optional[SheetsChartProperties] = None


class SpeakerSpotlightProperties(GSlidesBaseModel):
    """Represents properties of a speaker spotlight."""

    outline: Optional[Outline] = None
    shadow: Optional[Shadow] = None


class SpeakerSpotlight(GSlidesBaseModel):
    """Represents a speaker spotlight in a slide."""

    speakerSpotlightProperties: Optional[SpeakerSpotlightProperties] = None


class Group(GSlidesBaseModel):
    """Represents a group of page elements."""

    children: List["PageElement"]  # This will be a list of PageElement objects


class PropertyState(Enum):
    """The possible states of a property."""

    RENDERED = "RENDERED"
    NOT_RENDERED = "NOT_RENDERED"
    INHERIT = "INHERIT"


class StretchedPictureFill(GSlidesBaseModel):
    """Represents a stretched picture fill for a page background."""

    contentUrl: str
    size: Optional[Size] = None


class PageBackgroundFill(GSlidesBaseModel):
    """Represents the background fill of a page."""

    propertyState: Optional[PropertyState] = None
    solidFill: Optional[SolidFill] = None
    stretchedPictureFill: Optional[StretchedPictureFill] = None


class PredefinedLayout(Enum):
    """Enumeration of predefined slide layouts.

    These are common layouts in presentations. However, there is no guarantee that these layouts
    are present in the current master, as they may have been deleted or are not part of the
    design being used. Additionally, the placeholder images in each layout may have changed.
    """

    PREDEFINED_LAYOUT_UNSPECIFIED = "PREDEFINED_LAYOUT_UNSPECIFIED"
    BLANK = "BLANK"
    CAPTION_ONLY = "CAPTION_ONLY"
    TITLE = "TITLE"
    TITLE_AND_BODY = "TITLE_AND_BODY"
    TITLE_AND_TWO_COLUMNS = "TITLE_AND_TWO_COLUMNS"
    TITLE_ONLY = "TITLE_ONLY"
    SECTION_HEADER = "SECTION_HEADER"
    SECTION_TITLE_AND_DESCRIPTION = "SECTION_TITLE_AND_DESCRIPTION"
    ONE_COLUMN_TEXT = "ONE_COLUMN_TEXT"
    MAIN_POINT = "MAIN_POINT"
    BIG_NUMBER = "BIG_NUMBER"


class BulletGlyphPreset(Enum):
    """Enumeration of preset patterns of bullet glyphs for lists.

    These patterns use different kinds of bullets for different nesting levels:
    - ARROW: An arrow, corresponding to Unicode U+2794
    - ARROW3D: An arrow with 3D shading, corresponding to Unicode U+27a2
    - CHECKBOX: A hollow square, corresponding to Unicode U+274f
    - CIRCLE: A hollow circle, corresponding to Unicode U+25cb
    - DIAMOND: A solid diamond, corresponding to Unicode U+25c6
    - DIAMONDX: A diamond with an 'x', corresponding to Unicode U+2756
    - HOLLOWDIAMOND: A hollow diamond, corresponding to Unicode U+25c7
    - DISC: A solid circle, corresponding to Unicode U+25cf
    - SQUARE: A solid square, corresponding to Unicode U+25a0
    - STAR: A star, corresponding to Unicode U+2605
    - LEFTTRIANGLE: A triangle pointing left, corresponding to Unicode U+25c4
    - ALPHA: A lowercase letter, like 'a', 'b', or 'c'
    - UPPERALPHA: An uppercase letter, like 'A', 'B', or 'C'
    - DECIMAL: A number, like '1', '2', or '3'
    - ZERODECIMAL: A number where single digits are prefixed with zero, like '01', '02', '03'
    - ROMAN: A lowercase roman numeral, like 'i', 'ii', or 'iii'
    - UPPERROMAN: An uppercase roman numeral, like 'I', 'II', or 'III'
    """

    # Bulleted list presets
    BULLET_DISC_CIRCLE_SQUARE = "BULLET_DISC_CIRCLE_SQUARE"
    """A bulleted list with a DISC, CIRCLE and SQUARE bullet glyph for the first 3 list nesting levels."""

    BULLET_DIAMONDX_ARROW3D_SQUARE = "BULLET_DIAMONDX_ARROW3D_SQUARE"
    """A bulleted list with a DIAMONDX, ARROW3D and SQUARE bullet glyph for the first 3 list nesting levels."""

    BULLET_CHECKBOX = "BULLET_CHECKBOX"
    """A bulleted list with CHECKBOX bullet glyphs for all list nesting levels."""

    BULLET_ARROW_DIAMOND_DISC = "BULLET_ARROW_DIAMOND_DISC"
    """A bulleted list with a ARROW, DIAMOND and DISC bullet glyph for the first 3 list nesting levels."""

    BULLET_STAR_CIRCLE_SQUARE = "BULLET_STAR_CIRCLE_SQUARE"
    """A bulleted list with a STAR, CIRCLE and SQUARE bullet glyph for the first 3 list nesting levels."""

    BULLET_ARROW3D_CIRCLE_SQUARE = "BULLET_ARROW3D_CIRCLE_SQUARE"
    """A bulleted list with a ARROW3D, CIRCLE and SQUARE bullet glyph for the first 3 list nesting levels."""

    BULLET_LEFTTRIANGLE_DIAMOND_DISC = "BULLET_LEFTTRIANGLE_DIAMOND_DISC"
    """A bulleted list with a LEFTTRIANGLE, DIAMOND and DISC bullet glyph for the first 3 list nesting levels."""

    BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE = "BULLET_DIAMONDX_HOLLOWDIAMOND_SQUARE"
    """A bulleted list with a DIAMONDX, HOLLOWDIAMOND and SQUARE bullet glyph for the first 3 list nesting levels."""

    BULLET_DIAMOND_CIRCLE_SQUARE = "BULLET_DIAMOND_CIRCLE_SQUARE"
    """A bulleted list with a DIAMOND, CIRCLE and SQUARE bullet glyph for the first 3 list nesting levels."""

    # Numbered list presets
    NUMBERED_DIGIT_ALPHA_ROMAN = "NUMBERED_DIGIT_ALPHA_ROMAN"
    """A numbered list with DIGIT, ALPHA and ROMAN numeric glyphs for the first 3 list nesting levels, followed by periods."""

    NUMBERED_DIGIT_ALPHA_ROMAN_PARENS = "NUMBERED_DIGIT_ALPHA_ROMAN_PARENS"
    """A numbered list with DIGIT, ALPHA and ROMAN numeric glyphs for the first 3 list nesting levels, followed by parenthesis."""

    NUMBERED_DIGIT_NESTED = "NUMBERED_DIGIT_NESTED"
    """A numbered list with DIGIT numeric glyphs separated by periods, where each nesting level uses the previous nesting level's glyph as a prefix. For example: '1.', '1.1.', '2.', '2.2.'."""

    NUMBERED_UPPERALPHA_ALPHA_ROMAN = "NUMBERED_UPPERALPHA_ALPHA_ROMAN"
    """A numbered list with UPPERALPHA, ALPHA and ROMAN numeric glyphs for the first 3 list nesting levels, followed by periods."""

    NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT = "NUMBERED_UPPERROMAN_UPPERALPHA_DIGIT"
    """A numbered list with UPPERROMAN, UPPERALPHA and DIGIT numeric glyphs for the first 3 list nesting levels, followed by periods."""

    NUMBERED_ZERODIGIT_ALPHA_ROMAN = "NUMBERED_ZERODIGIT_ALPHA_ROMAN"
    """A numbered list with ZERODIGIT, ALPHA and ROMAN numeric glyphs for the first 3 list nesting levels, followed by periods."""


class LayoutReference(GSlidesBaseModel):
    """Represents a reference to a layout."""

    layoutId: Optional[str] = None
    predefinedLayout: Optional[PredefinedLayout] = None

    @model_validator(mode="after")
    def validate_exactly_one_field_set(self) -> "LayoutReference":
        """Validate that exactly one of layoutId or predefinedLayout is set."""
        if (self.layoutId is None and self.predefinedLayout is None) or (
            self.layoutId is not None and self.predefinedLayout is not None
        ):
            raise ValueError("Exactly one of layoutId or predefinedLayout must be set")
        return self


class MimeType(Enum):
    """The MIME type of the thumbnail image.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail#MimeType
    """

    PNG = "PNG"
    """The default MIME type and apparently the only one supported by the API."""


class ThumbnailSize(Enum):
    """The predefined thumbnail image sizes.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail#ThumbnailSize
    """

    THUMBNAIL_SIZE_UNSPECIFIED = "THUMBNAIL_SIZE_UNSPECIFIED"
    """The default thumbnail image size.

    The unspecified thumbnail size implies that the server chooses the size of the image in a way that might vary in the future.
    """

    LARGE = "LARGE"
    """The thumbnail image width is 1,600 px."""

    MEDIUM = "MEDIUM"
    """The thumbnail image width is 800 px."""

    SMALL = "SMALL"
    """The thumbnail image width is 200 px."""


class ThumbnailProperties(GSlidesBaseModel):
    """Provides control over the creation of page thumbnails.

    Reference: https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail#ThumbnailProperties
    """

    mimeType: Optional[MimeType] = None
    """The optional MIME type of the thumbnail image.

    If you don't specify the MIME type, the default is PNG.
    """

    thumbnailSize: Optional[ThumbnailSize] = None
    """The optional size of the thumbnail image.

    If you don't specify the size, the server chooses a default size for the image.
    """


# https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#pageelementproperties
# Note: When you initially create a PageElement, the API may modify the values of both size and transform, but the visual size will be unchanged.
class PageElementProperties(GSlidesBaseModel):
    """Represents properties of a page element."""

    pageObjectId: Optional[str] = None
    size: Optional[Size] = None
    transform: Optional[Transform] = None
