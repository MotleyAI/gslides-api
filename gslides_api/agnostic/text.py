"""Platform-agnostic text style classes for Google Slides and PowerPoint.

This module provides a split style architecture:
- MarkdownRenderableStyle: Properties that CAN be encoded in markdown (bold, italic, etc.)
- RichStyle: Properties that CANNOT be encoded in markdown (colors, fonts, etc.)

The `styles()` method returns only RichStyle objects, so text that differs only
in bold/italic/strikethrough is considered ONE style. Markdown formatting is stored
in the markdown string itself, while RichStyle is stored separately and reapplied
when writing.
"""

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class BaselineOffset(Enum):
    """Vertical offset for text (superscript/subscript)."""

    NONE = "none"
    SUPERSCRIPT = "superscript"
    SUBSCRIPT = "subscript"


class AbstractColor(BaseModel):
    """Platform-agnostic color representation using 0.0-1.0 scale.

    This matches Google Slides API color format and can be converted to/from
    various formats (RGB tuples, hex strings, etc.).

    Can represent either:
    - RGB color: red, green, blue values (0.0-1.0)
    - Theme color: theme_color string (e.g., "LIGHT1", "DARK1", "ACCENT1")

    If theme_color is set, it takes precedence over RGB values when converting
    back to Google Slides format.
    """

    red: float = 0.0
    green: float = 0.0
    blue: float = 0.0
    alpha: float = 1.0
    theme_color: Optional[str] = None  # e.g., "LIGHT1", "DARK1", "ACCENT1"

    @classmethod
    def from_rgb_tuple(cls, rgb: tuple[int, int, int]) -> "AbstractColor":
        """Create from RGB tuple with 0-255 values."""
        return cls(red=rgb[0] / 255, green=rgb[1] / 255, blue=rgb[2] / 255)

    @classmethod
    def from_rgb_float(cls, red: float, green: float, blue: float) -> "AbstractColor":
        """Create from RGB floats (0.0-1.0 scale)."""
        return cls(red=red, green=green, blue=blue)

    def to_rgb_tuple(self) -> tuple[int, int, int]:
        """Convert to RGB tuple with 0-255 values."""
        return (int(self.red * 255), int(self.green * 255), int(self.blue * 255))

    def to_hex(self) -> str:
        """Convert to hex color string (#RRGGBB)."""
        r, g, b = self.to_rgb_tuple()
        return f"#{r:02x}{g:02x}{b:02x}"


class MarkdownRenderableStyle(BaseModel):
    """Properties that CAN be encoded in standard markdown.

    These are stored IN the markdown string itself, not in the styles() list.
    When comparing styles for uniqueness, these properties are IGNORED.
    """

    bold: bool = False
    italic: bool = False
    strikethrough: bool = False
    is_code: bool = False  # Monospace/code span (detected via font family)
    hyperlink: Optional[str] = None


class RichStyle(BaseModel):
    """Properties that CANNOT be encoded in standard markdown.

    These are extracted via styles() and reapplied when writing.
    Uniqueness checking is done only on this class - text that differs
    only in bold/italic/strikethrough is considered ONE style.
    """

    # Font properties
    font_family: Optional[str] = None
    font_size_pt: Optional[float] = None  # Always in points
    font_weight: Optional[int] = None  # 100-900, 400=normal, 700=bold

    # Colors
    foreground_color: Optional[AbstractColor] = None
    background_color: Optional[AbstractColor] = None  # highlight/background

    # Non-markdown formatting
    underline: bool = False
    small_caps: bool = False
    all_caps: bool = False
    baseline_offset: BaselineOffset = BaselineOffset.NONE
    character_spacing: Optional[float] = None  # In points

    # Decorative (less commonly used)
    shadow: bool = False
    emboss: bool = False
    imprint: bool = False
    double_strike: bool = False

    def is_monospace(self) -> bool:
        """Check if the font family is a monospace font."""
        if not self.font_family:
            return False
        return self.font_family.lower() in [
            "courier new",
            "courier",
            "monospace",
            "consolas",
            "monaco",
            "lucida console",
            "dejavu sans mono",
            "source code pro",
            "fira code",
            "jetbrains mono",
        ]

    def is_default(self) -> bool:
        """Check if this is a default (empty) style with no properties set.

        A default style has no font, colors, or special formatting.
        """
        return (
            self.font_family is None
            and self.font_size_pt is None
            and self.font_weight is None
            and self.foreground_color is None
            and self.background_color is None
            and self.underline is False
            and self.small_caps is False
            and self.all_caps is False
            and self.baseline_offset == BaselineOffset.NONE
            and self.character_spacing is None
            and self.shadow is False
            and self.emboss is False
            and self.imprint is False
            and self.double_strike is False
        )


class FullTextStyle(BaseModel):
    """Complete style combining both markdown-renderable and rich parts.

    Used during extraction and application, but styles() returns only RichStyle.
    """

    markdown: MarkdownRenderableStyle = Field(default_factory=MarkdownRenderableStyle)
    rich: RichStyle = Field(default_factory=RichStyle)


class AbstractTextRun(BaseModel):
    """Platform-agnostic text run = content + full style.

    A text run is a contiguous piece of text with consistent styling.
    """

    content: str
    style: FullTextStyle = Field(default_factory=FullTextStyle)


class ParagraphAlignment(Enum):
    """Paragraph horizontal alignment."""

    LEFT = "l"
    CENTER = "ctr"
    RIGHT = "r"
    JUSTIFIED = "just"
    JUSTIFIED_LOW = "justLow"  # Low kashida justify
    DISTRIBUTED = "dist"
    THAI_DISTRIBUTED = "thaiDist"


class SpacingValue(BaseModel):
    """Spacing value that can be either points or percentage.

    In PPTX XML:
    - spcPts val="900" means 9pt (value is in 100ths of a point)
    - spcPct val="110000" means 110% (value is in 1/1000ths of a percent)
    """

    points: Optional[float] = None  # In points (e.g., 9.0 for 9pt)
    percentage: Optional[float] = None  # As decimal (e.g., 1.1 for 110%)

    @classmethod
    def from_pptx_pts(cls, val: str) -> "SpacingValue":
        """Create from PPTX spcPts val (100ths of a point)."""
        return cls(points=int(val) / 100)

    @classmethod
    def from_pptx_pct(cls, val: str) -> "SpacingValue":
        """Create from PPTX spcPct val (1/1000ths of percent)."""
        return cls(percentage=int(val) / 100000)

    def to_pptx_pts(self) -> str:
        """Convert to PPTX spcPts val string."""
        if self.points is not None:
            return str(int(self.points * 100))
        return "0"

    def to_pptx_pct(self) -> str:
        """Convert to PPTX spcPct val string."""
        if self.percentage is not None:
            return str(int(self.percentage * 100000))
        return "100000"


class ParagraphStyle(BaseModel):
    """Platform-agnostic paragraph-level formatting properties.

    These are distinct from text/run-level styles (FullTextStyle) and control
    paragraph layout: margins, indents, spacing, alignment, etc.

    All EMU (English Metric Unit) values are stored in their native EMU format.
    1 inch = 914400 EMU, 1 point = 12700 EMU.
    """

    # Margins and indents (EMU values stored as integers)
    margin_left: Optional[int] = None  # marL - left margin
    margin_right: Optional[int] = None  # marR - right margin
    indent: Optional[int] = None  # indent - first line (negative = hanging)

    # Alignment
    alignment: Optional[ParagraphAlignment] = None  # algn
    right_to_left: Optional[bool] = None  # rtl

    # Spacing
    line_spacing: Optional[SpacingValue] = None  # lnSpc
    space_before: Optional[SpacingValue] = None  # spcBef
    space_after: Optional[SpacingValue] = None  # spcAft

    # Other properties
    level: Optional[int] = None  # lvl - outline/list level (0-8)
    default_tab_size: Optional[int] = None  # defTabSz (EMU)

    @classmethod
    def from_pptx_pPr(cls, pPr, ns: str) -> "ParagraphStyle":
        """Create from a PPTX paragraph properties XML element.

        Args:
            pPr: The <a:pPr> XML element
            ns: The DrawingML namespace string

        Returns:
            ParagraphStyle with extracted properties
        """
        style = cls()

        # Extract attributes
        if pPr.get("marL"):
            style.margin_left = int(pPr.get("marL"))
        if pPr.get("marR"):
            style.margin_right = int(pPr.get("marR"))
        if pPr.get("indent"):
            style.indent = int(pPr.get("indent"))
        if pPr.get("algn"):
            try:
                style.alignment = ParagraphAlignment(pPr.get("algn"))
            except ValueError:
                pass
        if pPr.get("rtl"):
            style.right_to_left = pPr.get("rtl") == "1"
        if pPr.get("lvl"):
            style.level = int(pPr.get("lvl"))
        if pPr.get("defTabSz"):
            style.default_tab_size = int(pPr.get("defTabSz"))

        # Extract child elements for spacing
        lnSpc = pPr.find(f"{{{ns}}}lnSpc")
        if lnSpc is not None:
            spcPct = lnSpc.find(f"{{{ns}}}spcPct")
            spcPts = lnSpc.find(f"{{{ns}}}spcPts")
            if spcPct is not None and spcPct.get("val"):
                style.line_spacing = SpacingValue.from_pptx_pct(spcPct.get("val"))
            elif spcPts is not None and spcPts.get("val"):
                style.line_spacing = SpacingValue.from_pptx_pts(spcPts.get("val"))

        spcBef = pPr.find(f"{{{ns}}}spcBef")
        if spcBef is not None:
            spcPct = spcBef.find(f"{{{ns}}}spcPct")
            spcPts = spcBef.find(f"{{{ns}}}spcPts")
            if spcPct is not None and spcPct.get("val"):
                style.space_before = SpacingValue.from_pptx_pct(spcPct.get("val"))
            elif spcPts is not None and spcPts.get("val"):
                style.space_before = SpacingValue.from_pptx_pts(spcPts.get("val"))

        spcAft = pPr.find(f"{{{ns}}}spcAft")
        if spcAft is not None:
            spcPct = spcAft.find(f"{{{ns}}}spcPct")
            spcPts = spcAft.find(f"{{{ns}}}spcPts")
            if spcPct is not None and spcPct.get("val"):
                style.space_after = SpacingValue.from_pptx_pct(spcPct.get("val"))
            elif spcPts is not None and spcPts.get("val"):
                style.space_after = SpacingValue.from_pptx_pts(spcPts.get("val"))

        return style

    def has_bullet_properties(self) -> bool:
        """Check if this style looks like a bullet paragraph (has marL and negative indent)."""
        return (
            self.margin_left is not None
            and self.margin_left > 0
            and self.indent is not None
            and self.indent < 0
        )
