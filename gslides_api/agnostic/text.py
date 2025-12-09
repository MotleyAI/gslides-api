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
    """

    red: float = 0.0
    green: float = 0.0
    blue: float = 0.0
    alpha: float = 1.0

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
