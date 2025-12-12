"""Converters between Google Slides TextStyle and platform-agnostic styles.

This module provides bidirectional conversion between Google Slides API
text styles and the platform-agnostic MarkdownRenderableStyle/RichStyle classes.
"""

from typing import Optional

from gslides_api.agnostic.text import (
    AbstractColor,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    RichStyle,
)
from gslides_api.domain.domain import (
    Color,
    Dimension,
    OptionalColor,
    RgbColor,
    Unit,
)
from gslides_api.domain.text import (
    BaselineOffset as GSlidesBaselineOffset,
    Link,
    TextStyle,
    WeightedFontFamily,
)


# Monospace font families for code detection
MONOSPACE_FONTS = {
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
}


def _is_monospace(font_family: Optional[str]) -> bool:
    """Check if a font family is a monospace font."""
    if not font_family:
        return False
    return font_family.lower() in MONOSPACE_FONTS


def _dimension_to_pt(dimension: Optional[Dimension]) -> Optional[float]:
    """Convert a GSlides Dimension to points."""
    if dimension is None:
        return None

    if dimension.unit == Unit.PT:
        return dimension.magnitude
    elif dimension.unit == Unit.EMU:
        # 1 point = 12,700 EMUs
        return dimension.magnitude / 12700.0
    else:
        # UNIT_UNSPECIFIED or unknown - assume points
        return dimension.magnitude


def _pt_to_dimension(pt: Optional[float]) -> Optional[Dimension]:
    """Convert points to a GSlides Dimension."""
    if pt is None:
        return None
    return Dimension(magnitude=pt, unit=Unit.PT)


def _optional_color_to_abstract(opt_color: Optional[OptionalColor]) -> Optional[AbstractColor]:
    """Convert GSlides OptionalColor to AbstractColor."""
    if opt_color is None or opt_color.opaqueColor is None:
        return None

    color = opt_color.opaqueColor
    if color.rgbColor is None:
        return None

    rgb = color.rgbColor
    return AbstractColor(
        red=rgb.red if rgb.red is not None else 0.0,
        green=rgb.green if rgb.green is not None else 0.0,
        blue=rgb.blue if rgb.blue is not None else 0.0,
    )


def _abstract_to_optional_color(abstract_color: Optional[AbstractColor]) -> Optional[OptionalColor]:
    """Convert AbstractColor to GSlides OptionalColor."""
    if abstract_color is None:
        return None

    return OptionalColor(
        opaqueColor=Color(
            rgbColor=RgbColor(
                red=abstract_color.red,
                green=abstract_color.green,
                blue=abstract_color.blue,
            )
        )
    )


def _convert_baseline_to_abstract(baseline: Optional[GSlidesBaselineOffset]) -> BaselineOffset:
    """Convert GSlides BaselineOffset to abstract BaselineOffset."""
    if baseline is None:
        return BaselineOffset.NONE

    if baseline == GSlidesBaselineOffset.SUPERSCRIPT:
        return BaselineOffset.SUPERSCRIPT
    elif baseline == GSlidesBaselineOffset.SUBSCRIPT:
        return BaselineOffset.SUBSCRIPT
    else:
        return BaselineOffset.NONE


def _convert_baseline_to_gslides(baseline: BaselineOffset) -> Optional[GSlidesBaselineOffset]:
    """Convert abstract BaselineOffset to GSlides BaselineOffset."""
    if baseline == BaselineOffset.SUPERSCRIPT:
        return GSlidesBaselineOffset.SUPERSCRIPT
    elif baseline == BaselineOffset.SUBSCRIPT:
        return GSlidesBaselineOffset.SUBSCRIPT
    else:
        return None  # NONE means no explicit offset


def gslides_style_to_full(style: Optional[TextStyle]) -> FullTextStyle:
    """Convert GSlides TextStyle to FullTextStyle (both markdown and rich parts).

    Args:
        style: GSlides TextStyle object, or None

    Returns:
        FullTextStyle with both markdown-renderable and rich properties
    """
    if style is None:
        return FullTextStyle()

    # Extract markdown-renderable properties
    markdown = MarkdownRenderableStyle(
        bold=style.bold or False,
        italic=style.italic or False,
        strikethrough=style.strikethrough or False,
        is_code=_is_monospace(style.fontFamily),
        hyperlink=style.link.url if style.link else None,
    )

    # Extract rich properties (non-markdown-renderable)
    rich = RichStyle(
        font_family=style.fontFamily,
        font_size_pt=_dimension_to_pt(style.fontSize),
        font_weight=style.weightedFontFamily.weight if style.weightedFontFamily else None,
        foreground_color=_optional_color_to_abstract(style.foregroundColor),
        background_color=_optional_color_to_abstract(style.backgroundColor),
        underline=style.underline or False,
        small_caps=style.smallCaps or False,
        baseline_offset=_convert_baseline_to_abstract(style.baselineOffset),
    )

    return FullTextStyle(markdown=markdown, rich=rich)


def gslides_style_to_rich(style: Optional[TextStyle]) -> RichStyle:
    """Extract only RichStyle from GSlides TextStyle.

    This is used by the styles() method to get only the non-markdown-renderable
    properties for uniqueness checking.

    Args:
        style: GSlides TextStyle object, or None

    Returns:
        RichStyle with only non-markdown-renderable properties
    """
    return gslides_style_to_full(style).rich


def full_style_to_gslides(style: FullTextStyle) -> TextStyle:
    """Convert FullTextStyle back to GSlides TextStyle.

    Args:
        style: FullTextStyle with both markdown and rich parts

    Returns:
        GSlides TextStyle object
    """
    return rich_style_to_gslides(style.rich, style.markdown)


def rich_style_to_gslides(
    rich: RichStyle,
    markdown: Optional[MarkdownRenderableStyle] = None,
) -> TextStyle:
    """Convert RichStyle (+ optional markdown part) to GSlides TextStyle.

    This is the key function for reconstituting styles when making
    Google Slides API calls. The RichStyle contains the base style
    properties (colors, fonts), and the optional MarkdownRenderableStyle
    adds the formatting properties derived from markdown parsing.

    Args:
        rich: RichStyle with non-markdown-renderable properties
        markdown: Optional MarkdownRenderableStyle with markdown-derivable properties

    Returns:
        GSlides TextStyle object ready for API requests
    """
    # Start with markdown-renderable properties
    bold = markdown.bold if markdown else None
    italic = markdown.italic if markdown else None
    strikethrough = markdown.strikethrough if markdown else None

    # Hyperlink
    link = None
    if markdown and markdown.hyperlink:
        link = Link(url=markdown.hyperlink)

    # Font family - use rich style, or Courier New if markdown says it's code
    font_family = rich.font_family
    if markdown and markdown.is_code and not font_family:
        font_family = "Courier New"

    # Build weighted font family if we have weight
    weighted_font_family = None
    if rich.font_weight is not None:
        weighted_font_family = WeightedFontFamily(
            fontFamily=font_family or "Arial",  # Default font if not specified
            weight=rich.font_weight,
        )

    return TextStyle(
        bold=bold if bold else None,  # Don't set False explicitly
        italic=italic if italic else None,
        strikethrough=strikethrough if strikethrough else None,
        underline=rich.underline if rich.underline else None,
        smallCaps=rich.small_caps if rich.small_caps else None,
        fontFamily=font_family,
        fontSize=_pt_to_dimension(rich.font_size_pt),
        weightedFontFamily=weighted_font_family,
        foregroundColor=_abstract_to_optional_color(rich.foreground_color),
        backgroundColor=_abstract_to_optional_color(rich.background_color),
        baselineOffset=_convert_baseline_to_gslides(rich.baseline_offset),
        link=link,
    )


def markdown_style_to_gslides(markdown: MarkdownRenderableStyle) -> TextStyle:
    """Convert only MarkdownRenderableStyle to GSlides TextStyle.

    This creates a minimal TextStyle with only the properties that
    can be derived from markdown (bold, italic, strikethrough, code, hyperlink).

    Args:
        markdown: MarkdownRenderableStyle

    Returns:
        GSlides TextStyle with only markdown-derivable properties
    """
    link = None
    if markdown.hyperlink:
        link = Link(url=markdown.hyperlink)

    font_family = "Courier New" if markdown.is_code else None

    return TextStyle(
        bold=markdown.bold if markdown.bold else None,
        italic=markdown.italic if markdown.italic else None,
        strikethrough=markdown.strikethrough if markdown.strikethrough else None,
        fontFamily=font_family,
        link=link,
    )
