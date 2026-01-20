"""Converters between Google Slides TextStyle and platform-agnostic styles.

This module provides bidirectional conversion between Google Slides API
text styles and the platform-agnostic MarkdownRenderableStyle/RichStyle classes.
"""

from typing import List, Optional

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.text import (
    AbstractColor,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    RichStyle,
)
from gslides_api.agnostic.units import EMU_PER_PT
from gslides_api.domain.domain import (
    Color,
    Dimension,
    OptionalColor,
    RgbColor,
    ThemeColorType,
    Unit,
)
from gslides_api.domain.text import (
    BaselineOffset as GSlidesBaselineOffset,
    Link,
    TextElement,
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
        return dimension.magnitude / EMU_PER_PT
    else:
        # UNIT_UNSPECIFIED or unknown - assume points
        return dimension.magnitude


def _pt_to_dimension(pt: Optional[float]) -> Optional[Dimension]:
    """Convert points to a GSlides Dimension."""
    if pt is None:
        return None
    return Dimension(magnitude=pt, unit=Unit.PT)


def _optional_color_to_abstract(opt_color: Optional[OptionalColor]) -> Optional[AbstractColor]:
    """Convert GSlides OptionalColor to AbstractColor.

    Handles both RGB colors and theme colors. Theme colors (e.g., LIGHT1, DARK1)
    are preserved in the AbstractColor.theme_color field so they can be
    converted back without loss.
    """
    if opt_color is None or opt_color.opaqueColor is None:
        return None

    color = opt_color.opaqueColor

    # Check for theme color first - this takes precedence
    if color.themeColor is not None:
        return AbstractColor(theme_color=color.themeColor.value)

    # Fall back to RGB color
    if color.rgbColor is None:
        return None

    rgb = color.rgbColor
    return AbstractColor(
        red=rgb.red if rgb.red is not None else 0.0,
        green=rgb.green if rgb.green is not None else 0.0,
        blue=rgb.blue if rgb.blue is not None else 0.0,
    )


def _abstract_to_optional_color(abstract_color: Optional[AbstractColor]) -> Optional[OptionalColor]:
    """Convert AbstractColor to GSlides OptionalColor.

    If the AbstractColor has a theme_color set, it is converted back to a
    theme color (preserving the original). Otherwise, RGB values are used.
    """
    if abstract_color is None:
        return None

    # Check for theme color first - this takes precedence
    if abstract_color.theme_color is not None:
        try:
            theme_color_type = ThemeColorType(abstract_color.theme_color)
            return OptionalColor(opaqueColor=Color(themeColor=theme_color_type))
        except ValueError:
            # Invalid theme color string, fall through to RGB
            pass

    # Use RGB color
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


def _is_numbered_list_glyph(glyph: str) -> bool:
    """Determine if a glyph represents a numbered list item."""
    if not glyph:
        return False
    # Check if the glyph contains digits or letters (indicating numbering)
    return any(char.isdigit() for char in glyph) or any(
        char.isalpha() for char in glyph
    )


def text_elements_to_ir(elements: List[TextElement]) -> FormattedDocument:
    """Convert Google Slides TextElements to platform-agnostic IR.

    This function parses Google Slides TextElement objects and produces a
    FormattedDocument with FormattedParagraphs and FormattedLists.

    Args:
        elements: List of TextElement objects from Google Slides API

    Returns:
        FormattedDocument containing the converted content
    """
    if not elements:
        return FormattedDocument()

    result_elements = []
    current_paragraph_runs = []

    # Track list state
    current_list = None  # FormattedList being built
    current_list_item_runs = []  # Runs for the current list item
    pending_bullet_info = None  # (nesting_level, is_ordered) for next text
    in_list_item = False  # Whether we're currently building a list item
    current_item_nesting_level = 0  # Nesting level of the current item being built

    for te in elements:
        # Handle paragraph markers (for bullets and paragraph breaks)
        if te.paragraphMarker is not None:
            if te.paragraphMarker.bullet is not None:
                bullet = te.paragraphMarker.bullet
                nesting_level = bullet.nestingLevel if bullet.nestingLevel is not None else 0
                glyph = bullet.glyph if bullet.glyph else "●"
                is_ordered = _is_numbered_list_glyph(glyph)

                # Store bullet info for the next text run
                pending_bullet_info = (nesting_level, is_ordered)
            else:
                # Regular paragraph marker - no bullet
                # Flush any pending list item
                if in_list_item and current_list_item_runs and current_list is not None:
                    item_para = FormattedParagraph(runs=current_list_item_runs)
                    list_item = FormattedListItem(
                        paragraphs=[item_para],
                        nesting_level=current_item_nesting_level,
                    )
                    current_list.items.append(list_item)
                    current_list_item_runs = []
                    in_list_item = False

                pending_bullet_info = None
            continue

        # Handle text runs
        if te.textRun is not None:
            content = te.textRun.content
            style = gslides_style_to_full(te.textRun.style)

            # Check if we're starting a new bullet item
            if pending_bullet_info is not None:
                nesting_level, is_ordered = pending_bullet_info

                # Check if we need to start a new list or if the list type changed
                if current_list is None or current_list.ordered != is_ordered:
                    # Flush any pending paragraph
                    if current_paragraph_runs:
                        para = FormattedParagraph(runs=current_paragraph_runs)
                        result_elements.append(para)
                        current_paragraph_runs = []

                    # Flush the old list if type changed
                    if current_list is not None and current_list.ordered != is_ordered:
                        result_elements.append(current_list)
                        current_list = None

                    # Start a new list
                    current_list = FormattedList(ordered=is_ordered, items=[])

                # We're now in a list item
                in_list_item = True
                current_item_nesting_level = nesting_level  # Track for continuation runs

                # Strip the trailing newline from the content
                has_newline = "\n" in content
                item_content = content.rstrip("\n")

                if item_content:
                    item_run = FormattedTextRun(content=item_content, style=style)
                    current_list_item_runs.append(item_run)

                # If content ends with newline, complete this list item
                if has_newline and current_list_item_runs:
                    item_para = FormattedParagraph(runs=current_list_item_runs)
                    list_item = FormattedListItem(
                        paragraphs=[item_para],
                        nesting_level=nesting_level,
                    )
                    current_list.items.append(list_item)
                    current_list_item_runs = []
                    in_list_item = False

                pending_bullet_info = None  # Clear after processing first run
            elif in_list_item:
                # We're continuing a list item (subsequent run after bullet)
                has_newline = "\n" in content
                item_content = content.rstrip("\n")

                if item_content:
                    item_run = FormattedTextRun(content=item_content, style=style)
                    current_list_item_runs.append(item_run)

                # If content ends with newline, complete this list item
                if has_newline and current_list_item_runs and current_list is not None:
                    # Use the nesting level from when the item started
                    item_para = FormattedParagraph(runs=current_list_item_runs)
                    list_item = FormattedListItem(
                        paragraphs=[item_para],
                        nesting_level=current_item_nesting_level,
                    )
                    current_list.items.append(list_item)
                    current_list_item_runs = []
                    in_list_item = False
            else:
                # Regular text (not in a list)
                # Flush any pending list first
                if current_list is not None:
                    # Flush any remaining list item runs
                    if current_list_item_runs:
                        item_para = FormattedParagraph(runs=current_list_item_runs)
                        list_item = FormattedListItem(paragraphs=[item_para], nesting_level=current_item_nesting_level)
                        current_list.items.append(list_item)
                        current_list_item_runs = []
                    result_elements.append(current_list)
                    current_list = None
                    in_list_item = False

                # Create the formatted run
                run = FormattedTextRun(content=content, style=style)

                # Add the run to the current paragraph
                current_paragraph_runs.append(run)

                # Handle line breaks - if content ends with newline, complete the paragraph
                if "\n" in content:
                    # Create paragraph from accumulated runs
                    para = FormattedParagraph(runs=current_paragraph_runs)
                    result_elements.append(para)
                    current_paragraph_runs = []

    # Flush any remaining content
    if current_list is not None:
        # Flush any remaining list item runs
        if current_list_item_runs:
            item_para = FormattedParagraph(runs=current_list_item_runs)
            list_item = FormattedListItem(paragraphs=[item_para], nesting_level=current_item_nesting_level)
            current_list.items.append(list_item)
        result_elements.append(current_list)
    if current_paragraph_runs:
        para = FormattedParagraph(runs=current_paragraph_runs)
        result_elements.append(para)

    return FormattedDocument(elements=result_elements)
