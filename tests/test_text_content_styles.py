"""Tests for TextContent.styles() whitespace fallback behavior.

When skip_whitespace=True (the default), styles() should skip whitespace-only
text runs. If this yields no styles at all (e.g., the element contains only
whitespace runs), it should fall back to including whitespace runs.
"""

import pytest

from gslides_api.domain.domain import (
    Color,
    Dimension,
    OptionalColor,
    RgbColor,
    ThemeColorType,
    Unit,
)
from gslides_api.domain.text import (
    ParagraphMarker,
    TextElement,
    TextRun,
    TextStyle,
)
from gslides_api.element.text_content import TextContent


def _make_text_style(
    red: float = 0.0,
    green: float = 0.0,
    blue: float = 0.0,
    font_size_pt: float = 12.0,
    font_family: str = "Arial",
    theme_color: ThemeColorType | None = None,
) -> TextStyle:
    """Helper to create a TextStyle with a foreground color."""
    if theme_color is not None:
        foreground = OptionalColor(
            opaqueColor=Color(themeColor=theme_color)
        )
    else:
        foreground = OptionalColor(
            opaqueColor=Color(rgbColor=RgbColor(red=red, green=green, blue=blue))
        )
    return TextStyle(
        fontSize=Dimension(magnitude=font_size_pt, unit=Unit.PT),
        foregroundColor=foreground,
        fontFamily=font_family,
    )


def _make_text_element(content: str, style: TextStyle, start: int = 0) -> TextElement:
    """Helper to create a TextElement with a text run."""
    return TextElement(
        startIndex=start,
        endIndex=start + len(content),
        textRun=TextRun(content=content, style=style),
    )


class TestTextContentStylesWhitespaceFallback:
    """Test the skip_whitespace fallback in TextContent.styles()."""

    def test_skips_whitespace_when_non_whitespace_runs_exist(self):
        """When there are non-whitespace runs, whitespace-only runs should be skipped."""
        dark_blue = _make_text_style(red=0.24, green=0.28, blue=0.61)
        white = _make_text_style(theme_color=ThemeColorType.LIGHT1, font_size_pt=15.0)

        tc = TextContent(textElements=[
            TextElement(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
            _make_text_element("{Value}", style=dark_blue, start=0),
            _make_text_element("  ", style=white, start=7),
            _make_text_element("Label text\n", style=dark_blue, start=9),
        ])

        styles = tc.styles(skip_whitespace=True)
        assert styles is not None
        assert len(styles) == 1  # Only the dark blue style
        assert styles[0].foreground_color is not None
        assert styles[0].foreground_color.theme_color is None  # Not a theme color

    def test_includes_whitespace_when_skip_false(self):
        """When skip_whitespace=False, whitespace-only runs should be included."""
        dark_blue = _make_text_style(red=0.24, green=0.28, blue=0.61)
        white = _make_text_style(theme_color=ThemeColorType.LIGHT1, font_size_pt=15.0)

        tc = TextContent(textElements=[
            TextElement(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
            _make_text_element("{Value}", style=dark_blue, start=0),
            _make_text_element("  ", style=white, start=7),
            _make_text_element("Label text\n", style=dark_blue, start=9),
        ])

        styles = tc.styles(skip_whitespace=False)
        assert styles is not None
        assert len(styles) == 2  # Dark blue + white theme color

    def test_fallback_to_whitespace_when_only_whitespace_runs(self):
        """When all runs are whitespace, should fall back to including them."""
        white = _make_text_style(theme_color=ThemeColorType.LIGHT1, font_size_pt=15.0)

        tc = TextContent(textElements=[
            TextElement(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
            _make_text_element("  ", style=white, start=0),
            _make_text_element("\n", style=white, start=2),
        ])

        # With skip_whitespace=True, should fall back to including whitespace
        styles = tc.styles(skip_whitespace=True)
        assert styles is not None
        assert len(styles) == 1
        assert styles[0].foreground_color.theme_color == "LIGHT1"

    def test_returns_none_when_no_text_elements(self):
        """When there are no text elements, should return None."""
        tc = TextContent(textElements=None)
        assert tc.styles() is None

    def test_returns_none_when_text_elements_empty(self):
        """When text elements list is empty, should return None."""
        tc = TextContent(textElements=[])
        assert tc.styles() is None

    def test_multiple_distinct_non_whitespace_styles(self):
        """Multiple distinct non-whitespace styles should all be returned."""
        heading_style = _make_text_style(red=0.24, green=0.28, blue=0.61, font_size_pt=14.0)
        body_style = _make_text_style(red=0.34, green=0.38, blue=0.53, font_size_pt=11.0, font_family="Inter")
        white = _make_text_style(theme_color=ThemeColorType.LIGHT1, font_size_pt=15.0)

        tc = TextContent(textElements=[
            TextElement(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
            _make_text_element("{Value}", style=heading_style, start=0),
            _make_text_element("  ", style=white, start=7),
            _make_text_element("Label\n", style=heading_style, start=9),
            TextElement(startIndex=15, endIndex=15, paragraphMarker=ParagraphMarker()),
            _make_text_element("Body text\n", style=body_style, start=15),
        ])

        styles = tc.styles(skip_whitespace=True)
        assert styles is not None
        assert len(styles) == 2  # heading + body, not the white spacer

    def test_only_paragraph_markers_fallback_to_whitespace(self):
        """TextContent with only paragraph markers and whitespace should fall back."""
        some_style = _make_text_style(red=0.5, green=0.5, blue=0.5)

        tc = TextContent(textElements=[
            TextElement(startIndex=0, endIndex=0, paragraphMarker=ParagraphMarker()),
            _make_text_element("\n", style=some_style, start=0),
        ])

        # "\n" stripped is "", so with skip_whitespace=True it would be empty,
        # then fallback to skip_whitespace=False
        styles = tc.styles(skip_whitespace=True)
        assert styles is not None
        assert len(styles) == 1
