"""Tests for GSlides TextStyle <-> platform-agnostic style converters."""

import pytest

from gslides_api.agnostic.converters import (
    MONOSPACE_FONTS,
    _abstract_to_optional_color,
    _convert_baseline_to_abstract,
    _convert_baseline_to_gslides,
    _dimension_to_pt,
    _is_monospace,
    _optional_color_to_abstract,
    _pt_to_dimension,
    full_style_to_gslides,
    gslides_style_to_full,
    gslides_style_to_rich,
    markdown_style_to_gslides,
    rich_style_to_gslides,
)
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


class TestMonospaceDetection:
    """Tests for monospace font detection."""

    def test_monospace_fonts_detected(self):
        """Common monospace fonts should be detected."""
        for font in MONOSPACE_FONTS:
            assert _is_monospace(font) is True, f"{font} should be monospace"

    def test_monospace_case_insensitive(self):
        """Font detection should be case-insensitive."""
        assert _is_monospace("Courier New") is True
        assert _is_monospace("COURIER NEW") is True
        assert _is_monospace("courier new") is True

    def test_non_monospace_fonts(self):
        """Non-monospace fonts should not be detected."""
        assert _is_monospace("Arial") is False
        assert _is_monospace("Times New Roman") is False
        assert _is_monospace("Helvetica") is False

    def test_none_font(self):
        """None font should return False."""
        assert _is_monospace(None) is False

    def test_empty_font(self):
        """Empty font should return False."""
        assert _is_monospace("") is False


class TestDimensionConversion:
    """Tests for dimension (font size) conversion."""

    def test_pt_to_pt(self):
        """Points should pass through unchanged."""
        dimension = Dimension(magnitude=14.0, unit=Unit.PT)
        assert _dimension_to_pt(dimension) == 14.0

    def test_emu_to_pt(self):
        """EMUs should convert to points correctly."""
        # 1 point = 12,700 EMUs
        dimension = Dimension(magnitude=12700.0, unit=Unit.EMU)
        assert _dimension_to_pt(dimension) == 1.0

        dimension2 = Dimension(magnitude=127000.0, unit=Unit.EMU)
        assert _dimension_to_pt(dimension2) == 10.0

    def test_none_dimension(self):
        """None dimension should return None."""
        assert _dimension_to_pt(None) is None

    def test_zero_magnitude(self):
        """Dimension with zero magnitude should return zero."""
        dimension = Dimension(magnitude=0.0, unit=Unit.PT)
        assert _dimension_to_pt(dimension) == 0.0

    def test_pt_to_dimension(self):
        """Points should convert to Dimension correctly."""
        result = _pt_to_dimension(14.0)
        assert result is not None
        assert result.magnitude == 14.0
        assert result.unit == Unit.PT

    def test_pt_to_dimension_none(self):
        """None points should return None."""
        assert _pt_to_dimension(None) is None

    def test_dimension_roundtrip(self):
        """Dimension should roundtrip correctly."""
        original_pt = 14.5
        dimension = _pt_to_dimension(original_pt)
        result_pt = _dimension_to_pt(dimension)
        assert result_pt == original_pt


class TestColorConversion:
    """Tests for color conversion."""

    def test_optional_color_to_abstract(self):
        """GSlides OptionalColor should convert to AbstractColor."""
        opt_color = OptionalColor(
            opaqueColor=Color(
                rgbColor=RgbColor(red=1.0, green=0.5, blue=0.0)
            )
        )
        result = _optional_color_to_abstract(opt_color)
        assert result is not None
        assert result.red == 1.0
        assert result.green == 0.5
        assert result.blue == 0.0

    def test_optional_color_none(self):
        """None OptionalColor should return None."""
        assert _optional_color_to_abstract(None) is None

    def test_optional_color_no_opaque(self):
        """OptionalColor with no opaqueColor should return None."""
        opt_color = OptionalColor(opaqueColor=None)
        assert _optional_color_to_abstract(opt_color) is None

    def test_optional_color_no_rgb(self):
        """OptionalColor with no rgbColor should return None."""
        opt_color = OptionalColor(opaqueColor=Color(rgbColor=None))
        assert _optional_color_to_abstract(opt_color) is None

    def test_optional_color_partial_rgb(self):
        """OptionalColor with partial RGB should default missing to 0.0."""
        opt_color = OptionalColor(
            opaqueColor=Color(
                rgbColor=RgbColor(red=1.0, green=None, blue=None)
            )
        )
        result = _optional_color_to_abstract(opt_color)
        assert result is not None
        assert result.red == 1.0
        assert result.green == 0.0
        assert result.blue == 0.0

    def test_abstract_to_optional_color(self):
        """AbstractColor should convert to GSlides OptionalColor."""
        abstract = AbstractColor(red=1.0, green=0.5, blue=0.25)
        result = _abstract_to_optional_color(abstract)
        assert result is not None
        assert result.opaqueColor is not None
        assert result.opaqueColor.rgbColor is not None
        assert result.opaqueColor.rgbColor.red == 1.0
        assert result.opaqueColor.rgbColor.green == 0.5
        assert result.opaqueColor.rgbColor.blue == 0.25

    def test_abstract_to_optional_color_none(self):
        """None AbstractColor should return None."""
        assert _abstract_to_optional_color(None) is None

    def test_color_roundtrip(self):
        """Color should roundtrip correctly."""
        original = AbstractColor(red=0.8, green=0.4, blue=0.2)
        opt_color = _abstract_to_optional_color(original)
        result = _optional_color_to_abstract(opt_color)
        assert result == original


class TestBaselineConversion:
    """Tests for baseline offset conversion."""

    def test_superscript_to_abstract(self):
        """GSlides SUPERSCRIPT should convert to abstract SUPERSCRIPT."""
        result = _convert_baseline_to_abstract(GSlidesBaselineOffset.SUPERSCRIPT)
        assert result == BaselineOffset.SUPERSCRIPT

    def test_subscript_to_abstract(self):
        """GSlides SUBSCRIPT should convert to abstract SUBSCRIPT."""
        result = _convert_baseline_to_abstract(GSlidesBaselineOffset.SUBSCRIPT)
        assert result == BaselineOffset.SUBSCRIPT

    def test_none_to_abstract(self):
        """None baseline should convert to abstract NONE."""
        result = _convert_baseline_to_abstract(None)
        assert result == BaselineOffset.NONE

    def test_superscript_to_gslides(self):
        """Abstract SUPERSCRIPT should convert to GSlides SUPERSCRIPT."""
        result = _convert_baseline_to_gslides(BaselineOffset.SUPERSCRIPT)
        assert result == GSlidesBaselineOffset.SUPERSCRIPT

    def test_subscript_to_gslides(self):
        """Abstract SUBSCRIPT should convert to GSlides SUBSCRIPT."""
        result = _convert_baseline_to_gslides(BaselineOffset.SUBSCRIPT)
        assert result == GSlidesBaselineOffset.SUBSCRIPT

    def test_none_to_gslides(self):
        """Abstract NONE should convert to None."""
        result = _convert_baseline_to_gslides(BaselineOffset.NONE)
        assert result is None

    def test_baseline_roundtrip_superscript(self):
        """SUPERSCRIPT should roundtrip correctly."""
        result = _convert_baseline_to_gslides(
            _convert_baseline_to_abstract(GSlidesBaselineOffset.SUPERSCRIPT)
        )
        assert result == GSlidesBaselineOffset.SUPERSCRIPT

    def test_baseline_roundtrip_subscript(self):
        """SUBSCRIPT should roundtrip correctly."""
        result = _convert_baseline_to_gslides(
            _convert_baseline_to_abstract(GSlidesBaselineOffset.SUBSCRIPT)
        )
        assert result == GSlidesBaselineOffset.SUBSCRIPT


class TestGSlidesToFull:
    """Tests for gslides_style_to_full conversion."""

    def test_none_style(self):
        """None TextStyle should return default FullTextStyle."""
        result = gslides_style_to_full(None)
        assert result == FullTextStyle()

    def test_bold_italic_strikethrough(self):
        """Bold, italic, strikethrough should go to markdown part."""
        style = TextStyle(bold=True, italic=True, strikethrough=True)
        result = gslides_style_to_full(style)
        assert result.markdown.bold is True
        assert result.markdown.italic is True
        assert result.markdown.strikethrough is True

    def test_monospace_font_is_code(self):
        """Monospace font should set is_code to True."""
        style = TextStyle(fontFamily="Courier New")
        result = gslides_style_to_full(style)
        assert result.markdown.is_code is True

    def test_non_monospace_font_not_code(self):
        """Non-monospace font should leave is_code as False."""
        style = TextStyle(fontFamily="Arial")
        result = gslides_style_to_full(style)
        assert result.markdown.is_code is False

    def test_hyperlink(self):
        """Link should go to markdown hyperlink."""
        style = TextStyle(link=Link(url="https://example.com"))
        result = gslides_style_to_full(style)
        assert result.markdown.hyperlink == "https://example.com"

    def test_font_family_to_rich(self):
        """Font family should go to rich part."""
        style = TextStyle(fontFamily="Arial")
        result = gslides_style_to_full(style)
        assert result.rich.font_family == "Arial"

    def test_font_size_pt(self):
        """Font size should convert to points in rich part."""
        style = TextStyle(fontSize=Dimension(magnitude=14.0, unit=Unit.PT))
        result = gslides_style_to_full(style)
        assert result.rich.font_size_pt == 14.0

    def test_font_weight(self):
        """Font weight should go to rich part."""
        style = TextStyle(
            weightedFontFamily=WeightedFontFamily(fontFamily="Arial", weight=700)
        )
        result = gslides_style_to_full(style)
        assert result.rich.font_weight == 700

    def test_foreground_color(self):
        """Foreground color should go to rich part."""
        style = TextStyle(
            foregroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=1.0, green=0.0, blue=0.0)
                )
            )
        )
        result = gslides_style_to_full(style)
        assert result.rich.foreground_color is not None
        assert result.rich.foreground_color.red == 1.0
        assert result.rich.foreground_color.green == 0.0
        assert result.rich.foreground_color.blue == 0.0

    def test_background_color(self):
        """Background color should go to rich part."""
        style = TextStyle(
            backgroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=1.0, green=1.0, blue=0.0)
                )
            )
        )
        result = gslides_style_to_full(style)
        assert result.rich.background_color is not None
        assert result.rich.background_color.red == 1.0
        assert result.rich.background_color.green == 1.0

    def test_underline(self):
        """Underline should go to rich part."""
        style = TextStyle(underline=True)
        result = gslides_style_to_full(style)
        assert result.rich.underline is True

    def test_small_caps(self):
        """Small caps should go to rich part."""
        style = TextStyle(smallCaps=True)
        result = gslides_style_to_full(style)
        assert result.rich.small_caps is True

    def test_baseline_offset(self):
        """Baseline offset should go to rich part."""
        style = TextStyle(baselineOffset=GSlidesBaselineOffset.SUPERSCRIPT)
        result = gslides_style_to_full(style)
        assert result.rich.baseline_offset == BaselineOffset.SUPERSCRIPT

    def test_full_style_all_properties(self):
        """Full style with all properties should convert correctly."""
        style = TextStyle(
            bold=True,
            italic=True,
            strikethrough=True,
            underline=True,
            smallCaps=True,
            fontFamily="Consolas",  # monospace
            fontSize=Dimension(magnitude=12.0, unit=Unit.PT),
            weightedFontFamily=WeightedFontFamily(fontFamily="Consolas", weight=400),
            foregroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=0.5, green=0.5, blue=0.5)
                )
            ),
            backgroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=1.0, green=1.0, blue=0.0)
                )
            ),
            baselineOffset=GSlidesBaselineOffset.SUBSCRIPT,
            link=Link(url="https://test.com"),
        )
        result = gslides_style_to_full(style)

        # Check markdown part
        assert result.markdown.bold is True
        assert result.markdown.italic is True
        assert result.markdown.strikethrough is True
        assert result.markdown.is_code is True  # Consolas is monospace
        assert result.markdown.hyperlink == "https://test.com"

        # Check rich part
        assert result.rich.font_family == "Consolas"
        assert result.rich.font_size_pt == 12.0
        assert result.rich.font_weight == 400
        assert result.rich.foreground_color.red == 0.5
        assert result.rich.background_color.red == 1.0
        assert result.rich.underline is True
        assert result.rich.small_caps is True
        assert result.rich.baseline_offset == BaselineOffset.SUBSCRIPT


class TestGSlidesToRich:
    """Tests for gslides_style_to_rich - should only extract RichStyle."""

    def test_extracts_only_rich(self):
        """Should extract only RichStyle, ignoring markdown properties."""
        style = TextStyle(
            bold=True,  # markdown - ignored
            italic=True,  # markdown - ignored
            fontFamily="Arial",  # rich
            fontSize=Dimension(magnitude=14.0, unit=Unit.PT),  # rich
        )
        result = gslides_style_to_rich(style)

        # Should be a RichStyle
        assert isinstance(result, RichStyle)

        # Should have rich properties
        assert result.font_family == "Arial"
        assert result.font_size_pt == 14.0

    def test_uniqueness_ignores_bold(self):
        """Two styles differing only in bold should produce equal RichStyle."""
        style1 = TextStyle(bold=True, fontFamily="Arial", fontSize=Dimension(magnitude=14.0, unit=Unit.PT))
        style2 = TextStyle(bold=False, fontFamily="Arial", fontSize=Dimension(magnitude=14.0, unit=Unit.PT))

        rich1 = gslides_style_to_rich(style1)
        rich2 = gslides_style_to_rich(style2)

        assert rich1 == rich2

    def test_uniqueness_ignores_italic(self):
        """Two styles differing only in italic should produce equal RichStyle."""
        style1 = TextStyle(italic=True, fontFamily="Arial")
        style2 = TextStyle(italic=False, fontFamily="Arial")

        rich1 = gslides_style_to_rich(style1)
        rich2 = gslides_style_to_rich(style2)

        assert rich1 == rich2

    def test_uniqueness_different_color_is_different(self):
        """Two styles with different colors should produce different RichStyle."""
        style1 = TextStyle(
            fontFamily="Arial",
            foregroundColor=OptionalColor(
                opaqueColor=Color(rgbColor=RgbColor(red=1.0, green=0.0, blue=0.0))
            ),
        )
        style2 = TextStyle(
            fontFamily="Arial",
            foregroundColor=OptionalColor(
                opaqueColor=Color(rgbColor=RgbColor(red=0.0, green=1.0, blue=0.0))
            ),
        )

        rich1 = gslides_style_to_rich(style1)
        rich2 = gslides_style_to_rich(style2)

        assert rich1 != rich2


class TestRichStyleToGSlides:
    """Tests for rich_style_to_gslides conversion."""

    def test_basic_rich_style(self):
        """Basic RichStyle should convert to TextStyle."""
        rich = RichStyle(
            font_family="Arial",
            font_size_pt=14.0,
        )
        result = rich_style_to_gslides(rich)

        assert result.fontFamily == "Arial"
        assert result.fontSize.magnitude == 14.0
        assert result.fontSize.unit == Unit.PT

    def test_with_markdown_style(self):
        """RichStyle with MarkdownRenderableStyle should combine."""
        rich = RichStyle(font_family="Arial")
        markdown = MarkdownRenderableStyle(bold=True, italic=True)

        result = rich_style_to_gslides(rich, markdown)

        assert result.fontFamily == "Arial"
        assert result.bold is True
        assert result.italic is True

    def test_code_without_font_sets_courier(self):
        """is_code=True without font_family should set Courier New."""
        rich = RichStyle()  # no font_family
        markdown = MarkdownRenderableStyle(is_code=True)

        result = rich_style_to_gslides(rich, markdown)

        assert result.fontFamily == "Courier New"

    def test_code_with_font_preserves_font(self):
        """is_code=True with font_family should preserve original font."""
        rich = RichStyle(font_family="Fira Code")
        markdown = MarkdownRenderableStyle(is_code=True)

        result = rich_style_to_gslides(rich, markdown)

        assert result.fontFamily == "Fira Code"

    def test_hyperlink(self):
        """Hyperlink should be set from markdown style."""
        rich = RichStyle()
        markdown = MarkdownRenderableStyle(hyperlink="https://example.com")

        result = rich_style_to_gslides(rich, markdown)

        assert result.link is not None
        assert result.link.url == "https://example.com"

    def test_foreground_color(self):
        """Foreground color should convert correctly."""
        rich = RichStyle(
            foreground_color=AbstractColor(red=1.0, green=0.0, blue=0.0)
        )
        result = rich_style_to_gslides(rich)

        assert result.foregroundColor is not None
        assert result.foregroundColor.opaqueColor.rgbColor.red == 1.0
        assert result.foregroundColor.opaqueColor.rgbColor.green == 0.0
        assert result.foregroundColor.opaqueColor.rgbColor.blue == 0.0

    def test_background_color(self):
        """Background color should convert correctly."""
        rich = RichStyle(
            background_color=AbstractColor(red=1.0, green=1.0, blue=0.0)
        )
        result = rich_style_to_gslides(rich)

        assert result.backgroundColor is not None
        assert result.backgroundColor.opaqueColor.rgbColor.red == 1.0
        assert result.backgroundColor.opaqueColor.rgbColor.green == 1.0

    def test_font_weight_creates_weighted_font_family(self):
        """Font weight should create WeightedFontFamily."""
        rich = RichStyle(font_family="Arial", font_weight=700)
        result = rich_style_to_gslides(rich)

        assert result.weightedFontFamily is not None
        assert result.weightedFontFamily.fontFamily == "Arial"
        assert result.weightedFontFamily.weight == 700

    def test_font_weight_without_family_uses_arial(self):
        """Font weight without family should default to Arial."""
        rich = RichStyle(font_weight=700)
        result = rich_style_to_gslides(rich)

        assert result.weightedFontFamily is not None
        assert result.weightedFontFamily.fontFamily == "Arial"
        assert result.weightedFontFamily.weight == 700

    def test_underline(self):
        """Underline should convert correctly."""
        rich = RichStyle(underline=True)
        result = rich_style_to_gslides(rich)
        assert result.underline is True

    def test_small_caps(self):
        """Small caps should convert correctly."""
        rich = RichStyle(small_caps=True)
        result = rich_style_to_gslides(rich)
        assert result.smallCaps is True

    def test_baseline_offset_superscript(self):
        """Superscript should convert correctly."""
        rich = RichStyle(baseline_offset=BaselineOffset.SUPERSCRIPT)
        result = rich_style_to_gslides(rich)
        assert result.baselineOffset == GSlidesBaselineOffset.SUPERSCRIPT

    def test_baseline_offset_subscript(self):
        """Subscript should convert correctly."""
        rich = RichStyle(baseline_offset=BaselineOffset.SUBSCRIPT)
        result = rich_style_to_gslides(rich)
        assert result.baselineOffset == GSlidesBaselineOffset.SUBSCRIPT

    def test_false_values_become_none(self):
        """False values should become None (not set explicitly)."""
        rich = RichStyle()
        markdown = MarkdownRenderableStyle(bold=False, italic=False)

        result = rich_style_to_gslides(rich, markdown)

        assert result.bold is None
        assert result.italic is None
        assert result.strikethrough is None


class TestFullStyleToGSlides:
    """Tests for full_style_to_gslides conversion."""

    def test_combines_markdown_and_rich(self):
        """FullTextStyle should combine both parts."""
        full = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True, hyperlink="https://test.com"),
            rich=RichStyle(font_family="Arial", font_size_pt=14.0),
        )
        result = full_style_to_gslides(full)

        assert result.bold is True
        assert result.link.url == "https://test.com"
        assert result.fontFamily == "Arial"
        assert result.fontSize.magnitude == 14.0


class TestMarkdownStyleToGSlides:
    """Tests for markdown_style_to_gslides conversion."""

    def test_bold(self):
        """Bold should convert."""
        markdown = MarkdownRenderableStyle(bold=True)
        result = markdown_style_to_gslides(markdown)
        assert result.bold is True

    def test_italic(self):
        """Italic should convert."""
        markdown = MarkdownRenderableStyle(italic=True)
        result = markdown_style_to_gslides(markdown)
        assert result.italic is True

    def test_strikethrough(self):
        """Strikethrough should convert."""
        markdown = MarkdownRenderableStyle(strikethrough=True)
        result = markdown_style_to_gslides(markdown)
        assert result.strikethrough is True

    def test_code_sets_courier_new(self):
        """is_code should set Courier New font."""
        markdown = MarkdownRenderableStyle(is_code=True)
        result = markdown_style_to_gslides(markdown)
        assert result.fontFamily == "Courier New"

    def test_hyperlink(self):
        """Hyperlink should convert."""
        markdown = MarkdownRenderableStyle(hyperlink="https://example.com")
        result = markdown_style_to_gslides(markdown)
        assert result.link is not None
        assert result.link.url == "https://example.com"

    def test_no_rich_properties(self):
        """Should not include any rich properties."""
        markdown = MarkdownRenderableStyle(bold=True, italic=True)
        result = markdown_style_to_gslides(markdown)

        # These should be None since we only converted markdown
        assert result.fontSize is None
        assert result.foregroundColor is None
        assert result.backgroundColor is None
        assert result.underline is None
        assert result.smallCaps is None
        assert result.weightedFontFamily is None


class TestRoundTrip:
    """Tests for round-trip conversion: GSlides -> Abstract -> GSlides."""

    def test_roundtrip_full_style(self):
        """Full style should roundtrip correctly."""
        original = TextStyle(
            bold=True,
            italic=True,
            strikethrough=True,
            underline=True,
            smallCaps=True,
            fontFamily="Arial",
            fontSize=Dimension(magnitude=14.0, unit=Unit.PT),
            weightedFontFamily=WeightedFontFamily(fontFamily="Arial", weight=700),
            foregroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=1.0, green=0.0, blue=0.0)
                )
            ),
            backgroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=1.0, green=1.0, blue=0.0)
                )
            ),
            baselineOffset=GSlidesBaselineOffset.SUPERSCRIPT,
            link=Link(url="https://example.com"),
        )

        # Convert to FullTextStyle and back
        full = gslides_style_to_full(original)
        result = full_style_to_gslides(full)

        # Check all properties preserved
        assert result.bold is True
        assert result.italic is True
        assert result.strikethrough is True
        assert result.underline is True
        assert result.smallCaps is True
        assert result.fontFamily == "Arial"
        assert result.fontSize.magnitude == 14.0
        assert result.weightedFontFamily.fontFamily == "Arial"
        assert result.weightedFontFamily.weight == 700
        assert result.foregroundColor.opaqueColor.rgbColor.red == 1.0
        assert result.backgroundColor.opaqueColor.rgbColor.green == 1.0
        assert result.baselineOffset == GSlidesBaselineOffset.SUPERSCRIPT
        assert result.link.url == "https://example.com"

    def test_roundtrip_empty_style(self):
        """Empty style should roundtrip correctly."""
        original = TextStyle()
        full = gslides_style_to_full(original)
        result = full_style_to_gslides(full)

        # Should be essentially empty
        assert result.bold is None
        assert result.italic is None
        assert result.fontFamily is None

    def test_roundtrip_preserves_colors(self):
        """Colors should roundtrip precisely."""
        original = TextStyle(
            foregroundColor=OptionalColor(
                opaqueColor=Color(
                    rgbColor=RgbColor(red=0.123, green=0.456, blue=0.789)
                )
            )
        )

        full = gslides_style_to_full(original)
        result = full_style_to_gslides(full)

        assert result.foregroundColor.opaqueColor.rgbColor.red == 0.123
        assert result.foregroundColor.opaqueColor.rgbColor.green == 0.456
        assert result.foregroundColor.opaqueColor.rgbColor.blue == 0.789

    def test_roundtrip_font_size_emu(self):
        """EMU font size should roundtrip (converted to PT)."""
        # 127000 EMUs = 10 points
        original = TextStyle(fontSize=Dimension(magnitude=127000.0, unit=Unit.EMU))

        full = gslides_style_to_full(original)
        result = full_style_to_gslides(full)

        # Result will be in PT (10 points)
        assert result.fontSize.magnitude == 10.0
        assert result.fontSize.unit == Unit.PT
