"""Tests for platform-agnostic text style classes."""

import pytest

from gslides_api.agnostic.text import (
    AbstractColor,
    AbstractTextRun,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    ParagraphStyle,
    RichStyle,
    SpacingValue,
)


class TestAbstractColor:
    """Tests for AbstractColor class."""

    def test_default_color_is_black(self):
        """Default color should be black with full opacity."""
        color = AbstractColor()
        assert color.red == 0.0
        assert color.green == 0.0
        assert color.blue == 0.0
        assert color.alpha == 1.0

    def test_from_rgb_tuple(self):
        """Should convert from 0-255 RGB tuple."""
        color = AbstractColor.from_rgb_tuple((255, 128, 0))
        assert color.red == 1.0
        assert color.green == pytest.approx(128 / 255, rel=1e-3)
        assert color.blue == 0.0

    def test_from_rgb_float(self):
        """Should create from 0.0-1.0 floats."""
        color = AbstractColor.from_rgb_float(0.5, 0.25, 0.75)
        assert color.red == 0.5
        assert color.green == 0.25
        assert color.blue == 0.75

    def test_to_rgb_tuple(self):
        """Should convert to 0-255 RGB tuple."""
        color = AbstractColor(red=1.0, green=0.5, blue=0.0)
        r, g, b = color.to_rgb_tuple()
        assert r == 255
        assert g == 127  # int(0.5 * 255) = 127
        assert b == 0

    def test_to_hex(self):
        """Should convert to hex string."""
        color = AbstractColor(red=1.0, green=0.0, blue=0.0)
        assert color.to_hex() == "#ff0000"

        color2 = AbstractColor(red=0.0, green=1.0, blue=0.0)
        assert color2.to_hex() == "#00ff00"

        color3 = AbstractColor(red=0.0, green=0.0, blue=1.0)
        assert color3.to_hex() == "#0000ff"

    def test_roundtrip_rgb_tuple(self):
        """Should roundtrip through RGB tuple conversion."""
        original = (128, 64, 192)
        color = AbstractColor.from_rgb_tuple(original)
        result = color.to_rgb_tuple()
        assert result == original

    def test_color_equality(self):
        """Colors with same values should be equal."""
        c1 = AbstractColor(red=0.5, green=0.5, blue=0.5)
        c2 = AbstractColor(red=0.5, green=0.5, blue=0.5)
        assert c1 == c2

    def test_color_inequality(self):
        """Colors with different values should not be equal."""
        c1 = AbstractColor(red=0.5, green=0.5, blue=0.5)
        c2 = AbstractColor(red=0.6, green=0.5, blue=0.5)
        assert c1 != c2


class TestMarkdownRenderableStyle:
    """Tests for MarkdownRenderableStyle class."""

    def test_default_values(self):
        """All defaults should be False/None."""
        style = MarkdownRenderableStyle()
        assert style.bold is False
        assert style.italic is False
        assert style.strikethrough is False
        assert style.is_code is False
        assert style.hyperlink is None

    def test_bold_style(self):
        """Should support bold styling."""
        style = MarkdownRenderableStyle(bold=True)
        assert style.bold is True

    def test_combined_styles(self):
        """Should support combining multiple styles."""
        style = MarkdownRenderableStyle(bold=True, italic=True, strikethrough=True)
        assert style.bold is True
        assert style.italic is True
        assert style.strikethrough is True

    def test_hyperlink(self):
        """Should support hyperlinks."""
        style = MarkdownRenderableStyle(hyperlink="https://example.com")
        assert style.hyperlink == "https://example.com"

    def test_equality(self):
        """Styles with same values should be equal."""
        s1 = MarkdownRenderableStyle(bold=True, italic=True)
        s2 = MarkdownRenderableStyle(bold=True, italic=True)
        assert s1 == s2

    def test_inequality(self):
        """Styles with different values should not be equal."""
        s1 = MarkdownRenderableStyle(bold=True)
        s2 = MarkdownRenderableStyle(bold=False)
        assert s1 != s2


class TestRichStyle:
    """Tests for RichStyle class."""

    def test_default_values(self):
        """All defaults should be None/False/NONE."""
        style = RichStyle()
        assert style.font_family is None
        assert style.font_size_pt is None
        assert style.font_weight is None
        assert style.foreground_color is None
        assert style.background_color is None
        assert style.underline is False
        assert style.small_caps is False
        assert style.all_caps is False
        assert style.baseline_offset == BaselineOffset.NONE
        assert style.character_spacing is None
        assert style.shadow is False
        assert style.emboss is False
        assert style.imprint is False
        assert style.double_strike is False

    def test_font_properties(self):
        """Should support font properties."""
        style = RichStyle(
            font_family="Arial",
            font_size_pt=14.0,
            font_weight=700,
        )
        assert style.font_family == "Arial"
        assert style.font_size_pt == 14.0
        assert style.font_weight == 700

    def test_color_properties(self):
        """Should support color properties."""
        fg_color = AbstractColor(red=1.0, green=0.0, blue=0.0)
        bg_color = AbstractColor(red=1.0, green=1.0, blue=0.0)
        style = RichStyle(foreground_color=fg_color, background_color=bg_color)
        assert style.foreground_color == fg_color
        assert style.background_color == bg_color

    def test_is_monospace_common_fonts(self):
        """Should detect common monospace fonts."""
        monospace_fonts = [
            "Courier New",
            "courier",
            "MONOSPACE",
            "Consolas",
            "Monaco",
            "Lucida Console",
            "DejaVu Sans Mono",
            "Source Code Pro",
            "Fira Code",
            "JetBrains Mono",
        ]
        for font in monospace_fonts:
            style = RichStyle(font_family=font)
            assert style.is_monospace() is True, f"{font} should be detected as monospace"

    def test_is_monospace_non_monospace(self):
        """Should not detect non-monospace fonts as monospace."""
        non_monospace_fonts = ["Arial", "Times New Roman", "Helvetica", "Georgia"]
        for font in non_monospace_fonts:
            style = RichStyle(font_family=font)
            assert style.is_monospace() is False, f"{font} should not be detected as monospace"

    def test_is_monospace_none_font(self):
        """Should return False when font_family is None."""
        style = RichStyle(font_family=None)
        assert style.is_monospace() is False

    def test_baseline_offset_superscript(self):
        """Should support superscript baseline offset."""
        style = RichStyle(baseline_offset=BaselineOffset.SUPERSCRIPT)
        assert style.baseline_offset == BaselineOffset.SUPERSCRIPT

    def test_baseline_offset_subscript(self):
        """Should support subscript baseline offset."""
        style = RichStyle(baseline_offset=BaselineOffset.SUBSCRIPT)
        assert style.baseline_offset == BaselineOffset.SUBSCRIPT

    def test_equality_same_properties(self):
        """RichStyles with same properties should be equal."""
        s1 = RichStyle(font_family="Arial", font_size_pt=14.0, underline=True)
        s2 = RichStyle(font_family="Arial", font_size_pt=14.0, underline=True)
        assert s1 == s2

    def test_equality_different_properties(self):
        """RichStyles with different properties should not be equal."""
        s1 = RichStyle(font_family="Arial", font_size_pt=14.0)
        s2 = RichStyle(font_family="Arial", font_size_pt=16.0)
        assert s1 != s2

    def test_equality_ignores_none_vs_default(self):
        """RichStyles should compare properly with explicit vs implicit None."""
        s1 = RichStyle()
        s2 = RichStyle(font_family=None)
        assert s1 == s2

    def test_is_default_empty_style(self):
        """Empty RichStyle should be default."""
        style = RichStyle()
        assert style.is_default() is True

    def test_is_default_with_font_family(self):
        """RichStyle with font_family should not be default."""
        style = RichStyle(font_family="Arial")
        assert style.is_default() is False

    def test_is_default_with_font_size(self):
        """RichStyle with font_size_pt should not be default."""
        style = RichStyle(font_size_pt=14.0)
        assert style.is_default() is False

    def test_is_default_with_color(self):
        """RichStyle with foreground_color should not be default."""
        style = RichStyle(foreground_color=AbstractColor(red=1.0, green=0.0, blue=0.0))
        assert style.is_default() is False

    def test_is_default_with_underline(self):
        """RichStyle with underline should not be default."""
        style = RichStyle(underline=True)
        assert style.is_default() is False

    def test_is_default_with_baseline_offset(self):
        """RichStyle with baseline_offset should not be default."""
        style = RichStyle(baseline_offset=BaselineOffset.SUPERSCRIPT)
        assert style.is_default() is False


class TestFullTextStyle:
    """Tests for FullTextStyle class."""

    def test_default_values(self):
        """Should have default MarkdownRenderableStyle and RichStyle."""
        style = FullTextStyle()
        assert style.markdown == MarkdownRenderableStyle()
        assert style.rich == RichStyle()

    def test_with_markdown_style(self):
        """Should support setting markdown style."""
        md_style = MarkdownRenderableStyle(bold=True, italic=True)
        style = FullTextStyle(markdown=md_style)
        assert style.markdown.bold is True
        assert style.markdown.italic is True

    def test_with_rich_style(self):
        """Should support setting rich style."""
        rich_style = RichStyle(font_family="Arial", font_size_pt=14.0)
        style = FullTextStyle(rich=rich_style)
        assert style.rich.font_family == "Arial"
        assert style.rich.font_size_pt == 14.0

    def test_with_both_styles(self):
        """Should support setting both styles."""
        md_style = MarkdownRenderableStyle(bold=True)
        rich_style = RichStyle(font_family="Arial")
        style = FullTextStyle(markdown=md_style, rich=rich_style)
        assert style.markdown.bold is True
        assert style.rich.font_family == "Arial"

    def test_equality(self):
        """FullTextStyles with same components should be equal."""
        s1 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=RichStyle(font_family="Arial"),
        )
        s2 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=RichStyle(font_family="Arial"),
        )
        assert s1 == s2


class TestAbstractTextRun:
    """Tests for AbstractTextRun class."""

    def test_content_required(self):
        """Should require content."""
        run = AbstractTextRun(content="Hello")
        assert run.content == "Hello"

    def test_default_style(self):
        """Should have default FullTextStyle."""
        run = AbstractTextRun(content="Hello")
        assert run.style == FullTextStyle()

    def test_with_style(self):
        """Should support custom style."""
        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=RichStyle(font_family="Arial"),
        )
        run = AbstractTextRun(content="Hello", style=style)
        assert run.content == "Hello"
        assert run.style.markdown.bold is True
        assert run.style.rich.font_family == "Arial"

    def test_equality(self):
        """TextRuns with same content and style should be equal."""
        r1 = AbstractTextRun(content="Hello", style=FullTextStyle())
        r2 = AbstractTextRun(content="Hello", style=FullTextStyle())
        assert r1 == r2

    def test_inequality_content(self):
        """TextRuns with different content should not be equal."""
        r1 = AbstractTextRun(content="Hello")
        r2 = AbstractTextRun(content="World")
        assert r1 != r2

    def test_inequality_style(self):
        """TextRuns with different styles should not be equal."""
        r1 = AbstractTextRun(content="Hello", style=FullTextStyle())
        r2 = AbstractTextRun(
            content="Hello",
            style=FullTextStyle(markdown=MarkdownRenderableStyle(bold=True)),
        )
        assert r1 != r2


class TestStyleUniqueness:
    """Tests for style uniqueness checking behavior.

    The key design principle is that styles() should return only RichStyle,
    so text that differs only in bold/italic/strikethrough is ONE style.
    """

    def test_rich_style_uniqueness_ignores_markdown_properties(self):
        """Two texts with same RichStyle but different markdown formatting
        should be considered as having the same style for uniqueness purposes.
        """
        # Same rich style (Arial 14pt red)
        rich_style = RichStyle(
            font_family="Arial",
            font_size_pt=14.0,
            foreground_color=AbstractColor(red=1.0, green=0.0, blue=0.0),
        )

        # Different markdown styles
        full_style_1 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=False),
            rich=rich_style,
        )
        full_style_2 = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True),
            rich=rich_style,
        )

        # The RichStyle parts should be equal
        assert full_style_1.rich == full_style_2.rich

        # The FullTextStyles are different (because markdown differs)
        assert full_style_1 != full_style_2

        # But for styles() purposes, we only compare rich
        styles = []
        if full_style_1.rich not in styles:
            styles.append(full_style_1.rich)
        if full_style_2.rich not in styles:
            styles.append(full_style_2.rich)

        # Should only be ONE unique style
        assert len(styles) == 1

    def test_different_rich_styles_are_different(self):
        """Different rich styles should be considered different."""
        rich_style_1 = RichStyle(font_family="Arial", font_size_pt=14.0)
        rich_style_2 = RichStyle(font_family="Arial", font_size_pt=16.0)

        styles = []
        if rich_style_1 not in styles:
            styles.append(rich_style_1)
        if rich_style_2 not in styles:
            styles.append(rich_style_2)

        # Should be TWO unique styles
        assert len(styles) == 2


class TestSpacingValue:
    """Tests for SpacingValue class."""

    def test_default_values(self):
        """Should have None defaults."""
        spacing = SpacingValue()
        assert spacing.points is None
        assert spacing.percentage is None

    def test_from_pptx_pts(self):
        """Should create from PPTX spcPts val (100ths of a point)."""
        spacing = SpacingValue.from_pptx_pts("900")  # 9pt
        assert spacing.points == 9.0
        assert spacing.percentage is None

    def test_from_pptx_pct(self):
        """Should create from PPTX spcPct val (1/1000ths of percent)."""
        spacing = SpacingValue.from_pptx_pct("110000")  # 110%
        assert spacing.percentage == 1.1
        assert spacing.points is None

    def test_to_pptx_pts(self):
        """Should convert to PPTX spcPts val."""
        spacing = SpacingValue(points=9.0)
        assert spacing.to_pptx_pts() == "900"

    def test_to_pptx_pct(self):
        """Should convert to PPTX spcPct val."""
        spacing = SpacingValue(percentage=1.1)
        assert spacing.to_pptx_pct() == "110000"

    def test_roundtrip_pts(self):
        """Should roundtrip through PPTX pts format."""
        original = SpacingValue(points=12.0)
        roundtripped = SpacingValue.from_pptx_pts(original.to_pptx_pts())
        assert roundtripped.points == 12.0

    def test_roundtrip_pct(self):
        """Should roundtrip through PPTX pct format."""
        original = SpacingValue(percentage=1.15)
        roundtripped = SpacingValue.from_pptx_pct(original.to_pptx_pct())
        # Allow small floating point difference
        assert abs(roundtripped.percentage - 1.15) < 0.00001


class TestParagraphAlignment:
    """Tests for ParagraphAlignment enum."""

    def test_alignment_values(self):
        """Should have correct PPTX values."""
        from gslides_api.agnostic.text import ParagraphAlignment

        assert ParagraphAlignment.LEFT.value == "l"
        assert ParagraphAlignment.CENTER.value == "ctr"
        assert ParagraphAlignment.RIGHT.value == "r"
        assert ParagraphAlignment.JUSTIFIED.value == "just"


class TestParagraphStyle:
    """Tests for ParagraphStyle class."""

    def test_default_values(self):
        """Should have None defaults."""
        style = ParagraphStyle()
        assert style.margin_left is None
        assert style.margin_right is None
        assert style.indent is None
        assert style.alignment is None
        assert style.line_spacing is None
        assert style.space_before is None
        assert style.space_after is None
        assert style.level is None

    def test_with_margins_and_indent(self):
        """Should support margin and indent values."""
        style = ParagraphStyle(
            margin_left=520700,
            margin_right=0,
            indent=-209550,
        )
        assert style.margin_left == 520700
        assert style.margin_right == 0
        assert style.indent == -209550

    def test_with_spacing(self):
        """Should support spacing values."""
        style = ParagraphStyle(
            space_before=SpacingValue(points=9.0),
            space_after=SpacingValue(points=0.0),
            line_spacing=SpacingValue(percentage=1.1),
        )
        assert style.space_before.points == 9.0
        assert style.space_after.points == 0.0
        assert style.line_spacing.percentage == 1.1

    def test_has_bullet_properties_true(self):
        """Should detect bullet paragraph properties."""
        style = ParagraphStyle(
            margin_left=520700,
            indent=-209550,
        )
        assert style.has_bullet_properties() is True

    def test_has_bullet_properties_false_no_margin(self):
        """Should not detect bullet without margin."""
        style = ParagraphStyle(indent=-209550)
        assert style.has_bullet_properties() is False

    def test_has_bullet_properties_false_no_indent(self):
        """Should not detect bullet without indent."""
        style = ParagraphStyle(margin_left=520700)
        assert style.has_bullet_properties() is False

    def test_has_bullet_properties_false_positive_indent(self):
        """Should not detect bullet with positive indent."""
        style = ParagraphStyle(margin_left=520700, indent=100000)
        assert style.has_bullet_properties() is False

    def test_with_alignment(self):
        """Should support alignment."""
        from gslides_api.agnostic.text import ParagraphAlignment

        style = ParagraphStyle(alignment=ParagraphAlignment.CENTER)
        assert style.alignment == ParagraphAlignment.CENTER
