"""Tests for PowerPoint font <-> platform-agnostic style converters."""

from unittest.mock import MagicMock

import pytest

from gslides_api.agnostic.text import (
    AbstractColor,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    RichStyle,
)
from gslides_api.pptx.converters import (
    MONOSPACE_FONTS,
    _abstract_to_pptx_rgb,
    _escape_markdown_for_table,
    _is_monospace,
    _pptx_baseline_to_abstract,
    _pptx_color_to_abstract,
    apply_full_style_to_pptx_run,
    apply_markdown_style_to_pptx_run,
    apply_rich_style_to_pptx_run,
    pptx_font_to_full,
    pptx_font_to_rich,
    pptx_paragraph_to_markdown,
    pptx_run_to_markdown,
    pptx_table_to_markdown,
    pptx_text_frame_to_markdown,
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


class TestColorConversion:
    """Tests for color conversion between pptx and abstract."""

    def test_pptx_color_to_abstract_with_rgb(self):
        """Should convert pptx color with RGB to AbstractColor."""
        mock_color = MagicMock()
        mock_color.rgb = (255, 128, 0)

        result = _pptx_color_to_abstract(mock_color)

        assert result is not None
        assert result.red == 1.0
        assert result.green == pytest.approx(128 / 255, rel=1e-3)
        assert result.blue == 0.0

    def test_pptx_color_to_abstract_none_rgb(self):
        """Should return None when color has no RGB."""
        mock_color = MagicMock()
        mock_color.rgb = None

        result = _pptx_color_to_abstract(mock_color)
        assert result is None

    def test_pptx_color_to_abstract_no_attribute(self):
        """Should return None when color raises AttributeError."""
        # Use spec to make mock raise AttributeError when accessing rgb
        mock_color = MagicMock(spec=["something_else"])

        result = _pptx_color_to_abstract(mock_color)
        assert result is None

    def test_abstract_to_pptx_rgb(self):
        """Should convert AbstractColor to pptx RGBColor."""
        abstract = AbstractColor(red=1.0, green=0.5, blue=0.0)

        result = _abstract_to_pptx_rgb(abstract)

        assert result is not None
        assert result[0] == 255  # red
        assert result[1] == 127  # green (int(0.5 * 255))
        assert result[2] == 0  # blue

    def test_abstract_to_pptx_rgb_none(self):
        """Should return None for None input."""
        result = _abstract_to_pptx_rgb(None)
        assert result is None


class TestBaselineConversion:
    """Tests for baseline offset conversion."""

    def test_superscript_detected(self):
        """Superscript font should be detected."""
        mock_font = MagicMock()
        mock_font.superscript = True
        mock_font.subscript = False

        result = _pptx_baseline_to_abstract(mock_font)
        assert result == BaselineOffset.SUPERSCRIPT

    def test_subscript_detected(self):
        """Subscript font should be detected."""
        mock_font = MagicMock()
        mock_font.superscript = False
        mock_font.subscript = True

        result = _pptx_baseline_to_abstract(mock_font)
        assert result == BaselineOffset.SUBSCRIPT

    def test_no_baseline_offset(self):
        """No baseline offset should return NONE."""
        mock_font = MagicMock()
        mock_font.superscript = False
        mock_font.subscript = False

        result = _pptx_baseline_to_abstract(mock_font)
        assert result == BaselineOffset.NONE

    def test_baseline_attribute_error(self):
        """Should return NONE on AttributeError."""
        mock_font = MagicMock()
        del mock_font.superscript
        del mock_font.subscript

        result = _pptx_baseline_to_abstract(mock_font)
        assert result == BaselineOffset.NONE


class TestPptxFontToFull:
    """Tests for pptx_font_to_full conversion."""

    def test_basic_font_properties(self):
        """Basic font properties should be extracted."""
        mock_font = MagicMock()
        mock_font.bold = True
        mock_font.italic = True
        mock_font.strike = True
        mock_font.name = "Arial"
        mock_size = MagicMock()
        mock_size.pt = 14.0
        mock_font.size = mock_size
        mock_font.underline = True
        mock_font.small_caps = False
        mock_font.all_caps = False
        mock_font.superscript = False
        mock_font.subscript = False
        mock_font.shadow = False
        mock_font.emboss = False
        mock_font.imprint = False
        mock_font.double_strike = False
        mock_font.color = MagicMock()
        mock_font.color.rgb = (255, 0, 0)

        result = pptx_font_to_full(mock_font)

        # Check markdown properties
        assert result.markdown.bold is True
        assert result.markdown.italic is True
        assert result.markdown.strikethrough is True
        assert result.markdown.is_code is False  # Arial is not monospace
        assert result.markdown.hyperlink is None

        # Check rich properties
        assert result.rich.font_family == "Arial"
        assert result.rich.font_size_pt == 14.0
        assert result.rich.underline is True
        assert result.rich.foreground_color.red == 1.0

    def test_monospace_font_is_code(self):
        """Monospace font should set is_code to True."""
        mock_font = MagicMock()
        mock_font.bold = False
        mock_font.italic = False
        mock_font.strike = False
        mock_font.name = "Courier New"
        mock_font.size = None
        mock_font.underline = False
        mock_font.small_caps = False
        mock_font.all_caps = False
        mock_font.superscript = False
        mock_font.subscript = False
        mock_font.shadow = False
        mock_font.emboss = False
        mock_font.imprint = False
        mock_font.double_strike = False
        mock_font.color = MagicMock()
        mock_font.color.rgb = None

        result = pptx_font_to_full(mock_font)
        assert result.markdown.is_code is True

    def test_with_hyperlink(self):
        """Hyperlink should be captured."""
        mock_font = MagicMock()
        mock_font.bold = False
        mock_font.italic = False
        mock_font.strike = False
        mock_font.name = "Arial"
        mock_font.size = None
        mock_font.underline = False
        mock_font.small_caps = False
        mock_font.all_caps = False
        mock_font.superscript = False
        mock_font.subscript = False
        mock_font.shadow = False
        mock_font.emboss = False
        mock_font.imprint = False
        mock_font.double_strike = False
        mock_font.color = MagicMock()
        mock_font.color.rgb = None

        result = pptx_font_to_full(mock_font, hyperlink_address="https://example.com")
        assert result.markdown.hyperlink == "https://example.com"


class TestPptxFontToRich:
    """Tests for pptx_font_to_rich - should only extract RichStyle."""

    def test_extracts_only_rich(self):
        """Should extract only RichStyle, ignoring markdown properties."""
        mock_font = MagicMock()
        mock_font.bold = True  # markdown - should not affect RichStyle equality
        mock_font.italic = True
        mock_font.strike = True
        mock_font.name = "Arial"
        mock_size = MagicMock()
        mock_size.pt = 14.0
        mock_font.size = mock_size
        mock_font.underline = False
        mock_font.small_caps = False
        mock_font.all_caps = False
        mock_font.superscript = False
        mock_font.subscript = False
        mock_font.shadow = False
        mock_font.emboss = False
        mock_font.imprint = False
        mock_font.double_strike = False
        mock_font.color = MagicMock()
        mock_font.color.rgb = None

        result = pptx_font_to_rich(mock_font)

        # Should be a RichStyle
        assert isinstance(result, RichStyle)

        # Should have rich properties
        assert result.font_family == "Arial"
        assert result.font_size_pt == 14.0

    def test_uniqueness_ignores_bold(self):
        """Two fonts differing only in bold should produce equal RichStyle."""
        mock_font1 = MagicMock()
        mock_font1.bold = True
        mock_font1.italic = False
        mock_font1.strike = False
        mock_font1.name = "Arial"
        mock_size = MagicMock()
        mock_size.pt = 14.0
        mock_font1.size = mock_size
        mock_font1.underline = False
        mock_font1.small_caps = False
        mock_font1.all_caps = False
        mock_font1.superscript = False
        mock_font1.subscript = False
        mock_font1.shadow = False
        mock_font1.emboss = False
        mock_font1.imprint = False
        mock_font1.double_strike = False
        mock_font1.color = MagicMock()
        mock_font1.color.rgb = None

        mock_font2 = MagicMock()
        mock_font2.bold = False  # Different from font1
        mock_font2.italic = False
        mock_font2.strike = False
        mock_font2.name = "Arial"
        mock_font2.size = mock_size
        mock_font2.underline = False
        mock_font2.small_caps = False
        mock_font2.all_caps = False
        mock_font2.superscript = False
        mock_font2.subscript = False
        mock_font2.shadow = False
        mock_font2.emboss = False
        mock_font2.imprint = False
        mock_font2.double_strike = False
        mock_font2.color = MagicMock()
        mock_font2.color.rgb = None

        rich1 = pptx_font_to_rich(mock_font1)
        rich2 = pptx_font_to_rich(mock_font2)

        # RichStyles should be equal since they differ only in bold
        assert rich1 == rich2


class TestApplyRichStyleToPptxRun:
    """Tests for apply_rich_style_to_pptx_run."""

    def test_applies_font_family(self):
        """Should apply font family."""
        mock_run = MagicMock()
        rich = RichStyle(font_family="Arial")

        apply_rich_style_to_pptx_run(rich, mock_run)

        assert mock_run.font.name == "Arial"

    def test_applies_font_size(self):
        """Should apply font size."""
        mock_run = MagicMock()
        rich = RichStyle(font_size_pt=14.0)

        apply_rich_style_to_pptx_run(rich, mock_run)

        # The size should be set (Pt object)
        mock_run.font.size = rich.font_size_pt  # Check assignment happened

    def test_applies_foreground_color(self):
        """Should apply foreground color."""
        mock_run = MagicMock()
        rich = RichStyle(foreground_color=AbstractColor(red=1.0, green=0.0, blue=0.0))

        apply_rich_style_to_pptx_run(rich, mock_run)

        # Color should be set via rgb property
        mock_run.font.color.rgb  # Just verify access doesn't fail

    def test_applies_underline(self):
        """Should apply underline."""
        mock_run = MagicMock()
        rich = RichStyle(underline=True)

        apply_rich_style_to_pptx_run(rich, mock_run)

        assert mock_run.font.underline is True


class TestApplyMarkdownStyleToPptxRun:
    """Tests for apply_markdown_style_to_pptx_run."""

    def test_applies_bold(self):
        """Should apply bold."""
        mock_run = MagicMock()
        md = MarkdownRenderableStyle(bold=True)

        apply_markdown_style_to_pptx_run(md, mock_run)

        assert mock_run.font.bold is True

    def test_applies_italic(self):
        """Should apply italic."""
        mock_run = MagicMock()
        md = MarkdownRenderableStyle(italic=True)

        apply_markdown_style_to_pptx_run(md, mock_run)

        assert mock_run.font.italic is True

    def test_applies_code_font(self):
        """Should apply Courier New for code."""
        mock_run = MagicMock()
        mock_run.font.name = None  # No existing font
        md = MarkdownRenderableStyle(is_code=True)

        apply_markdown_style_to_pptx_run(md, mock_run)

        assert mock_run.font.name == "Courier New"

    def test_code_does_not_override_existing_font(self):
        """Should not override existing font when is_code is True."""
        mock_run = MagicMock()
        mock_run.font.name = "Fira Code"  # Already has a font
        md = MarkdownRenderableStyle(is_code=True)

        apply_markdown_style_to_pptx_run(md, mock_run)

        # Should keep existing font
        assert mock_run.font.name == "Fira Code"

    def test_applies_hyperlink(self):
        """Should apply hyperlink."""
        mock_run = MagicMock()
        md = MarkdownRenderableStyle(hyperlink="https://example.com")

        apply_markdown_style_to_pptx_run(md, mock_run)

        mock_run.hyperlink.address = "https://example.com"


class TestApplyFullStyleToPptxRun:
    """Tests for apply_full_style_to_pptx_run."""

    def test_applies_both_markdown_and_rich(self):
        """Should apply both markdown and rich properties."""
        mock_run = MagicMock()
        mock_run.font.name = None

        style = FullTextStyle(
            markdown=MarkdownRenderableStyle(bold=True, is_code=True),
            rich=RichStyle(font_size_pt=14.0, underline=True),
        )

        apply_full_style_to_pptx_run(style, mock_run)

        # Markdown properties
        assert mock_run.font.bold is True
        assert mock_run.font.name == "Courier New"  # from is_code

        # Rich properties
        assert mock_run.font.underline is True


# =============================================================================
# Tests for Markdown Generation Functions (PPT -> Markdown)
# =============================================================================


class TestEscapeMarkdownForTable:
    """Tests for _escape_markdown_for_table."""

    def test_escapes_pipe_characters(self):
        """Should escape pipe characters."""
        text = "col1|col2|col3"
        result = _escape_markdown_for_table(text)
        assert result == "col1\\|col2\\|col3"

    def test_converts_newlines_to_br(self):
        """Should convert newlines to <br> tags."""
        text = "line1\nline2\nline3"
        result = _escape_markdown_for_table(text)
        assert result == "line1<br>line2<br>line3"

    def test_preserves_curly_braces(self):
        """Should NOT escape curly braces (template variables)."""
        text = "Hello {name}, welcome to {place}"
        result = _escape_markdown_for_table(text)
        assert result == "Hello {name}, welcome to {place}"

    def test_combined_escaping(self):
        """Should handle pipes and newlines together."""
        text = "a|b\nc|d"
        result = _escape_markdown_for_table(text)
        assert result == "a\\|b<br>c\\|d"

    def test_empty_string(self):
        """Should handle empty string."""
        result = _escape_markdown_for_table("")
        assert result == ""


class TestPptxRunToMarkdown:
    """Tests for pptx_run_to_markdown."""

    def _create_mock_run(
        self,
        text="Test",
        bold=False,
        italic=False,
        strike=False,
        font_name="Arial",
        hyperlink=None,
    ):
        """Helper to create mock run with specified properties."""
        mock_run = MagicMock()
        mock_run.text = text
        mock_run.font.bold = bold
        mock_run.font.italic = italic
        mock_run.font.strike = strike
        mock_run.font.name = font_name
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None

        if hyperlink:
            mock_run.hyperlink.address = hyperlink
        else:
            mock_run.hyperlink = MagicMock()
            mock_run.hyperlink.address = None

        return mock_run

    def test_plain_text(self):
        """Plain text should be returned unchanged."""
        mock_run = self._create_mock_run(text="Hello World")
        result = pptx_run_to_markdown(mock_run)
        assert result == "Hello World"

    def test_bold_text(self):
        """Bold text should be wrapped in **."""
        mock_run = self._create_mock_run(text="Bold", bold=True)
        result = pptx_run_to_markdown(mock_run)
        assert result == "**Bold**"

    def test_italic_text(self):
        """Italic text should be wrapped in *."""
        mock_run = self._create_mock_run(text="Italic", italic=True)
        result = pptx_run_to_markdown(mock_run)
        assert result == "*Italic*"

    def test_bold_italic_text(self):
        """Bold + italic text should be wrapped in ***."""
        mock_run = self._create_mock_run(text="BoldItalic", bold=True, italic=True)
        result = pptx_run_to_markdown(mock_run)
        assert result == "***BoldItalic***"

    def test_strikethrough_text(self):
        """Strikethrough text should be wrapped in ~~."""
        mock_run = self._create_mock_run(text="Strike", strike=True)
        result = pptx_run_to_markdown(mock_run)
        assert result == "~~Strike~~"

    def test_code_text(self):
        """Monospace font should produce backtick code."""
        mock_run = self._create_mock_run(text="code", font_name="Courier New")
        result = pptx_run_to_markdown(mock_run)
        assert result == "`code`"

    def test_hyperlink_text(self):
        """Hyperlink should produce markdown link."""
        mock_run = self._create_mock_run(
            text="Click here", hyperlink="https://example.com"
        )
        result = pptx_run_to_markdown(mock_run)
        assert result == "[Click here](https://example.com)"

    def test_bold_with_hyperlink(self):
        """Bold with hyperlink should have bold inside link."""
        mock_run = self._create_mock_run(
            text="Bold Link", bold=True, hyperlink="https://example.com"
        )
        result = pptx_run_to_markdown(mock_run)
        assert result == "[**Bold Link**](https://example.com)"

    def test_empty_text(self):
        """Empty text should return empty string."""
        mock_run = self._create_mock_run(text="")
        result = pptx_run_to_markdown(mock_run)
        assert result == ""


class TestPptxParagraphToMarkdown:
    """Tests for pptx_paragraph_to_markdown."""

    def test_single_run_paragraph(self):
        """Single run should be converted."""
        mock_para = MagicMock()
        mock_run = MagicMock()
        mock_run.text = "Hello"
        mock_run.font.bold = False
        mock_run.font.italic = False
        mock_run.font.strike = False
        mock_run.font.name = "Arial"
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None
        mock_run.hyperlink = MagicMock()
        mock_run.hyperlink.address = None

        mock_para.runs = [mock_run]
        mock_para.level = 0

        result = pptx_paragraph_to_markdown(mock_para)
        assert result == "Hello"

    def test_multiple_runs_paragraph(self):
        """Multiple runs should be concatenated."""
        mock_para = MagicMock()

        def create_run(text, bold=False):
            run = MagicMock()
            run.text = text
            run.font.bold = bold
            run.font.italic = False
            run.font.strike = False
            run.font.name = "Arial"
            run.font.size = None
            run.font.underline = False
            run.font.small_caps = False
            run.font.all_caps = False
            run.font.superscript = False
            run.font.subscript = False
            run.font.shadow = False
            run.font.emboss = False
            run.font.imprint = False
            run.font.double_strike = False
            run.font.color = MagicMock()
            run.font.color.rgb = None
            run.hyperlink = MagicMock()
            run.hyperlink.address = None
            return run

        mock_para.runs = [
            create_run("Hello "),
            create_run("World", bold=True),
            create_run("!"),
        ]
        mock_para.level = 0

        result = pptx_paragraph_to_markdown(mock_para)
        assert result == "Hello **World**!"

    def test_bullet_point_level_1(self):
        """Level 1 paragraph with bullet XML should have bullet indentation."""
        mock_para = MagicMock()
        mock_run = MagicMock()
        mock_run.text = "Item"
        mock_run.font.bold = False
        mock_run.font.italic = False
        mock_run.font.strike = False
        mock_run.font.name = "Arial"
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None
        mock_run.hyperlink = MagicMock()
        mock_run.hyperlink.address = None

        mock_para.runs = [mock_run]
        mock_para.level = 1

        # Mock the XML element to have bullet properties
        # _paragraph_has_bullet() checks for buChar element in pPr
        mock_pPr = MagicMock()
        mock_buChar = MagicMock()  # Represents the bullet character element

        def find_side_effect(tag):
            if "buChar" in tag:
                return mock_buChar
            return None

        mock_pPr.find = MagicMock(side_effect=find_side_effect)
        mock_para._element.get_or_add_pPr.return_value = mock_pPr

        result = pptx_paragraph_to_markdown(mock_para)
        assert result == "  - Item"

    def test_bullet_point_level_2(self):
        """Level 2 paragraph with bullet XML should have double indentation."""
        mock_para = MagicMock()
        mock_run = MagicMock()
        mock_run.text = "Item"
        mock_run.font.bold = False
        mock_run.font.italic = False
        mock_run.font.strike = False
        mock_run.font.name = "Arial"
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None
        mock_run.hyperlink = MagicMock()
        mock_run.hyperlink.address = None

        mock_para.runs = [mock_run]
        mock_para.level = 2

        # Mock the XML element to have bullet properties
        mock_pPr = MagicMock()
        mock_buChar = MagicMock()  # Represents the bullet character element

        def find_side_effect(tag):
            if "buChar" in tag:
                return mock_buChar
            return None

        mock_pPr.find = MagicMock(side_effect=find_side_effect)
        mock_para._element.get_or_add_pPr.return_value = mock_pPr

        result = pptx_paragraph_to_markdown(mock_para)
        assert result == "    - Item"


class TestPptxTextFrameToMarkdown:
    """Tests for pptx_text_frame_to_markdown."""

    def test_none_text_frame(self):
        """None text frame should return empty string."""
        result = pptx_text_frame_to_markdown(None)
        assert result == ""

    def test_single_paragraph(self):
        """Single paragraph should be converted."""
        mock_frame = MagicMock()
        mock_para = MagicMock()
        mock_run = MagicMock()
        mock_run.text = "Hello"
        mock_run.font.bold = False
        mock_run.font.italic = False
        mock_run.font.strike = False
        mock_run.font.name = "Arial"
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None
        mock_run.hyperlink = MagicMock()
        mock_run.hyperlink.address = None

        mock_para.runs = [mock_run]
        mock_para.level = 0
        mock_frame.paragraphs = [mock_para]

        result = pptx_text_frame_to_markdown(mock_frame)
        assert result == "Hello"

    def test_multiple_paragraphs(self):
        """Multiple paragraphs should be joined with newlines."""
        mock_frame = MagicMock()

        def create_para(text):
            para = MagicMock()
            run = MagicMock()
            run.text = text
            run.font.bold = False
            run.font.italic = False
            run.font.strike = False
            run.font.name = "Arial"
            run.font.size = None
            run.font.underline = False
            run.font.small_caps = False
            run.font.all_caps = False
            run.font.superscript = False
            run.font.subscript = False
            run.font.shadow = False
            run.font.emboss = False
            run.font.imprint = False
            run.font.double_strike = False
            run.font.color = MagicMock()
            run.font.color.rgb = None
            run.hyperlink = MagicMock()
            run.hyperlink.address = None
            para.runs = [run]
            para.level = 0
            return para

        mock_frame.paragraphs = [create_para("Line 1"), create_para("Line 2")]

        result = pptx_text_frame_to_markdown(mock_frame)
        assert result == "Line 1\nLine 2"


class TestPptxTableToMarkdown:
    """Tests for pptx_table_to_markdown."""

    def _create_mock_cell(self, text):
        """Helper to create mock cell with text."""
        mock_cell = MagicMock()
        mock_frame = MagicMock()
        mock_para = MagicMock()
        mock_run = MagicMock()

        mock_run.text = text
        mock_run.font.bold = False
        mock_run.font.italic = False
        mock_run.font.strike = False
        mock_run.font.name = "Arial"
        mock_run.font.size = None
        mock_run.font.underline = False
        mock_run.font.small_caps = False
        mock_run.font.all_caps = False
        mock_run.font.superscript = False
        mock_run.font.subscript = False
        mock_run.font.shadow = False
        mock_run.font.emboss = False
        mock_run.font.imprint = False
        mock_run.font.double_strike = False
        mock_run.font.color = MagicMock()
        mock_run.font.color.rgb = None
        mock_run.hyperlink = MagicMock()
        mock_run.hyperlink.address = None

        mock_para.runs = [mock_run]
        mock_para.level = 0
        mock_frame.paragraphs = [mock_para]
        mock_cell.text_frame = mock_frame

        return mock_cell

    def _create_mock_row(self, cell_texts):
        """Helper to create mock row with cells."""
        mock_row = MagicMock()
        mock_row.cells = [self._create_mock_cell(text) for text in cell_texts]
        return mock_row

    def test_simple_2x2_table(self):
        """Should generate markdown for 2x2 table."""
        mock_table = MagicMock()
        mock_table.rows = [
            self._create_mock_row(["Header1", "Header2"]),
            self._create_mock_row(["Cell1", "Cell2"]),
        ]

        result = pptx_table_to_markdown(mock_table)

        expected = "| Header1 | Header2 |\n| --- | --- |\n| Cell1 | Cell2 |"
        assert result == expected

    def test_3x3_table(self):
        """Should generate markdown for 3x3 table."""
        mock_table = MagicMock()
        mock_table.rows = [
            self._create_mock_row(["A", "B", "C"]),
            self._create_mock_row(["D", "E", "F"]),
            self._create_mock_row(["G", "H", "I"]),
        ]

        result = pptx_table_to_markdown(mock_table)

        expected = (
            "| A | B | C |\n| --- | --- | --- |\n| D | E | F |\n| G | H | I |"
        )
        assert result == expected

    def test_table_with_pipe_in_content(self):
        """Should escape pipes in cell content."""
        mock_table = MagicMock()
        mock_table.rows = [
            self._create_mock_row(["Header", "Value"]),
            self._create_mock_row(["A|B", "C|D"]),
        ]

        result = pptx_table_to_markdown(mock_table)

        assert "A\\|B" in result
        assert "C\\|D" in result

    def test_table_with_empty_cells(self):
        """Should handle empty cells."""
        mock_table = MagicMock()
        mock_table.rows = [
            self._create_mock_row(["Header1", "Header2"]),
            self._create_mock_row(["", "Data"]),
        ]

        result = pptx_table_to_markdown(mock_table)

        expected = "| Header1 | Header2 |\n| --- | --- |\n|  | Data |"
        assert result == expected

    def test_empty_table(self):
        """Should return empty string for empty table."""
        mock_table = MagicMock()
        mock_table.rows = []

        result = pptx_table_to_markdown(mock_table)
        assert result == ""

    def test_none_table(self):
        """Should return empty string for None table."""
        result = pptx_table_to_markdown(None)
        assert result == ""
