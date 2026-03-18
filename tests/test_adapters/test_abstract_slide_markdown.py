"""Tests for AbstractSlide.markdown() method."""

from unittest.mock import MagicMock, PropertyMock

import pytest

from gslides_api.agnostic.element import MarkdownTableElement, TableData

from gslides_api.adapters.abstract_slides import (
    AbstractAltText,
    AbstractElement,
    AbstractImageElement,
    AbstractShapeElement,
    AbstractSlide,
    AbstractTableElement,
    _extract_font_size_from_table,
    _extract_font_size_pt,
)


def _make_shape_element(
    object_id="shape1",
    title=None,
    text="Hello World",
    x=0.5,
    y=0.3,
    w=9.0,
    h=1.2,
    font_size_pt=18.0,
):
    """Create a mock AbstractShapeElement."""
    elem = MagicMock(spec=AbstractShapeElement)
    elem.objectId = object_id
    elem.alt_text = AbstractAltText(title=title)
    elem.type = "SHAPE"
    type(elem).has_text = PropertyMock(return_value=True)
    elem.read_text.return_value = text
    elem.absolute_position.return_value = (x, y)
    elem.absolute_size.return_value = (w, h)

    # Create a mock style with font_size_pt attribute
    mock_style = MagicMock()
    mock_style.font_size_pt = font_size_pt
    elem.styles.return_value = [mock_style]

    return elem


def _make_image_element(
    object_id="img1",
    title="Chart",
    x=1.0,
    y=3.0,
    w=8.0,
    h=4.0,
):
    """Create a mock AbstractImageElement."""
    elem = MagicMock(spec=AbstractImageElement)
    elem.objectId = object_id
    elem.alt_text = AbstractAltText(title=title)
    elem.type = "IMAGE"
    elem.absolute_position.return_value = (x, y)
    elem.absolute_size.return_value = (w, h)
    return elem


def _make_table_element(
    object_id="table1",
    title="Data",
    x=0.5,
    y=7.5,
    w=9.0,
    h=2.0,
    headers=None,
    rows=None,
):
    """Create a mock AbstractTableElement."""
    if headers is None:
        headers = ["Metric", "Q3", "Q4"]
    if rows is None:
        rows = [["Revenue", "$1.2M", "$1.5M"], ["Growth", "8%", "12%"]]

    table_data = TableData(headers=headers, rows=rows)
    md_elem = MarkdownTableElement(name=title, content=table_data)

    elem = MagicMock(spec=AbstractTableElement)
    elem.objectId = object_id
    elem.alt_text = AbstractAltText(title=title)
    elem.type = "TABLE"
    elem.absolute_position.return_value = (x, y)
    elem.absolute_size.return_value = (w, h)
    elem.to_markdown_element.return_value = md_elem
    return elem


def _make_slide(elements):
    """Create a mock AbstractSlide with given elements."""
    slide = MagicMock(spec=AbstractSlide)
    slide.page_elements_flat = elements
    # Use the real markdown() method
    slide.markdown = AbstractSlide.markdown.__get__(slide, type(slide))
    return slide


class TestExtractFontSizePt:
    def test_none_styles(self):
        assert _extract_font_size_pt(None) == 12.0

    def test_empty_styles(self):
        assert _extract_font_size_pt([]) == 12.0

    def test_gslides_style(self):
        style = MagicMock()
        style.font_size_pt = 24.0
        assert _extract_font_size_pt([style]) == 24.0

    def test_pptx_style(self):
        fs = MagicMock()
        fs.pt = 18.0
        style = {"font_size": fs}
        assert _extract_font_size_pt([style]) == 18.0

    def test_multiple_styles_returns_max(self):
        s1 = MagicMock()
        s1.font_size_pt = 14.0
        s2 = MagicMock()
        s2.font_size_pt = 24.0
        assert _extract_font_size_pt([s1, s2]) == 24.0

    def test_gslides_style_none_font_size(self):
        style = MagicMock()
        style.font_size_pt = None
        assert _extract_font_size_pt([style]) == 12.0


class TestExtractFontSizeFromTable:
    def test_fallback_no_adapter_attributes(self):
        elem = MagicMock(spec=AbstractTableElement)
        elem.pptx_element = None
        elem.gslides_element = None
        assert _extract_font_size_from_table(elem) == 10.0

    def test_no_attributes_at_all(self):
        elem = MagicMock(spec=[])
        result = _extract_font_size_from_table(elem)
        assert result == 10.0


class TestSlideMarkdown:
    def test_text_element(self):
        shape = _make_shape_element(
            title="Title",
            text="# Quarterly Report",
            x=0.5,
            y=0.3,
            w=9.0,
            h=1.2,
            font_size_pt=18.0,
        )
        slide = _make_slide([shape])
        md = slide.markdown()

        assert "<!-- text: Title |" in md
        assert "pos=(0.5,0.3)" in md
        assert "size=(9.0,1.2)" in md
        assert "chars -->" in md
        assert "# Quarterly Report" in md

    def test_image_element(self):
        img = _make_image_element(
            title="Chart",
            x=1.0,
            y=3.0,
            w=8.0,
            h=4.0,
        )
        slide = _make_slide([img])
        md = slide.markdown()

        assert "<!-- image: Chart |" in md
        assert "pos=(1.0,3.0)" in md
        assert "size=(8.0,4.0)" in md

    def test_table_element(self):
        table = _make_table_element(
            title="Data",
            x=0.5,
            y=7.5,
            w=9.0,
            h=2.0,
        )
        slide = _make_slide([table])
        md = slide.markdown()

        assert "<!-- table: Data |" in md
        assert "pos=(0.5,7.5)" in md
        assert "size=(9.0,2.0)" in md
        assert "chars/col -->" in md
        assert "Metric" in md
        assert "Revenue" in md

    def test_mixed_elements(self):
        shape = _make_shape_element(
            title="Title",
            text="Hello",
            x=0.5,
            y=0.3,
            w=9.0,
            h=1.0,
        )
        img = _make_image_element(
            title="Chart",
            x=1.0,
            y=2.0,
            w=8.0,
            h=3.0,
        )
        table = _make_table_element(
            title="Data",
            x=0.5,
            y=6.0,
            w=9.0,
            h=2.0,
        )
        slide = _make_slide([shape, img, table])
        md = slide.markdown()

        # Check all three elements are present, separated by double newlines
        parts = md.split("\n\n")
        assert len(parts) == 3
        assert "text: Title" in parts[0]
        assert "image: Chart" in parts[1]
        assert "table: Data" in parts[2]

    def test_element_uses_object_id_when_no_title(self):
        shape = _make_shape_element(object_id="abc123", title=None)
        slide = _make_slide([shape])
        md = slide.markdown()

        assert "text: abc123 |" in md

    def test_empty_slide(self):
        slide = _make_slide([])
        md = slide.markdown()
        assert md == ""

    def test_unknown_element_type(self):
        elem = MagicMock(spec=AbstractElement)
        elem.objectId = "group1"
        elem.alt_text = AbstractAltText(title="MyGroup")
        elem.type = "GROUP"
        elem.absolute_position.return_value = (0.0, 0.0)
        elem.absolute_size.return_value = (10.0, 7.5)

        slide = _make_slide([elem])
        md = slide.markdown()

        assert "<!-- GROUP: MyGroup |" in md
        assert "pos=(0.0,0.0)" in md

    def test_char_capacity_calculation(self):
        """Verify that char capacity is included and reasonable."""
        shape = _make_shape_element(
            title="Body",
            text="Some text",
            w=9.0,
            h=5.0,
            font_size_pt=12.0,
        )
        slide = _make_slide([shape])
        md = slide.markdown()

        # Extract the char count from the markdown
        import re
        match = re.search(r"~(\d+) chars", md)
        assert match is not None
        chars = int(match.group(1))
        # For a 9x5 inch box with 12pt font, we expect a reasonable number
        assert chars > 100
        assert chars < 10000

    def test_table_col_chars_calculation(self):
        """Verify per-column char capacity for tables."""
        table = _make_table_element(
            title="BigTable",
            w=9.0,
            h=3.0,
            headers=["A", "B", "C"],
            rows=[["1", "2", "3"]],
        )
        slide = _make_slide([table])
        md = slide.markdown()

        import re
        match = re.search(r"~(\d+) chars/col", md)
        assert match is not None
        chars_per_col = int(match.group(1))
        # 9 inches / 3 cols = 3 inches per col, with 10pt font
        assert chars_per_col > 10
        assert chars_per_col < 1000

    def test_shape_without_text(self):
        """Shape element without text should not appear as text type."""
        elem = MagicMock(spec=AbstractShapeElement)
        elem.objectId = "empty_shape"
        elem.alt_text = AbstractAltText(title="EmptyBox")
        elem.type = "SHAPE"
        type(elem).has_text = PropertyMock(return_value=False)
        elem.absolute_position.return_value = (1.0, 1.0)
        elem.absolute_size.return_value = (3.0, 2.0)

        slide = _make_slide([elem])
        md = slide.markdown()

        # Should fall through to the generic case since has_text is False
        assert "<!-- SHAPE: EmptyBox |" in md
        assert "text:" not in md
