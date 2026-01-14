"""Tests for SlideLayoutLibrary.slide_from_markdown functionality."""

import pytest

from gslides_api.agnostic.element import ContentType
from gslides_api.agnostic.presentation import MarkdownSlide

from gslides_api.agnostic.library import SlideLayoutLibrary, example_slides


@pytest.fixture
def library():
    """Create a SlideLayoutLibrary with example slides."""
    return SlideLayoutLibrary(slides=example_slides)


class TestSlideFromMarkdown:
    """Tests for slide_from_markdown method."""

    @pytest.mark.parametrize("slide", example_slides, ids=lambda s: s.name)
    def test_match_all_slide_types_by_name(self, library, slide):
        """Parse each slide's to_markdown() output and verify it matches back."""
        markdown = slide.to_markdown()
        result = library.slide_from_markdown(markdown)
        assert result.name == slide.name
        assert len(result.elements) == len(slide.elements)
        for result_el, expected_el in zip(result.elements, slide.elements):
            assert result_el.name == expected_el.name

    def test_name_match_wrong_element_names(self, library):
        """Should raise ValueError when slide name matches but element names don't."""
        # Create markdown with correct slide name but wrong element names
        markdown = """<!-- slide: Title -->
<!-- text: WrongName1 -->
Some title text

<!-- text: WrongName2 -->
Some subtitle text"""
        with pytest.raises(ValueError, match="Element names don't match"):
            library.slide_from_markdown(markdown)

    def test_subset_matching_with_missing_elements(self, library):
        """Should match when parsed elements are a subset of library elements."""
        # Title slide has 2 elements (Title, Subtitle), provide only 1
        markdown = """<!-- slide: Title -->
<!-- text: Title -->
Just a title, no subtitle"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Title"
        assert len(result.elements) == 2  # Both elements should be present
        assert result.elements[0].name == "Title"
        assert result.elements[0].content == "Just a title, no subtitle"
        assert result.elements[1].name == "Subtitle"
        assert result.elements[1].content is None  # None content for missing element

    def test_name_match_wrong_element_types(self, library):
        """Should raise ValueError when element names match but types don't."""
        # Chart and text slide expects chart + text, provide table + text
        markdown = """<!-- slide: Chart and text slide -->
<!-- text: Title -->
Chart Slide Title

<!-- table: Chart -->
| A | B |
|---|---|
| 1 | 2 |

<!-- text: Text -->
Some text content"""
        with pytest.raises(ValueError, match="Element types don't match"):
            library.slide_from_markdown(markdown)

    def test_match_by_element_names_no_slide_name(self, library):
        """Should match by element names when no slide name comment is present."""
        # Title slide elements without slide name comment
        markdown = """<!-- text: Title -->
Main Title Here

<!-- text: Subtitle -->
A subtitle"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Title"
        assert len(result.elements) == 2
        assert result.elements[0].name == "Title"
        assert result.elements[1].name == "Subtitle"

    @pytest.mark.parametrize(
        "content_type,content_markdown",
        [
            ("text", "<!-- text: Content -->\nSome text content"),
            (
                "table",
                """<!-- table: Content -->
| Header1 | Header2 |
|---------|---------|
| Cell1   | Cell2   |""",
            ),
            (
                "chart",
                """<!-- chart: Content -->
```json
{"type": "bar", "data": [1, 2, 3]}
```""",
            ),
        ],
        ids=["text", "table", "chart"],
    )
    def test_content_any_matches_specific_types(self, library, content_type, content_markdown):
        """ContentElement (ANY) in library should match specific content types."""
        # "Header and single content" has Title + Content (ANY)
        markdown = f"""<!-- slide: Header and single content -->
<!-- text: Title -->
A Title

{content_markdown}"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Header and single content"
        assert len(result.elements) == 2
        assert result.elements[1].name == "Content"

    def test_no_match_raises(self, library):
        """Should raise ValueError when no matching layout is found."""
        # Create markdown with element count/types that don't match any library slide
        # Using 4 text elements - no library slide has this pattern
        markdown = """<!-- text: UniqueElement1 -->
Content 1

<!-- text: UniqueElement2 -->
Content 2

<!-- text: UniqueElement3 -->
Content 3

<!-- text: UniqueElement4 -->
Content 4"""
        with pytest.raises(ValueError, match="No matching slide layout found"):
            library.slide_from_markdown(markdown)

    def test_match_section_header_single_element(self, library):
        """Should match section header slide with single element."""
        markdown = """<!-- slide: Section header -->
<!-- text: Section header -->
## My Section"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Section header"
        assert len(result.elements) == 1
        assert result.elements[0].name == "Section header"

    def test_match_comparison_slide_multiple_any_elements(self, library):
        """Should match comparison slide with multiple ANY content elements."""
        markdown = """<!-- slide: Comparison -->
<!-- text: Title -->
Comparison Title

<!-- text: Subtitle 1 -->
First Option

<!-- text: Content 1 -->
Description of first option

<!-- text: Subtitle 2 -->
Second Option

<!-- text: Content 2 -->
Description of second option"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Comparison"
        assert len(result.elements) == 5

    def test_subset_matching_preserves_library_order(self, library):
        """Elements should be ordered per library template, not parsed order."""
        # Provide Subtitle before Title (wrong order) - should still work
        markdown = """<!-- slide: Title -->
<!-- text: Subtitle -->
The subtitle first

<!-- text: Title -->
The title second"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Title"
        assert len(result.elements) == 2
        # Elements should be in library order (Title, Subtitle)
        assert result.elements[0].name == "Title"
        assert result.elements[0].content == "The title second"
        assert result.elements[1].name == "Subtitle"
        assert result.elements[1].content == "The subtitle first"

    def test_subset_matching_empty_parsed_slide(self, library):
        """Empty parsed slide should fill all elements with None content."""
        markdown = """<!-- slide: Title -->"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Title"
        assert len(result.elements) == 2
        assert result.elements[0].name == "Title"
        assert result.elements[0].content is None
        assert result.elements[1].name == "Subtitle"
        assert result.elements[1].content is None

    def test_subset_matching_all_elements_provided(self, library):
        """Full match should still work with all elements provided."""
        markdown = """<!-- slide: Title -->
<!-- text: Title -->
Main Title

<!-- text: Subtitle -->
The Subtitle"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Title"
        assert len(result.elements) == 2
        assert result.elements[0].content == "Main Title"
        assert result.elements[1].content == "The Subtitle"

    def test_subset_matching_fills_different_element_types(self, library):
        """Missing elements should be filled with correct type from template."""
        # Chart and text slide has: Title (text), Chart (chart), Text (text)
        # Only provide Title
        markdown = """<!-- slide: Chart and text slide -->
<!-- text: Title -->
Chart Slide Title"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Chart and text slide"
        assert len(result.elements) == 3
        assert result.elements[0].name == "Title"
        assert result.elements[0].content == "Chart Slide Title"
        assert result.elements[0].content_type == ContentType.TEXT
        assert result.elements[1].name == "Chart"
        assert result.elements[1].content is None
        assert result.elements[1].content_type == ContentType.CHART
        assert result.elements[2].name == "Text"
        assert result.elements[2].content is None
        assert result.elements[2].content_type == ContentType.TEXT

    def test_subset_matching_fills_table_with_none(self, library):
        """Missing table elements should be filled with None content."""
        # Header and table has: Title (text), Table (table)
        markdown = """<!-- slide: Header and table -->
<!-- text: Title -->
Table Slide Title"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Header and table"
        assert len(result.elements) == 2
        assert result.elements[1].name == "Table"
        assert result.elements[1].content_type == ContentType.TABLE
        # Missing table should have None content
        assert result.elements[1].content is None

    def test_subset_matching_fills_any_with_none(self, library):
        """Missing ANY elements should be filled with None content."""
        # Header and single content has: Title (text), Content (any)
        markdown = """<!-- slide: Header and single content -->
<!-- text: Title -->
Content Slide Title"""
        result = library.slide_from_markdown(markdown)
        assert result.name == "Header and single content"
        assert len(result.elements) == 2
        assert result.elements[1].name == "Content"
        assert result.elements[1].content_type == ContentType.ANY
        assert result.elements[1].content is None
