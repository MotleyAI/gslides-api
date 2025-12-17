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

    def test_name_match_wrong_element_count(self, library):
        """Should raise ValueError when slide name matches but element count differs."""
        # Title slide has 2 elements, provide only 1
        markdown = """<!-- slide: Title -->
<!-- text: Title -->
Just a title, no subtitle"""
        with pytest.raises(ValueError, match="Element names don't match"):
            library.slide_from_markdown(markdown)

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
