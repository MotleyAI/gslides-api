import logging
import pytest

from gslides_api.markdown.domain import (
    ContentType,
    MarkdownDeck, 
    MarkdownSlide,
    MarkdownSlideElement,
    example_md
)


class TestContentType:
    def test_content_type_enum_values(self):
        assert ContentType.TEXT == "text"
        assert ContentType.IMAGE == "image" 
        assert ContentType.CHART == "chart"
        assert ContentType.TABLE == "table"


class TestMarkdownSlideElement:
    def test_create_element(self):
        element = MarkdownSlideElement(
            name="Test",
            content="Some content",
            content_type=ContentType.TEXT
        )
        assert element.name == "Test"
        assert element.content == "Some content"
        assert element.content_type == ContentType.TEXT
        assert element.metadata == {}

    def test_element_with_metadata(self):
        element = MarkdownSlideElement(
            name="Test",
            content="Some content",
            content_type=ContentType.TEXT,
            metadata={"key": "value"}
        )
        assert element.metadata == {"key": "value"}

    def test_to_markdown_with_comment(self):
        element = MarkdownSlideElement(
            name="TestElement",
            content="## Header\n\nSome content",
            content_type=ContentType.TEXT
        )
        result = element.to_markdown()
        expected = "<!-- text: TestElement -->\n## Header\n\nSome content"
        assert result == expected

    def test_to_markdown_default_text_no_comment(self):
        element = MarkdownSlideElement(
            name="Default",
            content="# Title\n\nDefault content",
            content_type=ContentType.TEXT
        )
        result = element.to_markdown()
        expected = "# Title\n\nDefault content"
        assert result == expected

    def test_to_markdown_strips_trailing_whitespace(self):
        element = MarkdownSlideElement(
            name="Test",
            content="Content with trailing spaces   \n  ",
            content_type=ContentType.TEXT
        )
        result = element.to_markdown()
        expected = "<!-- text: Test -->\nContent with trailing spaces"
        assert result == expected


class TestMarkdownSlide:
    def test_create_empty_slide(self):
        slide = MarkdownSlide()
        assert slide.elements == []

    def test_slide_with_elements(self):
        elements = [
            MarkdownSlideElement(
                name="Default",
                content="# Title",
                content_type=ContentType.TEXT
            ),
            MarkdownSlideElement(
                name="Image1",
                content="![alt](url)",
                content_type=ContentType.IMAGE
            )
        ]
        slide = MarkdownSlide(elements=elements)
        assert len(slide.elements) == 2

    def test_to_markdown(self):
        elements = [
            MarkdownSlideElement(
                name="Default",
                content="# Slide Title",
                content_type=ContentType.TEXT
            ),
            MarkdownSlideElement(
                name="Description",
                content="Some description text",
                content_type=ContentType.TEXT
            )
        ]
        slide = MarkdownSlide(elements=elements)
        result = slide.to_markdown()
        expected = "# Slide Title\n\n<!-- text: Description -->\nSome description text"
        assert result == expected

    def test_from_markdown_simple(self):
        markdown = "# Title\n\nSome content"
        slide = MarkdownSlide.from_markdown(markdown)
        
        assert len(slide.elements) == 1
        assert slide.elements[0].name == "Default"
        assert slide.elements[0].content == "# Title\n\nSome content"
        assert slide.elements[0].content_type == ContentType.TEXT

    def test_from_markdown_with_comments(self):
        markdown = """# Title

<!-- text: Section1 -->
## Introduction

Content here

<!-- image: Img1 -->
![Image](url.jpg)"""

        slide = MarkdownSlide.from_markdown(markdown)
        
        assert len(slide.elements) == 3
        
        # Default text element
        assert slide.elements[0].name == "Default"
        assert slide.elements[0].content == "# Title"
        assert slide.elements[0].content_type == ContentType.TEXT
        
        # Section1 text element
        assert slide.elements[1].name == "Section1"
        assert slide.elements[1].content == "## Introduction\n\nContent here"
        assert slide.elements[1].content_type == ContentType.TEXT
        
        # Image element
        assert slide.elements[2].name == "Img1"
        assert slide.elements[2].content == "![Image](url.jpg)"
        assert slide.elements[2].content_type == ContentType.IMAGE

    def test_from_markdown_invalid_element_type_warn(self, caplog):
        markdown = """# Title

<!-- invalid: BadType -->
Some content"""

        with caplog.at_level(logging.WARNING):
            slide = MarkdownSlide.from_markdown(markdown, on_invalid_element="warn")
        
        assert len(slide.elements) == 2
        assert slide.elements[1].content_type == ContentType.TEXT
        assert "Invalid element type 'invalid'" in caplog.text

    def test_from_markdown_invalid_element_type_raise(self):
        markdown = """# Title

<!-- invalid: BadType -->
Some content"""

        with pytest.raises(ValueError, match="Invalid element type: invalid"):
            MarkdownSlide.from_markdown(markdown, on_invalid_element="raise")

    def test_from_markdown_empty_content_after_comment_ignored(self):
        markdown = """# Title

<!-- text: EmptySection -->


<!-- text: ValidSection -->
Valid content"""

        slide = MarkdownSlide.from_markdown(markdown)
        
        # Should have Default + ValidSection only (EmptySection ignored due to empty content)
        assert len(slide.elements) == 2
        assert slide.elements[0].name == "Default"
        assert slide.elements[1].name == "ValidSection"


class TestMarkdownDeck:
    def test_create_empty_deck(self):
        deck = MarkdownDeck()
        assert deck.slides == []

    def test_deck_with_slides(self):
        slides = [
            MarkdownSlide(elements=[
                MarkdownSlideElement(
                    name="Default",
                    content="# Slide 1",
                    content_type=ContentType.TEXT
                )
            ])
        ]
        deck = MarkdownDeck(slides=slides)
        assert len(deck.slides) == 1

    def test_dumps_single_slide(self):
        deck = MarkdownDeck(slides=[
            MarkdownSlide(elements=[
                MarkdownSlideElement(
                    name="Default",
                    content="# Title",
                    content_type=ContentType.TEXT
                )
            ])
        ])
        result = deck.dumps()
        expected = "---\n# Title\n"
        assert result == expected

    def test_dumps_multiple_slides(self):
        deck = MarkdownDeck(slides=[
            MarkdownSlide(elements=[
                MarkdownSlideElement(
                    name="Default",
                    content="# Slide 1",
                    content_type=ContentType.TEXT
                )
            ]),
            MarkdownSlide(elements=[
                MarkdownSlideElement(
                    name="Default",
                    content="# Slide 2",
                    content_type=ContentType.TEXT
                )
            ])
        ])
        result = deck.dumps()
        expected = "---\n# Slide 1\n\n---\n# Slide 2\n"
        assert result == expected

    def test_dumps_empty_slides_filtered_out(self):
        deck = MarkdownDeck(slides=[
            MarkdownSlide(elements=[]),  # Empty slide
            MarkdownSlide(elements=[
                MarkdownSlideElement(
                    name="Default",
                    content="# Valid Slide",
                    content_type=ContentType.TEXT
                )
            ])
        ])
        result = deck.dumps()
        expected = "---\n# Valid Slide\n"
        assert result == expected

    def test_loads_simple(self):
        markdown = "# Single Slide"
        deck = MarkdownDeck.loads(markdown)
        
        assert len(deck.slides) == 1
        assert len(deck.slides[0].elements) == 1
        assert deck.slides[0].elements[0].content == "# Single Slide"

    def test_loads_multiple_slides(self):
        markdown = """# Slide 1

---
# Slide 2

Some content"""
        
        deck = MarkdownDeck.loads(markdown)
        
        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2\n\nSome content"

    def test_loads_with_leading_separator(self):
        markdown = """---
# Slide 1

---
# Slide 2"""
        
        deck = MarkdownDeck.loads(markdown)
        
        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2"

    def test_loads_ignores_empty_slides(self):
        markdown = """# Slide 1

---

---
# Slide 2"""
        
        deck = MarkdownDeck.loads(markdown)
        
        assert len(deck.slides) == 2
        assert deck.slides[0].elements[0].content == "# Slide 1"
        assert deck.slides[1].elements[0].content == "# Slide 2"

    def test_loads_passes_invalid_element_option(self):
        markdown = """# Slide 1

<!-- invalid: BadType -->
Content"""
        
        with pytest.raises(ValueError):
            MarkdownDeck.loads(markdown, on_invalid_element="raise")


class TestFullCycleRoundTrip:
    """Test that loading and dumping preserves content exactly."""
    
    def test_example_md_round_trip(self):
        """Test that example_md can be loaded and dumped back to identical content."""
        deck = MarkdownDeck.loads(example_md)
        output = deck.dumps()
        
        # Strip whitespace for comparison (allow minor formatting differences)
        assert example_md.strip() == output.strip()

    def test_simple_content_round_trip(self):
        original = "---\n# Simple Slide\n\nJust some text\n"
        
        deck = MarkdownDeck.loads(original)
        output = deck.dumps()
        
        assert original == output

    def test_complex_content_round_trip(self):
        original = """---
# Slide 1

<!-- text: Intro -->
## Introduction

Welcome to the presentation

<!-- image: Logo -->
![Logo](https://example.com/logo.png)

---
# Slide 2

<!-- table: Data -->
| Name | Value |
|------|-------|
| A    | 1     |
| B    | 2     |
"""
        
        deck = MarkdownDeck.loads(original)
        output = deck.dumps()
        
        assert original == output

    def test_no_leading_separator_round_trip(self):
        original = """# First Slide

Some content

---
# Second Slide

More content
"""
        
        deck = MarkdownDeck.loads(original)
        output = deck.dumps()
        
        # Should add leading --- when dumping
        expected = "---\n" + original
        assert expected == output

    def test_whitespace_preservation(self):
        """Test that internal whitespace in content is preserved."""
        original = """---
# Title

<!-- text: Content -->
Line 1

Line 2 with   spaces

    Indented line
"""
        
        deck = MarkdownDeck.loads(original)
        output = deck.dumps()
        
        assert original == output