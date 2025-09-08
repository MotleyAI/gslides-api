"""
Integration tests for markdown functionality with Google Slides API.

These tests verify that markdown text can be written to and read from Google Slides
with proper formatting preservation. Tests are skipped if GSLIDES_CREDENTIALS_PATH
environment variable is not set.
"""

import logging
import os

import pytest

from gslides_api.presentation import Presentation, Slide
from gslides_api.element.shape import Shape, ShapeElement
from gslides_api.element.text_content import TextContent
from gslides_api.domain.domain import Dimension, Size, Transform, Unit
from gslides_api.domain.text import ShapeProperties, TextStyle, Type as ShapeType
from gslides_api.client import api_client, initialize_credentials

logger = logging.getLogger(__name__)


class TestMarkdownIntegration:
    """Integration tests for markdown functionality with Google Slides."""

    @pytest.fixture(scope="class", autouse=True)
    def setup_credentials(self):
        """Initialize Google Slides API credentials if available."""
        credential_location = os.getenv("GSLIDES_CREDENTIALS_PATH")
        if credential_location:
            initialize_credentials(credential_location)

    @pytest.fixture(scope="class")
    def presentation(self):
        """Create a test presentation and clean up after all tests."""
        # Create test presentation
        test_presentation = Presentation.create_blank("Markdown Integration Test")
        yield test_presentation

        # Cleanup: delete the presentation after all tests
        try:
            api_client.delete_file(test_presentation.presentationId)
        except Exception as e:
            print(f"Warning: Failed to delete test presentation: {e}")

    @pytest.fixture(scope="class")
    def source_slide(self, presentation):
        """Create a source slide with test elements for duplication."""
        # Create a new slide for testing
        slide = Slide.create_blank(presentation.presentationId)

        # Create text element with alt title "text_1"
        text_element = ShapeElement(
            objectId="temp_text_1",
            presentation_id=presentation.presentationId,
            slide_id=slide.objectId,
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=200, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=50.0, translateY=100.0, unit="EMU"
            ),
            shape=Shape(
                shapeProperties=ShapeProperties(),
                shapeType=ShapeType.TEXT_BOX,
                text=TextContent(textElements=[]),
            ),
        )
        text_element_id = text_element.create_copy(
            parent_id=slide.objectId, presentation_id=presentation.presentationId
        )
        # Set alt text for identification and add initial content
        text_element.objectId = text_element_id
        text_element.set_alt_text(title="text_1")
        text_element.write_text("* Sample bullet point\\n* Another bullet point", as_markdown=True)

        # Create title element with alt title "title_1"
        title_element = ShapeElement(
            objectId="temp_title_1",
            presentation_id=presentation.presentationId,
            slide_id=slide.objectId,
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=50.0, translateY=50.0, unit="EMU"
            ),
            shape=Shape(
                shapeProperties=ShapeProperties(),
                shapeType=ShapeType.TEXT_BOX,
                text=TextContent(textElements=[]),
            ),
        )
        title_element_id = title_element.create_copy(
            parent_id=slide.objectId, presentation_id=presentation.presentationId
        )
        # Set alt text for identification and add initial content
        title_element.objectId = title_element_id
        title_element.set_alt_text(title="title_1")
        title_element.write_text("Sample Title", as_markdown=True)

        # Refresh the slide to get the updated elements
        return Slide.from_ids(presentation.presentationId, slide.objectId)

    @pytest.fixture(scope="class")
    def source_slide_2(self, presentation):
        """Create a second source slide with test elements for duplication."""
        # Create a new slide for testing
        slide = Slide.create_blank(presentation.presentationId)

        # Create text element with alt title "text"
        text_element = ShapeElement(
            objectId="temp_text_2",
            presentation_id=presentation.presentationId,
            slide_id=slide.objectId,
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=300, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=50.0, translateY=100.0, unit="EMU"
            ),
            shape=Shape(
                shapeProperties=ShapeProperties(),
                shapeType=ShapeType.TEXT_BOX,
                text=TextContent(textElements=[]),
            ),
        )
        text_element_id = text_element.create_copy(
            parent_id=slide.objectId, presentation_id=presentation.presentationId
        )
        # Set alt text for identification and add initial content
        text_element.objectId = text_element_id
        text_element.set_alt_text(title="text")
        # Write content with header and body that should create multiple styles
        text_element.write_text(
            "# Sample Header\nThis is body text with different formatting", as_markdown=True
        )

        # Refresh the slide to get the updated elements
        return Slide.from_ids(presentation.presentationId, slide.objectId)

    @pytest.fixture
    def test_slide(self, source_slide):
        """Create a test slide and ensure cleanup after test."""
        new_slide = source_slide.duplicate()
        yield new_slide
        # Cleanup: delete the slide after the test
        new_slide.delete()

    @pytest.fixture
    def test_slide_2(self, source_slide_2):
        """Create a test slide and ensure cleanup after test."""
        new_slide = source_slide_2.duplicate()
        yield new_slide
        # Cleanup: delete the slide after the test
        new_slide.delete()

    @pytest.mark.skipif(
        not os.getenv("GSLIDES_CREDENTIALS_PATH"),
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set",
    )
    class TestTextDeletion:
        """Test text deletion functionality."""

        def test_delete_text_with_bullets(self, test_slide):
            """Test deleting text from element with bullet points."""
            # Delete text from text element
            test_slide.get_element_by_alt_title("text_1").delete_text()

            # Sync and verify text is empty
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == ""

        def test_delete_text_without_bullets(self, test_slide):
            """Test deleting text from title element (no bullets)."""
            # Delete text from title element
            test_slide.get_element_by_alt_title("title_1").delete_text()

            # Sync and verify text is empty
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("title_1").read_text()
            assert re_md == ""

        def test_write_simple_markdown(self, test_slide):
            """Test writing simple markdown with bullets and numbered lists."""
            md = "Oh what a text\n* Bullet points\n* And more\n1. Numbered items\n2. And more"

            # Write markdown to text element
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)

            # Sync and verify content matches
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_write_medium_markdown(self, test_slide):
            """Test writing markdown with formatting and nested bullets."""
            medium_md = """This is a ***very*** *important* report with **bold** text.

* It illustrates **bullet points**
    * With nested sub-points
    * And even more `code` blocks"""

            # Write markdown to text element
            test_slide.get_element_by_alt_title("text_1").write_text(medium_md, as_markdown=True)

            # Sync from cloud
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()

            assert re_md == medium_md

        """Test writing complex markdown content with various formatting."""

        @pytest.fixture
        def complex_markdown(self):
            """Complex markdown content for testing."""
            return """This is a ***very*** *important* report with **bold** text.

* It illustrates **bullet points**
    * With nested sub-points
    * And even more `code` blocks
        * Third level nesting
* And even `code` blocks
* Plus *italic* formatting
    * Nested italic *emphasis*
    * With **bold** nested items

Here's a [link to Google](https://google.com) for testing hyperlinks.

Some ~~strikethrough~~ text to test deletion formatting.

Ordered list example:
1. First numbered item
    1. Nested numbered sub-item
    2. Another nested item with **bold**
        1. Third level numbering
2. Second with `inline code`
    1. Nested under third
    2. Final nested item

Mixed content with [links](https://example.com) and ~~crossed out~~ text."""

        def test_write_complex_markdown(self, test_slide, complex_markdown):
            """Test writing complex markdown with all supported features."""
            # Write complex markdown to text element
            test_slide.get_element_by_alt_title("text_1").write_text(
                complex_markdown, as_markdown=True
            )

            # Sync from cloud
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()

            # Note: This test currently doesn't assert exact equality due to known formatting differences
            # The test verifies that the operation completes without error and content is written
            if re_md != complex_markdown:
                # Find where the strings start differing
                min_len = min(len(re_md), len(complex_markdown))
                diff_pos = 0
                for i in range(min_len):
                    if re_md[i] != complex_markdown[i]:
                        diff_pos = i
                        break
                else:
                    # Strings are identical up to the shorter length
                    diff_pos = min_len

                # Show context around the difference
                context_start = max(0, diff_pos - 50)
                context_end = min(len(re_md), diff_pos + 50)
                expected_context_end = min(len(complex_markdown), diff_pos + 50)

                print(f"\nStrings differ at position {diff_pos}")
                print(
                    f"Expected: ...{repr(complex_markdown[context_start:expected_context_end])}..."
                )
                print(f"Actual:   ...{repr(re_md[context_start:context_end])}...")
                print(f"Expected length: {len(complex_markdown)}, Actual length: {len(re_md)}")

            assert re_md == complex_markdown

    def test_header_style(self, test_slide):
        old_element = test_slide.get_element_by_alt_title("title_1")
        old_style = old_element.shape.text.textElements[1].textRun.style
        old_element.write_text("New Title", as_markdown=True)
        test_slide.sync_from_cloud()
        new_element = test_slide.get_element_by_alt_title("title_1")
        new_style = new_element.shape.text.textElements[1].textRun.style
        assert old_style == new_style

    def test_header_style_2(self, test_slide_2):
        old_element = test_slide_2.get_element_by_alt_title("text")
        old_style_1 = old_element.styles[0]
        old_style_2 = old_element.styles[1]
        old_element.write_text("# New Title\nNew text", as_markdown=True)
        test_slide_2.sync_from_cloud()
        new_element = test_slide_2.get_element_by_alt_title("text")
        new_style_1 = new_element.styles[0]
        new_style_2 = new_element.styles[1]
        print("Testing header style...")
        compare_styles(old_style_1, new_style_1)
        print("Testing text style...")
        compare_styles(old_style_2, new_style_2)

    def test_bullet_style(self, test_slide_2):
        old_element = test_slide_2.get_element_by_alt_title("text")
        text = """# This is a *very important* report.
* It *illustrates* **bullet points** 
* And even `code` blocks
* And some more bullet points

Not to mention other text"""
        old_element.write_text(text, as_markdown=True)
        test_slide_2.sync_from_cloud()
        new_element = test_slide_2.get_element_by_alt_title("text")
        new_text = new_element.shape.text
        bullet_count = 0
        for e in new_text.textElements:
            # Make sure bullet points have proper bullet structures
            if e.paragraphMarker is not None and e.paragraphMarker.bullet is not None:
                bullet_count += 1
                # Check that bullet has proper structure (glyph should be set)
                assert e.paragraphMarker.bullet.glyph is not None
                # Check that bulletStyle exists (even if foregroundColor is None in fresh presentations)
                assert e.paragraphMarker.bullet.bulletStyle is not None

        # Ensure we found some bullet points
        assert bullet_count > 0, "Should have found bullet points in the markdown text"

        print("Testing header style...")

    def test_line_after_bullets(self, test_slide_2):
        old_element = test_slide_2.get_element_by_alt_title("text")
        text = """This is a very important report.
* Here is a bullet point
* And another

And text outside of the list item."""
        old_element.write_text(text, as_markdown=True)
        test_slide_2.sync_from_cloud()
        new_element = test_slide_2.get_element_by_alt_title("text")
        new_text = new_element.read_text()
        assert new_text.strip() == text.strip()
        # TODO: fix to_markdown to insert extra newline after lists
        print("Testing line after list...")

    def test_line_after_numbered_list(self, test_slide_2):
        old_element = test_slide_2.get_element_by_alt_title("text")
        text = """This is a very important report.
1. Here is a numbered item
2. And another

And text outside of the list item."""
        old_element.write_text(text, as_markdown=True)
        test_slide_2.sync_from_cloud()
        new_element = test_slide_2.get_element_by_alt_title("text")
        new_text = new_element.read_text()
        assert new_text.strip() == text.strip()
        # TODO: fix to_markdown to insert extra newline after lists
        print("Testing line after list...")

    # his one reproduces a bug in GS API, namely no support for inserting newlines into list items
    def test_newline_in_list(self, test_slide_2):
        old_element = test_slide_2.get_element_by_alt_title("text")
        text = """# This is a very important report.
* Here is a bullet point
* And another
And some more text that belongs in the last list item """
        with pytest.raises(ValueError):
            old_element.write_text(text, as_markdown=True)

    @pytest.mark.skipif(
        not os.getenv("GSLIDES_CREDENTIALS_PATH"),
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set",
    )
    class TestIndividualFormattingTypes:
        """Test individual formatting types in standalone lines and bullet lists."""

        def test_strikethrough_standalone(self, test_slide):
            """Test strikethrough formatting in a standalone line."""
            md = "This is regular text with ~~strikethrough~~ formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_emphasis_standalone(self, test_slide):
            """Test emphasis formatting in a standalone line."""
            md = "This is regular text with *emphasis* formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_bold_standalone(self, test_slide):
            """Test bold formatting in a standalone line."""
            md = "This is regular text with **bold** formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_bold_emphasis_standalone(self, test_slide):
            """Test bold emphasis formatting in a standalone line."""
            md = "This is regular text with ***bold emphasis*** formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_code_standalone(self, test_slide):
            """Test code formatting in a standalone line."""
            md = "This is regular text with `code` formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_link_standalone(self, test_slide):
            """Test link formatting in a standalone line."""
            md = "This is regular text with a [link to Google](https://google.com) formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_strikethrough_in_bullet(self, test_slide):
            """Test strikethrough formatting within a bullet list."""
            md = "* This is regular text with ~~strikethrough~~ formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_emphasis_in_bullet(self, test_slide):
            """Test emphasis formatting within a bullet list."""
            md = "* This is regular text with *emphasis* formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_bold_in_bullet(self, test_slide):
            """Test bold formatting within a bullet list."""
            md = "* This is regular text with **bold** formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_bold_emphasis_in_bullet(self, test_slide):
            """Test bold emphasis formatting within a bullet list."""
            md = "* This is regular text with ***bold emphasis*** formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_code_in_bullet(self, test_slide):
            """Test code formatting within a bullet list."""
            md = "* This is regular text with `code` formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_link_in_bullet(self, test_slide):
            """Test link formatting within a bullet list."""
            md = "* This is regular text with a [link to Google](https://google.com) formatting."
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

        def test_simple_nested_numbered_list(self, test_slide):
            """Test simple nested numbered list reconstruction."""
            md = "1. First item\n    1. Nested item\n    2. Another nested item\n2. Second item"
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md


# Utility functions for test setup and teardown
def create_test_slide(source_slide):
    """Create a new test slide by duplicating the source slide."""
    return source_slide.duplicate()


def cleanup_test_slide(test_slide):
    """Delete the test slide to clean up after testing."""
    test_slide.delete()


def compare_styles(style1: TextStyle, style2: TextStyle):
    """Print the values of every attribute which is not equal between style1 and style2."""

    # Get all field names from the TextStyle model
    field_names = style1.model_fields.keys()

    differences_found = False
    message = []
    for field_name in field_names:
        value1 = getattr(style1, field_name, None)
        value2 = getattr(style2, field_name, None)

        if value1 != value2:
            if field_name == "weightedFontFamily":
                logger.warning(
                    f"{field_name}:\n  style1: {value1}\n  style2: {value2}, "
                    f"not raising because mangling font weights is 'normal' GSlides API behavior"
                )
            else:
                message.append(f"{field_name}:\n  style1: {value1}\n  style2: {value2}")

    if len(message) > 0:
        print("\n".join(message))
        raise AssertionError("\n".join(message))
