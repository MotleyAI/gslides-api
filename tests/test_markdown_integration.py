"""
Integration tests for markdown functionality with Google Slides API.

These tests verify that markdown text can be written to and read from Google Slides
with proper formatting preservation. Tests are skipped if GSLIDES_CREDENTIALS_PATH
environment variable is not set.
"""

import os
import pytest
from gslides_api import Presentation, Slide, initialize_credentials


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
        """Get the test presentation."""
        presentation_id = "1FHbC3ZXsEDUUNtQbxyyDQ3EFjwwt13_WovJAiYxhmOU"
        return Presentation.from_id(presentation_id)

    @pytest.fixture(scope="class")
    def source_slide(self, presentation):
        """Get the source slide for duplication."""
        return presentation.slides[1]

    @pytest.fixture
    def test_slide(self, source_slide):
        """Create a test slide and ensure cleanup after test."""
        new_slide = source_slide.duplicate()
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

            # Note: This test currently doesn't assert equality due to known formatting differences
            # The test verifies that the operation completes without error
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
        old_element.write_text("New Title", as_markdown=True)
        test_slide.sync_from_cloud()
        new_element = test_slide.get_element_by_alt_title("title_1")
        new_style = new_element.text_element.text_run.text_style
        assert old_style == new_style

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
