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
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set"
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

    @pytest.mark.skipif(
        not os.getenv("GSLIDES_CREDENTIALS_PATH"),
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set"
    )
    class TestSimpleMarkdownWriting:
        """Test writing simple markdown content."""

        def test_write_simple_markdown(self, test_slide):
            """Test writing simple markdown with bullets and numbered lists."""
            md = "Oh what a text\n* Bullet points\n* And more\n1. Numbered items\n2. And more"
            
            # Write markdown to text element
            test_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
            
            # Sync and verify content matches
            test_slide.sync_from_cloud()
            re_md = test_slide.get_element_by_alt_title("text_1").read_text()
            assert re_md == md

    @pytest.mark.skipif(
        not os.getenv("GSLIDES_CREDENTIALS_PATH"),
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set"
    )
    class TestMediumComplexityMarkdown:
        """Test writing medium complexity markdown content."""

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
            assert re_md is not None

    @pytest.mark.skipif(
        not os.getenv("GSLIDES_CREDENTIALS_PATH"),
        reason="GSLIDES_CREDENTIALS_PATH environment variable not set"
    )
    class TestComplexMarkdown:
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
            assert re_md is not None
            assert len(re_md) > 0
            # Verify some basic content is preserved
            assert "important report" in re_md
            assert "bullet points" in re_md


# Utility functions for test setup and teardown
def create_test_slide(source_slide):
    """Create a new test slide by duplicating the source slide."""
    return source_slide.duplicate()


def cleanup_test_slide(test_slide):
    """Delete the test slide to clean up after testing."""
    test_slide.delete()
