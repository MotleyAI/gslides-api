import os
import pytest
import re

from gslides_api.adapters.pptx_adapter import PowerPointAPIClient
from gslides_api.adapters.abstract_slides import AbstractPresentation


@pytest.fixture
def test_pptx_path():
    """Path to the test PPTX file."""
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(here, "Samplead Master Deck Template.pptx")


@pytest.fixture
def pptx_presentation(test_pptx_path):
    """Load the test presentation."""
    api_client = PowerPointAPIClient()
    presentation = AbstractPresentation.from_id(
        api_client=api_client, presentation_id=test_pptx_path
    )
    return presentation


class TestPPTXTextExtraction:
    """Test suite for PowerPoint text extraction without HTML/escaping."""

    def test_first_text_box_in_first_slide_no_escaping(self, pptx_presentation):
        """Test that the first text box in the first slide extracts text correctly.

        The text box contains 'QBR\\n{quarter}' with bold and color formatting.
        It should NOT:
        - Escape curly brackets (\\{quarter\\})
        - Add extra underscores (__QBR__)
        - Include HTML span tags for colors
        """
        # Get the first slide
        first_slide = pptx_presentation.slides[0]

        # Get text boxes (shapes with text frames)
        text_boxes = [
            elem
            for elem in first_slide.page_elements_flat
            if hasattr(elem, "has_text") and elem.has_text
        ]

        assert len(text_boxes) > 0, "First slide should have at least one text box"

        # Get the first NON-EMPTY text box (matches ingestion behavior)
        first_text_box = None
        for text_box in text_boxes:
            text_content = text_box.read_text(as_markdown=True)
            if text_content.strip():
                first_text_box = text_box
                break

        assert first_text_box is not None, "First slide should have at least one non-empty text box"

        # Extract text as markdown
        text_content = first_text_box.read_text(as_markdown=True)

        # Verify the text content
        assert text_content is not None, "Text content should not be None"

        # Main assertions: verify correct extraction
        assert (
            "{quarter}" in text_content
        ), "Template variable {quarter} should be present without escaping"
        assert "\\{" not in text_content, "Curly brackets should NOT be escaped"
        assert "\\}" not in text_content, "Curly brackets should NOT be escaped"
        assert "<span" not in text_content, "HTML span tags should NOT be in text content"
        assert "style=" not in text_content, "Inline CSS should NOT be in text content"

        # Check for QBR text (may have markdown bold formatting but no extra underscores with spaces)
        assert "QBR" in text_content, "Text 'QBR' should be present"
        # The text should not have the pptx2md pattern of " __text__ " (spaces + underscores)
        assert (
            " __QBR__ " not in text_content
        ), "Should not have pptx2md-style bold formatting with spaces"

        # The text should be clean: either "QBR" or "**QBR**" (markdown bold)
        # but not " __QBR__ " (pptx2md style)
        assert "QBR" in text_content.replace("**", ""), "QBR should be in the text"

    def test_first_text_box_styles_separate_from_content(self, pptx_presentation):
        """Test that style information is available separately from text content.

        Following the gslides-api pattern, styles should be extractable
        via the styles() method rather than embedded in text.
        """
        # Get the first slide
        first_slide = pptx_presentation.slides[0]

        # Get the first NON-EMPTY text box
        text_boxes = [
            elem
            for elem in first_slide.page_elements_flat
            if hasattr(elem, "has_text") and elem.has_text
        ]

        first_text_box = None
        for text_box in text_boxes:
            text_content = text_box.read_text(as_markdown=True)
            if text_content.strip():
                first_text_box = text_box
                break

        assert first_text_box is not None, "First slide should have at least one non-empty text box"

        # Get styles separately
        # Note: styles() method needs to be implemented in PowerPointShapeElement
        try:
            styles = first_text_box.styles()

            # Verify styles are returned
            assert styles is not None, "Styles should be extractable"

            # The styles should contain information about bold and color
            # (exact structure depends on implementation, but should be similar to gslides-api)

        except NotImplementedError:
            pytest.skip("styles() method not yet implemented for PowerPoint elements")

    def test_template_variable_recognition(self, pptx_presentation):
        """Test that template variables in curly brackets are preserved for parsing."""
        first_slide = pptx_presentation.slides[0]

        text_boxes = [
            elem
            for elem in first_slide.page_elements_flat
            if hasattr(elem, "has_text") and elem.has_text
        ]

        # Check first NON-EMPTY text box for template variables
        first_text_box = None
        for text_box in text_boxes:
            text_content = text_box.read_text(as_markdown=True)
            if text_content.strip():
                first_text_box = text_box
                break

        assert first_text_box is not None, "First slide should have at least one non-empty text box"

        text_content = first_text_box.read_text(as_markdown=True)

        # Verify template variables are unescaped and recognizable
        template_vars = re.findall(r"\{([^}]+)\}", text_content)

        assert len(template_vars) > 0, "Should find at least one template variable"
        assert "quarter" in template_vars, "Should recognize {quarter} as a template variable"

    def test_paragraph_newline_preservation(self, pptx_presentation):
        """Test that newlines between paragraphs are preserved."""
        first_slide = pptx_presentation.slides[0]

        text_boxes = [
            elem
            for elem in first_slide.page_elements_flat
            if hasattr(elem, "has_text") and elem.has_text
        ]

        # Find first NON-EMPTY text box
        first_text_box = None
        for text_box in text_boxes:
            text_content = text_box.read_text(as_markdown=True)
            if text_content.strip():
                first_text_box = text_box
                break

        assert first_text_box is not None, "First slide should have at least one non-empty text box"

        text_content = first_text_box.read_text(as_markdown=True)

        # The first text box should contain a newline between "QBR" and "{quarter}"
        assert "\n" in text_content, "Newlines should be preserved in text extraction"
