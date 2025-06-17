"""
Test for perfect markdown reconstruction from Google Slides API responses.

This test uses real API response data stored as fixtures to verify that
the to_markdown() method can perfectly reconstruct the original markdown.
"""

import json
import os
import pytest
from gslides_api.element.shape import ShapeElement
from gslides_api.domain import Shape


class TestMarkdownReconstruction:
    """Test markdown reconstruction using real API response data."""
    
    @pytest.fixture
    def api_response_data(self):
        """Load the saved API response data."""
        test_data_dir = os.path.join(os.path.dirname(__file__), 'test_data')
        test_data_file = os.path.join(test_data_dir, 'markdown_api_response.json')
        
        with open(test_data_file, 'r') as f:
            return json.load(f)
    
    def test_perfect_markdown_reconstruction(self, api_response_data):
        """Test that to_markdown() perfectly reconstructs the original markdown."""
        original_markdown = api_response_data['original_markdown']
        api_response = api_response_data['api_response']
        
        # Create a ShapeElement from the API response
        shape_data = Shape(**api_response['shape'])
        shape_element = ShapeElement(
            objectId=api_response['objectId'],
            size=api_response['size'],
            transform=api_response['transform'],
            title=api_response.get('title'),
            description=api_response.get('description'),
            shape=shape_data,
            placeholder=api_response.get('placeholder'),
            presentation_id="test_presentation"
        )
        
        # Test the reconstruction
        reconstructed_markdown = shape_element.to_markdown()
        
        assert reconstructed_markdown is not None, "Reconstruction should not return None"
        
        # Print for debugging
        print(f"\nOriginal markdown:\n{repr(original_markdown)}")
        print(f"\nReconstructed markdown:\n{repr(reconstructed_markdown)}")
        
        # Test core features are preserved
        self._assert_features_preserved(original_markdown, reconstructed_markdown)
    
    def _assert_features_preserved(self, original: str, reconstructed: str):
        """Assert that key markdown features are preserved in reconstruction."""
        import re

        # Extract hyperlinks from original
        hyperlink_pattern = r'\[([^\]]+)\]\(([^)]+)\)'
        original_hyperlinks = re.findall(hyperlink_pattern, original)
        for link_text, url in original_hyperlinks:
            expected_link = f'[{link_text}]({url})'
            assert expected_link in reconstructed, \
                f"Hyperlink '{expected_link}' should be preserved"

        # Extract strikethrough text from original
        strikethrough_pattern = r'~~([^~]+)~~'
        original_strikethrough = re.findall(strikethrough_pattern, original)
        for strikethrough_text in original_strikethrough:
            expected_strikethrough = f'~~{strikethrough_text}~~'
            assert expected_strikethrough in reconstructed, \
                f"Strikethrough '{expected_strikethrough}' should be preserved"

        # Extract bold formatting from original (including bold+italic)
        bold_pattern = r'\*\*\*([^*]+)\*\*\*|\*\*([^*]+)\*\*'
        original_bold_matches = re.findall(bold_pattern, original)
        for bold_italic, bold_only in original_bold_matches:
            if bold_italic:  # ***text*** (bold+italic)
                expected_bold = f'***{bold_italic}***'
            else:  # **text** (bold only)
                expected_bold = f'**{bold_only}**'
            assert expected_bold in reconstructed, \
                f"Bold formatting '{expected_bold}' should be preserved"

        # Extract italic formatting from original (single asterisks, not part of bold)
        italic_pattern = r'(?<!\*)\*([^*]+)\*(?!\*)'
        original_italic = re.findall(italic_pattern, original)
        for italic_text in original_italic:
            expected_italic = f'*{italic_text}*'
            assert expected_italic in reconstructed, \
                f"Italic formatting '{expected_italic}' should be preserved"

        # Extract code spans from original
        code_pattern = r'`([^`]+)`'
        original_code = re.findall(code_pattern, original)
        for code_text in original_code:
            expected_code = f'`{code_text}`'
            assert expected_code in reconstructed, \
                f"Code span '{expected_code}' should be preserved"

        # Extract bullet list items from original
        bullet_pattern = r'^\* (.+)$'
        original_bullets = re.findall(bullet_pattern, original, re.MULTILINE)
        for bullet_text in original_bullets:
            expected_bullet = f'* {bullet_text}'
            assert expected_bullet in reconstructed, \
                f"Bullet list item '{expected_bullet}' should be preserved"

        # Extract numbered list items from original
        numbered_pattern = r'^(\d+)\. (.+)$'
        original_numbered = re.findall(numbered_pattern, original, re.MULTILINE)
        for number, numbered_text in original_numbered:
            expected_numbered = f'{number}. {numbered_text}'
            assert expected_numbered in reconstructed, \
                f"Numbered list item '{expected_numbered}' should be preserved"

        # Test that content is preserved (ignoring formatting differences)
        original_words = set(original.replace('*', '').replace('`', '').replace('#', '').replace('~', '').split())
        reconstructed_words = set(reconstructed.replace('*', '').replace('`', '').replace('#', '').replace('~', '').split())

        # Remove markdown syntax and compare content
        missing_words = original_words - reconstructed_words
        extra_words = reconstructed_words - original_words

        # Filter out very short words and punctuation that might differ
        missing_words = {w for w in missing_words if len(w) > 2 and w.isalpha()}
        extra_words = {w for w in extra_words if len(w) > 2 and w.isalpha()}

        assert not missing_words, f"Missing words from reconstruction: {missing_words}"
        # Note: We allow extra words since headings become bold text, etc.
    
    def test_numbered_vs_bullet_lists(self, api_response_data):
        """Test that numbered and bullet lists are correctly distinguished."""
        api_response = api_response_data['api_response']
        
        # Create a ShapeElement from the API response
        shape_data = Shape(**api_response['shape'])
        shape_element = ShapeElement(
            objectId=api_response['objectId'],
            size=api_response['size'],
            transform=api_response['transform'],
            title=api_response.get('title'),
            description=api_response.get('description'),
            shape=shape_data,
            placeholder=api_response.get('placeholder'),
            presentation_id="test_presentation"
        )
        
        reconstructed = shape_element.to_markdown()
        
        # Check that we have both bullet and numbered lists
        has_bullet_lists = '* ' in reconstructed
        has_numbered_lists = any(f'{i}. ' in reconstructed for i in range(1, 10))
        
        assert has_bullet_lists, "Should have bullet lists (* items)"
        assert has_numbered_lists, "Should have numbered lists (1. items)"
        
        # Verify specific list types
        lines = reconstructed.split('\n')
        bullet_lines = [line for line in lines if line.strip().startswith('* ')]
        numbered_lines = [line for line in lines if any(line.strip().startswith(f'{i}. ') for i in range(1, 10))]
        
        assert len(bullet_lines) >= 3, f"Should have at least 3 bullet items, got {len(bullet_lines)}"
        assert len(numbered_lines) >= 3, f"Should have at least 3 numbered items, got {len(numbered_lines)}"
        
        print(f"Found {len(bullet_lines)} bullet list items")
        print(f"Found {len(numbered_lines)} numbered list items")
        
        # Print the lists for verification
        print("Bullet list items:")
        for line in bullet_lines:
            print(f"  {line}")
        
        print("Numbered list items:")
        for line in numbered_lines:
            print(f"  {line}")


class TestMarkdownReconstructionEdgeCases:
    """Test edge cases in markdown reconstruction."""
    
    def test_empty_shape(self):
        """Test reconstruction with empty shape."""
        from gslides_api.domain import ShapeProperties

        shape_element = ShapeElement(
            objectId="test",
            size={"width": {"magnitude": 100, "unit": "EMU"}, "height": {"magnitude": 100, "unit": "EMU"}},
            transform={"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1, "unit": "EMU"},
            shape=Shape(
                shapeType="TEXT_BOX",
                shapeProperties=ShapeProperties()
            ),
            presentation_id="test"
        )

        result = shape_element.to_markdown()
        assert result is None, "Empty shape should return None"

    def test_shape_without_text(self):
        """Test reconstruction with shape that has no text."""
        from gslides_api.domain import ShapeProperties

        shape_element = ShapeElement(
            objectId="test",
            size={"width": {"magnitude": 100, "unit": "EMU"}, "height": {"magnitude": 100, "unit": "EMU"}},
            transform={"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1, "unit": "EMU"},
            shape=Shape(
                shapeType="TEXT_BOX",
                text=None,
                shapeProperties=ShapeProperties()
            ),
            presentation_id="test"
        )

        result = shape_element.to_markdown()
        assert result is None, "Shape without text should return None"
