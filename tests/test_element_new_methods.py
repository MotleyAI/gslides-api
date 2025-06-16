"""Tests for new methods added to element classes."""

import json
import os
import pytest
from unittest.mock import patch, MagicMock
from pydantic import TypeAdapter

from gslides_api.element import PageElement, ImageElement
from gslides_api.shape_element import ShapeElement
from gslides_api.domain import (
    Size,
    Transform,
    Shape,
    Image,
    ShapeProperties,
    ShapeType,
    ImageReplaceMethod,
)
from gslides_api.presentation import Presentation


@pytest.fixture
def presentation_data():
    """Load presentation data from presentation_output.json."""
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "presentation_output.json"), "r") as f:
        return json.load(f)


@pytest.fixture
def presentation_object(presentation_data):
    """Create a Presentation object from the test data."""
    return Presentation.model_validate(presentation_data)


@pytest.fixture
def slide_3_shape_element(presentation_object):
    """Get the first shape element from slide at index 3."""
    slide_3 = presentation_object.slides[3]
    # Get the first page element which should be a shape
    page_element_data = slide_3.pageElements[0]

    # Use TypeAdapter to validate and create the correct element type
    page_element_adapter = TypeAdapter(PageElement)
    element = page_element_adapter.validate_python(page_element_data.model_dump())

    assert isinstance(element, ShapeElement), f"Expected ShapeElement, got {type(element)}"
    return element


class TestShapeElementToMarkdown:
    """Test ShapeElement.to_markdown() method."""

    def test_to_markdown_with_text(self, slide_3_shape_element):
        """Test to_markdown() with a shape element that has text."""
        result = slide_3_shape_element.to_markdown()

        # The first element should contain "Section title & body slide\n"
        assert result is not None
        assert "Section title & body slide" in result

    def test_to_markdown_without_text(self):
        """Test to_markdown() with a shape element that has no text."""
        shape_element = ShapeElement(
            objectId="test_shape",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            shape=Shape(
                shapeType=ShapeType.RECTANGLE,
                shapeProperties=ShapeProperties(),
                # No text property
            ),
        )

        result = shape_element.to_markdown()
        assert result is None

    def test_to_markdown_with_empty_text_elements(self):
        """Test to_markdown() with a shape that has text but no text elements."""
        from gslides_api.domain import Text

        shape_element = ShapeElement(
            objectId="test_shape",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            shape=Shape(
                shapeType=ShapeType.TEXT_BOX,
                shapeProperties=ShapeProperties(),
                text=Text(textElements=[]),  # Empty text elements
            ),
        )

        result = shape_element.to_markdown()
        assert result is None


class TestShapeElementWriteMethods:
    """Test ShapeElement write methods."""

    def test_write_plain_text_requests_not_implemented(self, slide_3_shape_element):
        """Test that _write_plain_text_requests raises NotImplementedError."""
        with pytest.raises(
            NotImplementedError, match="Writing plain text to shape elements is not supported yet"
        ):
            slide_3_shape_element._write_plain_text_requests("test text")

    @patch("gslides_api.element.markdown_processor")
    def test_write_markdown_requests(self, mock_processor, slide_3_shape_element):
        """Test write_markdown_requests method."""
        mock_processor.create_slides_requests.return_value = [{"test": "request"}]

        result = slide_3_shape_element._write_markdown_requests("# Test markdown")

        mock_processor.create_slides_requests.assert_called_once_with(
            slide_3_shape_element.objectId, "# Test markdown"
        )
        assert result == [{"test": "request"}]

    # Note: write_text method doesn't exist in the actual implementation
    # The tests above cover the actual methods that exist


class TestImageElementToMarkdown:
    """Test ImageElement.to_markdown() method."""

    def test_to_markdown_with_source_url(self):
        """Test to_markdown() with an image that has a source URL."""
        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            title="Test Image",
            image=Image(sourceUrl="https://example.com/image.jpg"),
        )

        result = image_element.to_markdown()
        assert result == "![Test Image](https://example.com/image.jpg)"

    def test_to_markdown_without_source_url(self):
        """Test to_markdown() with an image that has no source URL."""
        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(contentUrl="https://example.com/content.jpg"),  # No sourceUrl
        )

        result = image_element.to_markdown()
        assert result is None

    def test_to_markdown_without_title(self):
        """Test to_markdown() with an image that has no title."""
        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(sourceUrl="https://example.com/image.jpg"),
            # No title
        )

        result = image_element.to_markdown()
        assert result == "![Image](https://example.com/image.jpg)"


class TestImageElementReplaceImageRequests:
    """Test ImageElement._replace_image_requests() method."""

    @patch("gslides_api.element.image_url_is_valid")
    def test_replace_image_requests_success(self, mock_validate):
        """Test successful image replacement request generation."""
        mock_validate.return_value = True

        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(contentUrl="https://example.com/old_image.jpg"),
        )

        result = image_element._replace_image_requests("https://example.com/new_image.jpg")

        mock_validate.assert_called_once_with("https://example.com/new_image.jpg")
        expected_request = [
            {
                "replaceImage": {
                    "imageObjectId": "test_image",
                    "url": "https://example.com/new_image.jpg",
                }
            }
        ]
        assert result == expected_request

    def test_replace_image_requests_invalid_url_protocol(self):
        """Test _replace_image_requests with invalid URL protocol."""
        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(contentUrl="https://example.com/image.jpg"),
        )

        with pytest.raises(ValueError, match="Image URL must start with http:// or https://"):
            image_element._replace_image_requests("ftp://example.com/image.jpg")

    @patch("gslides_api.element.image_url_is_valid")
    def test_replace_image_requests_url_not_accessible(self, mock_validate):
        """Test _replace_image_requests with URL that's not accessible."""
        mock_validate.return_value = False

        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(contentUrl="https://example.com/image.jpg"),
        )

        with pytest.raises(ValueError, match="Image URL is not accessible or invalid"):
            image_element._replace_image_requests("https://invalid-url.com/image.jpg")

    @patch("gslides_api.element.image_url_is_valid")
    def test_replace_image_requests_with_method(self, mock_validate):
        """Test _replace_image_requests with ImageReplaceMethod."""
        mock_validate.return_value = True

        image_element = ImageElement(
            objectId="test_image",
            size=Size(width=100, height=100),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
            image=Image(contentUrl="https://example.com/old_image.jpg"),
        )

        result = image_element._replace_image_requests(
            "https://example.com/new_image.jpg", ImageReplaceMethod.CENTER_CROP
        )

        mock_validate.assert_called_once_with("https://example.com/new_image.jpg")
        expected_request = [
            {
                "replaceImage": {
                    "imageObjectId": "test_image",
                    "url": "https://example.com/new_image.jpg",
                    "imageReplaceMethod": "CENTER_CROP",
                }
            }
        ]
        assert result == expected_request


class TestElementsFromPresentationData:
    """Test elements created from actual presentation data."""

    def test_slide_3_elements_types(self, presentation_object):
        """Test that slide 3 elements are correctly typed."""
        slide_3 = presentation_object.slides[3]
        page_element_adapter = TypeAdapter(PageElement)

        for i, page_element_data in enumerate(slide_3.pageElements):
            element = page_element_adapter.validate_python(page_element_data.model_dump())
            assert isinstance(element, ShapeElement), f"Element {i} should be a ShapeElement"
            assert element.shape is not None
            assert element.shape.shapeType == ShapeType.TEXT_BOX

    def test_slide_3_all_elements_to_markdown(self, presentation_object):
        """Test to_markdown() on all elements from slide 3."""
        slide_3 = presentation_object.slides[3]
        page_element_adapter = TypeAdapter(PageElement)

        expected_texts = ["Section title & body slide", "This is a subtitle", "This is the body"]

        for i, page_element_data in enumerate(slide_3.pageElements):
            element = page_element_adapter.validate_python(page_element_data.model_dump())
            result = element.to_markdown()

            assert result is not None, f"Element {i} should have markdown content"
            assert expected_texts[i] in result, f"Element {i} should contain '{expected_texts[i]}'"
