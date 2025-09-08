import os
import tempfile
from unittest.mock import Mock, patch

import pytest
import requests

from gslides_api.domain.domain import Image, ImageData, Transform
from gslides_api.element.base import ElementKind
from gslides_api.element.image import ImageElement


def test_image_data_creation():
    """Test creating ImageData directly."""
    content = b"fake image data"
    mime_type = "image/jpeg"
    filename = "test.jpg"

    image_data = ImageData(content=content, mime_type=mime_type, filename=filename)

    assert image_data.content == content
    assert image_data.mime_type == mime_type
    assert image_data.filename == filename


def test_image_data_save_to_file():
    """Test saving ImageData to a file."""
    content = b"fake image data"
    mime_type = "image/jpeg"
    filename = "test.jpg"

    image_data = ImageData(content=content, mime_type=mime_type, filename=filename)

    with tempfile.TemporaryDirectory() as temp_dir:
        # Test saving to a specific file path
        file_path = os.path.join(temp_dir, "output.jpg")
        saved_path = image_data.save_to_file(file_path)

        assert saved_path == file_path
        assert os.path.exists(file_path)

        with open(file_path, "rb") as f:
            assert f.read() == content


def test_image_data_save_to_directory():
    """Test saving ImageData to a directory using filename hint."""
    content = b"fake image data"
    mime_type = "image/jpeg"
    filename = "test.jpg"

    image_data = ImageData(content=content, mime_type=mime_type, filename=filename)

    with tempfile.TemporaryDirectory() as temp_dir:
        saved_path = image_data.save_to_file(temp_dir)
        expected_path = os.path.join(temp_dir, filename)

        assert saved_path == expected_path
        assert os.path.exists(expected_path)

        with open(expected_path, "rb") as f:
            assert f.read() == content


def test_image_data_save_to_directory_no_filename():
    """Test saving ImageData to a directory without filename hint."""
    content = b"fake image data"
    mime_type = "image/png"

    image_data = ImageData(content=content, mime_type=mime_type, filename=None)

    with tempfile.TemporaryDirectory() as temp_dir:
        saved_path = image_data.save_to_file(temp_dir)
        expected_path = os.path.join(temp_dir, "image.png")

        assert saved_path == expected_path
        assert os.path.exists(expected_path)

        with open(expected_path, "rb") as f:
            assert f.read() == content


def test_image_data_get_extension():
    """Test getting file extension from MIME type."""
    image_data = ImageData(content=b"test", mime_type="image/jpeg")
    assert image_data.get_extension() in [".jpg", ".jpeg"]

    image_data = ImageData(content=b"test", mime_type="image/png")
    assert image_data.get_extension() == ".png"

    image_data = ImageData(content=b"test", mime_type="unknown/type")
    assert image_data.get_extension() == ".bin"


@patch("gslides_api.element.image.requests.get")
def test_image_element_get_image_data(mock_get):
    """Test ImageElement.get_image_data method."""
    # Mock the HTTP response
    mock_response = Mock()
    mock_response.content = b"fake image data"
    mock_response.headers = {"content-type": "image/jpeg"}
    mock_response.raise_for_status.return_value = None
    mock_get.return_value = mock_response

    # Create an ImageElement
    image = Image(contentUrl="https://example.com/image.jpg")
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    # Test the method
    image_data = element.get_image_data()

    assert image_data.content == b"fake image data"
    assert image_data.mime_type == "image/jpeg"
    assert image_data.filename == "image.jpg"

    # Verify the HTTP request was made
    mock_get.assert_called_once_with("https://example.com/image.jpg", timeout=30)


@patch("gslides_api.element.image.requests.get")
def test_image_element_get_image_data_fallback_to_source_url(mock_get):
    """Test ImageElement.get_image_data falls back to sourceUrl."""
    # Mock the HTTP response
    mock_response = Mock()
    mock_response.content = b"fake image data"
    mock_response.headers = {"content-type": "image/png"}
    mock_response.raise_for_status.return_value = None
    mock_get.return_value = mock_response

    # Create an ImageElement with only sourceUrl
    image = Image(sourceUrl="https://example.com/source.png")
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    # Test the method
    image_data = element.get_image_data()

    assert image_data.content == b"fake image data"
    assert image_data.mime_type == "image/png"

    # Verify the HTTP request was made with sourceUrl
    mock_get.assert_called_once_with("https://example.com/source.png", timeout=30)


def test_image_element_get_image_data_no_url():
    """Test ImageElement.get_image_data raises error when no URL available."""
    # Create an ImageElement without URLs
    image = Image()
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    with pytest.raises(ValueError, match="No image URL available"):
        element.get_image_data()


@patch("gslides_api.element.image.requests.get")
def test_image_element_get_image_data_http_error(mock_get):
    """Test ImageElement.get_image_data handles HTTP errors."""
    # Mock HTTP error
    mock_get.side_effect = requests.RequestException("Network error")

    # Create an ImageElement
    image = Image(contentUrl="https://example.com/image.jpg")
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    with pytest.raises(requests.RequestException):
        element.get_image_data()


@patch("gslides_api.element.image.requests.get")
def test_image_element_get_image_data_empty_content(mock_get):
    """Test ImageElement.get_image_data handles empty content."""
    # Mock empty response
    mock_response = Mock()
    mock_response.content = b""
    mock_response.headers = {"content-type": "image/jpeg"}
    mock_response.raise_for_status.return_value = None
    mock_get.return_value = mock_response

    # Create an ImageElement
    image = Image(contentUrl="https://example.com/image.jpg")
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    with pytest.raises(ValueError, match="Downloaded image is empty"):
        element.get_image_data()


@patch("gslides_api.element.image.requests.get")
def test_image_element_get_image_data_mime_type_fallback(mock_get):
    """Test ImageElement.get_image_data MIME type detection from URL."""
    # Mock response with generic content type
    mock_response = Mock()
    mock_response.content = b"fake image data"
    mock_response.headers = {"content-type": "application/octet-stream"}
    mock_response.raise_for_status.return_value = None
    mock_get.return_value = mock_response

    # Create an ImageElement with .png in URL
    image = Image(contentUrl="https://example.com/path/image.png")
    transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
    element = ImageElement(
        objectId="test-id", image=image, transform=transform, type=ElementKind.IMAGE
    )

    # Test the method
    image_data = element.get_image_data()

    # Should have guessed PNG from URL
    assert image_data.mime_type == "image/png"
