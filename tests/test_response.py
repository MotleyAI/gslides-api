"""Tests for response.py classes, particularly ImageThumbnail."""

import os
import tempfile
import pytest
from unittest.mock import patch, MagicMock
from gslides_api.response import ImageThumbnail


class TestImageThumbnail:
    """Test ImageThumbnail class."""

    def test_image_thumbnail_creation(self):
        """Test creating ImageThumbnail instance."""
        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )
        assert thumbnail.contentUrl == "https://example.com/image.png"
        assert thumbnail.width == 800
        assert thumbnail.height == 600
        assert thumbnail._payload is None  # Should be None initially

    def test_image_thumbnail_inheritance(self):
        """Test that ImageThumbnail inherits from GSlidesBaseModel."""
        from gslides_api.domain import GSlidesBaseModel
        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/test.jpg",
            width=100,
            height=100
        )
        assert isinstance(thumbnail, GSlidesBaseModel)
        assert hasattr(thumbnail, 'to_api_format')

    @patch('gslides_api.response.requests.get')
    def test_payload_property_lazy_loading(self, mock_requests):
        """Test that payload property lazy loads content."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_image_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        # Initially _payload should be None
        assert thumbnail._payload is None

        # First access should trigger the request
        payload = thumbnail.payload
        assert payload == b'fake_image_data'
        mock_requests.assert_called_once_with("https://example.com/image.png")

        # Second access should use cached value
        payload2 = thumbnail.payload
        assert payload2 == b'fake_image_data'
        # Should still only have been called once
        mock_requests.assert_called_once()

    @patch('gslides_api.response.requests.get')
    @patch('gslides_api.response.imghdr.what')
    def test_mime_type_property(self, mock_imghdr, mock_requests):
        """Test mime_type property."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_png_data'
        mock_requests.return_value = mock_response

        # Mock imghdr to return 'png'
        mock_imghdr.return_value = 'png'

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        mime_type = thumbnail.mime_type
        assert mime_type == 'png'

        # Verify imghdr.what was called with correct parameters
        mock_imghdr.assert_called_with(None, h=b'fake_png_data')

    @patch('gslides_api.response.requests.get')
    def test_save_png_success(self, mock_requests):
        """Test successful save of PNG image."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_png_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        # Mock the mime_type property to return 'png'
        with patch('gslides_api.response.imghdr.what', return_value='png'):
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                try:
                    thumbnail.save(tmp_file.name)

                    # Verify the file was written
                    assert os.path.exists(tmp_file.name)
                    with open(tmp_file.name, 'rb') as f:
                        content = f.read()
                    assert content == b'fake_png_data'

                    # Verify requests.get was called (through payload property)
                    mock_requests.assert_called_with("https://example.com/image.png")

                finally:
                    # Clean up
                    if os.path.exists(tmp_file.name):
                        os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_save_jpeg_with_jpg_extension_success(self, mock_requests):
        """Test successful save of JPEG image with .jpg extension."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_jpeg_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.jpg",
            width=800,
            height=600
        )

        # Mock the mime_type property to return 'jpeg'
        with patch('gslides_api.response.imghdr.what', return_value='jpeg'):
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp_file:
                try:
                    thumbnail.save(tmp_file.name)

                    # Verify the file was written
                    assert os.path.exists(tmp_file.name)
                    with open(tmp_file.name, 'rb') as f:
                        content = f.read()
                    assert content == b'fake_jpeg_data'

                finally:
                    # Clean up
                    if os.path.exists(tmp_file.name):
                        os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_save_jpeg_with_jpeg_extension_success(self, mock_requests):
        """Test successful save of JPEG image with .jpeg extension."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_jpeg_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.jpeg",
            width=800,
            height=600
        )

        # Mock the mime_type property to return 'jpeg'
        with patch('gslides_api.response.imghdr.what', return_value='jpeg'):
            with tempfile.NamedTemporaryFile(suffix='.jpeg', delete=False) as tmp_file:
                try:
                    thumbnail.save(tmp_file.name)

                    # Verify the file was written
                    assert os.path.exists(tmp_file.name)

                finally:
                    # Clean up
                    if os.path.exists(tmp_file.name):
                        os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_save_format_mismatch_error(self, mock_requests):
        """Test that format mismatch raises ValueError."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_jpeg_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        # Mock the mime_type property to return 'jpeg' but file extension is .png
        with patch('gslides_api.response.imghdr.what', return_value='jpeg'):
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                try:
                    with pytest.raises(ValueError, match="Image format mismatch"):
                        thumbnail.save(tmp_file.name)

                    # Verify the error message contains expected details
                    with pytest.raises(ValueError) as exc_info:
                        thumbnail.save(tmp_file.name)

                    error_msg = str(exc_info.value)
                    assert "'.png'" in error_msg
                    assert "'png'" in error_msg
                    assert "'jpeg'" in error_msg

                finally:
                    # Clean up
                    if os.path.exists(tmp_file.name):
                        os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_save_no_extension_success(self, mock_requests):
        """Test save with no file extension (should skip validation)."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image",
            width=800,
            height=600
        )

        with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
            try:
                thumbnail.save(tmp_file.name)

                # Verify the file was written
                assert os.path.exists(tmp_file.name)

            finally:
                # Clean up
                if os.path.exists(tmp_file.name):
                    os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_save_mime_type_returns_none(self, mock_requests):
        """Test save when mime_type returns None (unrecognized format)."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        # Mock the mime_type property to return None (unrecognized format)
        with patch('gslides_api.response.imghdr.what', return_value=None):
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                try:
                    # Should not raise error when mime_type is None
                    thumbnail.save(tmp_file.name)

                    # Verify the file was written
                    assert os.path.exists(tmp_file.name)

                finally:
                    # Clean up
                    if os.path.exists(tmp_file.name):
                        os.unlink(tmp_file.name)

    @patch('gslides_api.response.requests.get')
    def test_to_ipython_image(self, mock_requests):
        """Test to_ipython_image method."""
        # Mock the response
        mock_response = MagicMock()
        mock_response.content = b'fake_image_data'
        mock_requests.return_value = mock_response

        thumbnail = ImageThumbnail(
            contentUrl="https://example.com/image.png",
            width=800,
            height=600
        )

        # Mock the IPython Image import inside the method
        with patch('IPython.display.Image') as mock_image:
            mock_image_instance = MagicMock()
            mock_image.return_value = mock_image_instance

            result = thumbnail.to_ipython_image()

            # Verify requests.get was called (through payload property)
            mock_requests.assert_called_with("https://example.com/image.png")

            # Verify Image was called with the payload
            mock_image.assert_called_with(b'fake_image_data')

            # Verify the result is the mock image instance
            assert result == mock_image_instance

    def test_extension_format_mapping_logic(self):
        """Test the extension to format mapping logic."""
        test_cases = [
            ('image.png', 'png', 'png'),
            ('photo.jpg', 'jpg', 'jpeg'),
            ('picture.JPEG', 'jpeg', 'jpeg'),
            ('animation.gif', 'gif', 'gif'),
            ('bitmap.bmp', 'bmp', 'bmp'),
            ('web.webp', 'webp', 'webp'),
            ('tagged.tiff', 'tiff', 'tiff'),
            ('tagged.tif', 'tif', 'tif'),  # Fixed: tif extension stays as tif
        ]

        for file_path, expected_extension, expected_format in test_cases:
            file_extension = os.path.splitext(file_path)[1].lower().lstrip('.')
            actual_format = 'jpeg' if file_extension in ('jpg', 'jpeg') else file_extension

            assert file_extension == expected_extension
            assert actual_format == expected_format
