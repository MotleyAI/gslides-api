import os
from unittest.mock import MagicMock, patch

import pytest

from gslides_api.client import api_client


class TestUploadImageToDrive:
    """Test the upload_image_to_drive function."""

    def test_supported_png_format(self):
        """Test that PNG format is correctly detected and processed."""
        with patch.object(api_client, "drive_srvc") as mock_drive_service, patch(
            "gslides_api.client.MediaFileUpload"
        ) as mock_media_upload, patch("os.path.basename") as mock_basename:

            # Setup mocks
            mock_basename.return_value = "test_image.png"
            mock_drive_service.files().create().execute.return_value = {
                "id": "test_file_id"
            }
            mock_drive_service.permissions().create().execute.return_value = {}

            # Test PNG format
            result = api_client.upload_image_to_drive("test_image.png")

            # Verify correct MIME type was used
            mock_media_upload.assert_called_once_with(
                "test_image.png", mimetype="image/png"
            )

            # Verify file metadata has correct MIME type
            create_call_args = mock_drive_service.files().create.call_args
            file_metadata = create_call_args[1]["body"]
            assert file_metadata["mimeType"] == "image/png"

            # Verify return URL
            assert result == "https://drive.google.com/uc?id=test_file_id"

    def test_supported_jpeg_formats(self):
        """Test that JPEG formats (.jpg and .jpeg) are correctly detected."""
        test_cases = [
            ("test_image.jpg", "image/jpeg"),
            ("test_image.jpeg", "image/jpeg"),
            ("test_image.JPG", "image/jpeg"),  # Test case insensitive
            ("test_image.JPEG", "image/jpeg"),
        ]

        for image_path, expected_mime_type in test_cases:
            with patch.object(api_client, "drive_srvc") as mock_drive_service, patch(
                "gslides_api.client.MediaFileUpload"
            ) as mock_media_upload, patch("os.path.basename") as mock_basename:

                # Setup mocks
                mock_basename.return_value = os.path.basename(image_path)
                mock_drive_service.files().create().execute.return_value = {
                    "id": "test_file_id"
                }
                mock_drive_service.permissions().create().execute.return_value = {}

                # Test the format
                result = api_client.upload_image_to_drive(image_path)

                # Verify correct MIME type was used
                mock_media_upload.assert_called_once_with(
                    image_path, mimetype=expected_mime_type
                )

                # Verify file metadata has correct MIME type
                create_call_args = mock_drive_service.files().create.call_args
                file_metadata = create_call_args[1]["body"]
                assert file_metadata["mimeType"] == expected_mime_type

    def test_supported_gif_format(self):
        """Test that GIF format is correctly detected and processed."""
        with patch.object(api_client, "drive_srvc") as mock_drive_service, patch(
            "gslides_api.client.MediaFileUpload"
        ) as mock_media_upload, patch("os.path.basename") as mock_basename:

            # Setup mocks
            mock_basename.return_value = "test_image.gif"
            mock_drive_service.files().create().execute.return_value = {
                "id": "test_file_id"
            }
            mock_drive_service.permissions().create().execute.return_value = {}

            # Test GIF format
            result = api_client.upload_image_to_drive("test_image.gif")

            # Verify correct MIME type was used
            mock_media_upload.assert_called_once_with(
                "test_image.gif", mimetype="image/gif"
            )

            # Verify file metadata has correct MIME type
            create_call_args = mock_drive_service.files().create.call_args
            file_metadata = create_call_args[1]["body"]
            assert file_metadata["mimeType"] == "image/gif"

    def test_unsupported_format_raises_value_error(self):
        """Test that unsupported image formats raise ValueError."""
        unsupported_formats = [
            "test_image.bmp",
            "test_image.tiff",
            "test_image.webp",
            "test_image.svg",
            "test_image.pdf",
            "test_image.txt",
            "test_image",  # No extension
        ]

        for image_path in unsupported_formats:
            with pytest.raises(ValueError) as exc_info:
                api_client.upload_image_to_drive(image_path)

            # Verify error message contains expected information
            error_message = str(exc_info.value)
            assert "Unsupported image format" in error_message
            assert "Supported formats are:" in error_message
            assert ".png" in error_message
            assert ".jpg" in error_message
            assert ".jpeg" in error_message
            assert ".gif" in error_message

    def test_case_insensitive_extension_detection(self):
        """Test that file extension detection is case insensitive."""
        test_cases = [
            "test_image.PNG",
            "test_image.Png",
            "test_image.pNg",
            "test_image.JPG",
            "test_image.Jpg",
            "test_image.JPEG",
            "test_image.Jpeg",
            "test_image.GIF",
            "test_image.Gif",
        ]

        for image_path in test_cases:
            with patch.object(api_client, "drive_srvc") as mock_drive_service, patch(
                "gslides_api.client.MediaFileUpload"
            ) as mock_media_upload, patch("os.path.basename") as mock_basename:

                # Setup mocks
                mock_basename.return_value = os.path.basename(image_path)
                mock_drive_service.files().create().execute.return_value = {
                    "id": "test_file_id"
                }
                mock_drive_service.permissions().create().execute.return_value = {}

                # Should not raise an exception
                result = api_client.upload_image_to_drive(image_path)
                assert result == "https://drive.google.com/uc?id=test_file_id"

    def test_path_with_directories(self):
        """Test that function works with full file paths including directories."""
        test_paths = [
            "/path/to/image.png",
            "relative/path/image.jpg",
            "C:\\Windows\\path\\image.gif",
            "./local/image.jpeg",
        ]

        for image_path in test_paths:
            with patch.object(api_client, "drive_srvc") as mock_drive_service, patch(
                "gslides_api.client.MediaFileUpload"
            ) as mock_media_upload, patch("os.path.basename") as mock_basename:

                # Setup mocks
                mock_basename.return_value = os.path.basename(image_path)
                mock_drive_service.files().create().execute.return_value = {
                    "id": "test_file_id"
                }
                mock_drive_service.permissions().create().execute.return_value = {}

                # Should not raise an exception
                result = api_client.upload_image_to_drive(image_path)
                assert result == "https://drive.google.com/uc?id=test_file_id"

    def test_error_message_format(self):
        """Test that error messages are properly formatted."""
        with pytest.raises(ValueError) as exc_info:
            api_client.upload_image_to_drive("test_image.bmp")

        error_message = str(exc_info.value)

        # Check that the error message includes the unsupported extension
        assert "'.bmp'" in error_message

        # Check that all supported formats are listed
        assert ".png, .jpg, .jpeg, .gif" in error_message
