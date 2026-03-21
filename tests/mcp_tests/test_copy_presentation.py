"""Tests for the copy_presentation MCP tool."""

import json
from unittest.mock import Mock, patch

import pytest

from gslides_api.mcp.server import copy_presentation


@pytest.fixture
def mock_api_client():
    """Create a mock API client."""
    client = Mock()
    client.copy_presentation.return_value = {"id": "new_pres_id_123"}
    client.flush_batch_update.return_value = None
    return client


@pytest.fixture
def mock_presentation():
    """Create a mock Presentation object."""
    pres = Mock()
    pres.title = "My Presentation"
    pres.presentationId = "original_pres_id"
    return pres


class TestCopyPresentation:
    """Tests for the copy_presentation tool."""

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_with_default_title(self, mock_get_client, mock_pres_class, mock_api_client, mock_presentation):
        """Test copying a presentation with default title."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = mock_presentation

        result = json.loads(copy_presentation("original_pres_id"))

        assert result["success"] is True
        assert result["message"] == "Successfully copied presentation 'My Presentation'"
        assert result["details"]["original_presentation_id"] == "original_pres_id"
        assert result["details"]["new_presentation_id"] == "new_pres_id_123"
        assert result["details"]["new_title"] == "Copy of My Presentation"
        assert "docs.google.com/presentation/d/new_pres_id_123/edit" in result["details"]["new_presentation_url"]

        mock_api_client.copy_presentation.assert_called_once_with("original_pres_id", "Copy of My Presentation", None)

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_with_custom_title(self, mock_get_client, mock_pres_class, mock_api_client, mock_presentation):
        """Test copying a presentation with a custom title."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = mock_presentation

        result = json.loads(copy_presentation("original_pres_id", copy_title="Custom Title"))

        assert result["success"] is True
        assert result["details"]["new_title"] == "Custom Title"
        mock_api_client.copy_presentation.assert_called_once_with("original_pres_id", "Custom Title", None)

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_with_folder_id(self, mock_get_client, mock_pres_class, mock_api_client, mock_presentation):
        """Test copying a presentation into a specific folder."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = mock_presentation

        result = json.loads(copy_presentation("original_pres_id", copy_title="In Folder", folder_id="folder_abc"))

        assert result["success"] is True
        mock_api_client.copy_presentation.assert_called_once_with("original_pres_id", "In Folder", "folder_abc")

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_with_url_input(self, mock_get_client, mock_pres_class, mock_api_client, mock_presentation):
        """Test copying a presentation using a Google Slides URL."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = mock_presentation

        url = "https://docs.google.com/presentation/d/abc123_xyz/edit"
        result = json.loads(copy_presentation(url, copy_title="From URL"))

        assert result["success"] is True
        mock_pres_class.from_id.assert_called_once_with("abc123_xyz", api_client=mock_api_client)
        mock_api_client.copy_presentation.assert_called_once_with("abc123_xyz", "From URL", None)

    def test_copy_with_invalid_url(self):
        """Test copying with an invalid Google Slides URL."""
        result = json.loads(copy_presentation("https://example.com/not-a-slides-url"))

        assert result["error"] is True
        assert result["error_type"] == "ValidationError"

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_with_untitled_presentation(self, mock_get_client, mock_pres_class, mock_api_client):
        """Test copying a presentation with no title defaults correctly."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.title = None
        mock_pres.presentationId = "pres_no_title"
        mock_pres_class.from_id.return_value = mock_pres

        result = json.loads(copy_presentation("pres_no_title"))

        assert result["success"] is True
        assert result["details"]["new_title"] == "Copy of Untitled"
        mock_api_client.copy_presentation.assert_called_once_with("pres_no_title", "Copy of Untitled", None)

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_presentation_api_error(self, mock_get_client, mock_pres_class, mock_api_client, mock_presentation):
        """Test error handling when the API call fails."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = mock_presentation
        mock_api_client.copy_presentation.side_effect = Exception("Drive API quota exceeded")

        result = json.loads(copy_presentation("original_pres_id", copy_title="Will Fail"))

        assert result["error"] is True
        assert "PresentationError" in result["error_type"]
        assert "Drive API quota exceeded" in result["message"]

    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_copy_presentation_load_error(self, mock_get_client, mock_pres_class, mock_api_client):
        """Test error handling when loading the original presentation fails."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.side_effect = Exception("Presentation not found")

        result = json.loads(copy_presentation("nonexistent_id"))

        assert result["error"] is True
        assert "Presentation not found" in result["message"]
