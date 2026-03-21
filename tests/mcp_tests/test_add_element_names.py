"""Tests for the add_element_names MCP tool."""

import json
from unittest.mock import Mock, patch

import pytest

from gslides_api.adapters.add_names import SlideElementNames
from gslides_api.mcp.server import add_element_names


class TestAddElementNames:
    """Tests for the add_element_names tool."""

    @patch("gslides_api.mcp.server.name_slides")
    @patch("gslides_api.mcp.server.GSlidesAPIClient")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_successful_call(
        self, mock_get_client, mock_gslides_class, mock_name_slides,
    ):
        """Test successful call returns SuccessResponse with slide/element names."""
        mock_client = Mock()
        mock_get_client.return_value = mock_client
        mock_gslides_client = Mock()
        mock_gslides_class.return_value = mock_gslides_client

        mock_name_slides.return_value = {
            "Intro": SlideElementNames(
                image_names=["Image_1"],
                text_names=["Title", "Text_1"],
                chart_names=["Chart_1"],
                table_names=[],
            ),
            "Summary": SlideElementNames(
                image_names=[],
                text_names=["Title"],
                chart_names=[],
                table_names=["Table_1"],
            ),
        }

        result = json.loads(add_element_names("pres_123"))

        assert result["success"] is True
        assert "Successfully named 2 slides" in result["message"]
        details = result["details"]["slide_element_names"]
        assert details["Intro"]["text_names"] == ["Title", "Text_1"]
        assert details["Intro"]["image_names"] == ["Image_1"]
        assert details["Intro"]["chart_names"] == ["Chart_1"]
        assert details["Intro"]["table_names"] == []
        assert details["Summary"]["table_names"] == ["Table_1"]

        mock_name_slides.assert_called_once_with(
            "pres_123",
            name_elements=True,
            api_client=mock_gslides_client,
            skip_empty_text_boxes=False,
            min_image_size_cm=4.0,
        )
        mock_client.flush_batch_update.assert_called_once()

    @patch("gslides_api.mcp.server.name_slides")
    @patch("gslides_api.mcp.server.GSlidesAPIClient")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_with_custom_parameters(
        self, mock_get_client, mock_gslides_class, mock_name_slides,
    ):
        """Test that custom parameters are passed through."""
        mock_client = Mock()
        mock_get_client.return_value = mock_client
        mock_gslides_client = Mock()
        mock_gslides_class.return_value = mock_gslides_client
        mock_name_slides.return_value = {}

        result = json.loads(add_element_names(
            "pres_123",
            skip_empty_text_boxes=True,
            min_image_size_cm=2.0,
        ))

        assert result["success"] is True
        mock_name_slides.assert_called_once_with(
            "pres_123",
            name_elements=True,
            api_client=mock_gslides_client,
            skip_empty_text_boxes=True,
            min_image_size_cm=2.0,
        )

    def test_invalid_presentation_url(self):
        """Test that invalid presentation URL returns validation error."""
        result = json.loads(add_element_names("https://example.com/bad-url"))

        assert result["error"] is True
        assert result["error_type"] == "ValidationError"

    @patch("gslides_api.mcp.server.name_slides")
    @patch("gslides_api.mcp.server.GSlidesAPIClient")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_exception_handling(
        self, mock_get_client, mock_gslides_class, mock_name_slides,
    ):
        """Test that exceptions are caught and returned as error responses."""
        mock_client = Mock()
        mock_get_client.return_value = mock_client
        mock_gslides_class.return_value = Mock()
        mock_name_slides.side_effect = RuntimeError("API connection failed")

        result = json.loads(add_element_names("pres_123"))

        assert result["error"] is True
        assert "API connection failed" in result["message"]

    @patch("gslides_api.mcp.server.name_slides")
    @patch("gslides_api.mcp.server.GSlidesAPIClient")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_google_slides_url_parsed(
        self, mock_get_client, mock_gslides_class, mock_name_slides,
    ):
        """Test that a full Google Slides URL is parsed to extract the presentation ID."""
        mock_client = Mock()
        mock_get_client.return_value = mock_client
        mock_gslides_class.return_value = Mock()
        mock_name_slides.return_value = {}

        url = "https://docs.google.com/presentation/d/abc123xyz/edit"
        result = json.loads(add_element_names(url))

        assert result["success"] is True
        mock_name_slides.assert_called_once()
        call_args = mock_name_slides.call_args
        assert call_args[0][0] == "abc123xyz"
