"""Tests for gslides_api.mcp.utils module."""

import pytest

from gslides_api.mcp.utils import (
    create_error_response,
    element_not_found_error,
    parse_presentation_id,
    slide_not_found_error,
    validation_error,
)


class TestParsePresentationId:
    """Tests for parse_presentation_id function."""

    def test_parse_simple_id(self):
        """Test parsing a simple presentation ID."""
        pres_id = "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"
        result = parse_presentation_id(pres_id)
        assert result == pres_id

    def test_parse_url_with_edit(self):
        """Test parsing a Google Slides URL with /edit."""
        url = "https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_url_without_edit(self):
        """Test parsing a Google Slides URL without /edit."""
        url = "https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_url_with_slide_anchor(self):
        """Test parsing a Google Slides URL with slide anchor."""
        url = "https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit#slide=id.p"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_url_with_query_params(self):
        """Test parsing a Google Slides URL with query parameters."""
        url = "https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit?usp=sharing"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_url_with_query_and_anchor(self):
        """Test parsing a URL with both query params and anchor."""
        url = "https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit?usp=sharing#slide=id.g123"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_http_url(self):
        """Test parsing an HTTP (non-HTTPS) URL."""
        url = "http://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit"
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_url_with_whitespace(self):
        """Test parsing a URL with leading/trailing whitespace."""
        url = "  https://docs.google.com/presentation/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit  "
        result = parse_presentation_id(url)
        assert result == "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

    def test_parse_id_with_underscores_and_hyphens(self):
        """Test parsing an ID with underscores and hyphens."""
        pres_id = "abc-123_XYZ-456_789"
        result = parse_presentation_id(pres_id)
        assert result == pres_id

    def test_invalid_url_format(self):
        """Test that invalid URL format raises ValueError."""
        url = "https://example.com/presentation/d/123"
        with pytest.raises(ValueError) as exc_info:
            parse_presentation_id(url)
        assert "Invalid Google Slides URL format" in str(exc_info.value)

    def test_invalid_url_no_id(self):
        """Test that URL without ID raises ValueError."""
        url = "https://docs.google.com/presentation/d/"
        with pytest.raises(ValueError) as exc_info:
            parse_presentation_id(url)
        assert "Invalid Google Slides URL format" in str(exc_info.value)

    def test_invalid_url_different_service(self):
        """Test that URL for different Google service raises ValueError."""
        url = "https://docs.google.com/document/d/1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM/edit"
        with pytest.raises(ValueError) as exc_info:
            parse_presentation_id(url)
        assert "Invalid Google Slides URL format" in str(exc_info.value)


class TestErrorResponses:
    """Tests for error response creation functions."""

    def test_create_error_response(self):
        """Test creating a basic error response."""
        error = create_error_response(
            error_type="TestError",
            message="Test error message",
            key1="value1",
            key2="value2",
        )
        assert error.error is True
        assert error.error_type == "TestError"
        assert error.message == "Test error message"
        assert error.details["key1"] == "value1"
        assert error.details["key2"] == "value2"

    def test_slide_not_found_error(self):
        """Test creating a slide not found error."""
        error = slide_not_found_error(
            presentation_id="pres123",
            slide_name="Introduction",
            available_slides=["Cover", "Overview", "Conclusion"],
        )
        assert error.error is True
        assert error.error_type == "SlideNotFound"
        assert "Introduction" in error.message
        assert error.details["presentation_id"] == "pres123"
        assert error.details["searched_slide_name"] == "Introduction"
        assert error.details["available_slides"] == ["Cover", "Overview", "Conclusion"]

    def test_element_not_found_error(self):
        """Test creating an element not found error."""
        error = element_not_found_error(
            presentation_id="pres123",
            slide_name="Introduction",
            element_name="Title",
            available_elements=["Subtitle", "Body", "Image"],
        )
        assert error.error is True
        assert error.error_type == "ElementNotFound"
        assert "Title" in error.message
        assert "Introduction" in error.message
        assert error.details["presentation_id"] == "pres123"
        assert error.details["slide_name"] == "Introduction"
        assert error.details["searched_element_name"] == "Title"
        assert error.details["available_elements"] == ["Subtitle", "Body", "Image"]

    def test_validation_error(self):
        """Test creating a validation error."""
        error = validation_error(
            field="presentation_id",
            message="Invalid format",
            value="bad-value",
        )
        assert error.error is True
        assert error.error_type == "ValidationError"
        assert "Invalid format" in error.message
        assert error.details["field"] == "presentation_id"
        assert error.details["invalid_value"] == "bad-value"

    def test_validation_error_no_value(self):
        """Test creating a validation error without value."""
        error = validation_error(
            field="slide_name",
            message="Field is required",
        )
        assert error.error is True
        assert error.error_type == "ValidationError"
        assert error.details["field"] == "slide_name"
        assert "invalid_value" not in error.details
