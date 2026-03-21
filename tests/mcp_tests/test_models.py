"""Tests for gslides_api.mcp.models module."""

import pytest

from gslides_api.mcp.models import (
    ErrorResponse,
    OutputFormat,
    SuccessResponse,
    ThumbnailSizeOption,
)


class TestOutputFormat:
    """Tests for OutputFormat enum."""

    def test_raw_format(self):
        """Test RAW format value."""
        assert OutputFormat.RAW.value == "raw"
        assert OutputFormat("raw") == OutputFormat.RAW

    def test_domain_format(self):
        """Test DOMAIN format value."""
        assert OutputFormat.DOMAIN.value == "domain"
        assert OutputFormat("domain") == OutputFormat.DOMAIN

    def test_markdown_format(self):
        """Test MARKDOWN format value."""
        assert OutputFormat.MARKDOWN.value == "markdown"
        assert OutputFormat("markdown") == OutputFormat.MARKDOWN

    def test_invalid_format(self):
        """Test that invalid format raises ValueError."""
        with pytest.raises(ValueError):
            OutputFormat("invalid")

    def test_outline_format_removed(self):
        """Test that OUTLINE format no longer exists."""
        with pytest.raises(ValueError):
            OutputFormat("outline")


class TestThumbnailSizeOption:
    """Tests for ThumbnailSizeOption enum."""

    def test_small_size(self):
        """Test SMALL size value."""
        assert ThumbnailSizeOption.SMALL.value == "SMALL"

    def test_medium_size(self):
        """Test MEDIUM size value."""
        assert ThumbnailSizeOption.MEDIUM.value == "MEDIUM"

    def test_large_size(self):
        """Test LARGE size value."""
        assert ThumbnailSizeOption.LARGE.value == "LARGE"


class TestErrorResponse:
    """Tests for ErrorResponse model."""

    def test_basic_error_response(self):
        """Test creating a basic error response."""
        error = ErrorResponse(
            error_type="TestError",
            message="Test message",
        )
        assert error.error is True
        assert error.error_type == "TestError"
        assert error.message == "Test message"
        assert error.details == {}

    def test_error_response_with_details(self):
        """Test creating an error response with details."""
        error = ErrorResponse(
            error_type="ValidationError",
            message="Invalid input",
            details={"field": "name", "value": "bad"},
        )
        assert error.details == {"field": "name", "value": "bad"}

    def test_error_response_model_dump(self):
        """Test that error response can be serialized."""
        error = ErrorResponse(
            error_type="TestError",
            message="Test message",
            details={"key": "value"},
        )
        data = error.model_dump()
        assert data["error"] is True
        assert data["error_type"] == "TestError"
        assert data["message"] == "Test message"
        assert data["details"] == {"key": "value"}


class TestSuccessResponse:
    """Tests for SuccessResponse model."""

    def test_basic_success_response(self):
        """Test creating a basic success response."""
        response = SuccessResponse(
            message="Operation completed",
        )
        assert response.success is True
        assert response.message == "Operation completed"
        assert response.details == {}

    def test_success_response_with_details(self):
        """Test creating a success response with details."""
        response = SuccessResponse(
            message="Slide copied",
            details={"new_slide_id": "abc123", "position": 5},
        )
        assert response.details == {"new_slide_id": "abc123", "position": 5}
