"""Tests for gslides_api.mcp.models module."""

import pytest

from gslides_api.mcp.models import (
    ElementOutline,
    ErrorResponse,
    OutputFormat,
    PresentationOutline,
    SlideOutline,
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

    def test_outline_format(self):
        """Test OUTLINE format value."""
        assert OutputFormat.OUTLINE.value == "outline"
        assert OutputFormat("outline") == OutputFormat.OUTLINE

    def test_invalid_format(self):
        """Test that invalid format raises ValueError."""
        with pytest.raises(ValueError):
            OutputFormat("invalid")


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


class TestElementOutline:
    """Tests for ElementOutline model."""

    def test_minimal_element_outline(self):
        """Test creating a minimal element outline."""
        outline = ElementOutline(
            element_id="elem123",
            type="shape",
        )
        assert outline.element_id == "elem123"
        assert outline.type == "shape"
        assert outline.element_name is None
        assert outline.alt_description is None
        assert outline.content_markdown is None

    def test_full_element_outline(self):
        """Test creating a full element outline."""
        outline = ElementOutline(
            element_name="Title",
            element_id="elem123",
            type="shape",
            alt_description="Main title text box",
            content_markdown="# Welcome",
        )
        assert outline.element_name == "Title"
        assert outline.element_id == "elem123"
        assert outline.type == "shape"
        assert outline.alt_description == "Main title text box"
        assert outline.content_markdown == "# Welcome"


class TestSlideOutline:
    """Tests for SlideOutline model."""

    def test_minimal_slide_outline(self):
        """Test creating a minimal slide outline."""
        outline = SlideOutline(
            slide_id="slide123",
        )
        assert outline.slide_id == "slide123"
        assert outline.slide_name is None
        assert outline.elements == []

    def test_slide_outline_with_elements(self):
        """Test creating a slide outline with elements."""
        elements = [
            ElementOutline(element_id="e1", type="shape"),
            ElementOutline(element_id="e2", type="image"),
        ]
        outline = SlideOutline(
            slide_name="Introduction",
            slide_id="slide123",
            elements=elements,
        )
        assert outline.slide_name == "Introduction"
        assert len(outline.elements) == 2
        assert outline.elements[0].element_id == "e1"


class TestPresentationOutline:
    """Tests for PresentationOutline model."""

    def test_minimal_presentation_outline(self):
        """Test creating a minimal presentation outline."""
        outline = PresentationOutline(
            presentation_id="pres123",
            title="My Presentation",
        )
        assert outline.presentation_id == "pres123"
        assert outline.title == "My Presentation"
        assert outline.slides == []

    def test_presentation_outline_with_slides(self):
        """Test creating a presentation outline with slides."""
        slides = [
            SlideOutline(slide_id="s1", slide_name="Cover"),
            SlideOutline(slide_id="s2", slide_name="Content"),
        ]
        outline = PresentationOutline(
            presentation_id="pres123",
            title="My Presentation",
            slides=slides,
        )
        assert len(outline.slides) == 2
        assert outline.slides[0].slide_name == "Cover"
        assert outline.slides[1].slide_name == "Content"

    def test_presentation_outline_model_dump(self):
        """Test that presentation outline can be serialized."""
        outline = PresentationOutline(
            presentation_id="pres123",
            title="My Presentation",
            slides=[
                SlideOutline(
                    slide_id="s1",
                    slide_name="Cover",
                    elements=[
                        ElementOutline(
                            element_name="Title",
                            element_id="e1",
                            type="shape",
                            content_markdown="# Welcome",
                        )
                    ],
                )
            ],
        )
        data = outline.model_dump()
        assert data["presentation_id"] == "pres123"
        assert data["title"] == "My Presentation"
        assert len(data["slides"]) == 1
        assert data["slides"][0]["slide_name"] == "Cover"
        assert data["slides"][0]["elements"][0]["element_name"] == "Title"
