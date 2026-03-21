"""Models for the gslides-api MCP server."""

from enum import Enum
from typing import Any, Dict

from pydantic import BaseModel, Field


class OutputFormat(str, Enum):
    """Output format for presentation/slide/element data."""

    RAW = "raw"  # Raw Google Slides API JSON response
    DOMAIN = "domain"  # gslides-api domain object model_dump()
    MARKDOWN = "markdown"  # Slide markdown layout representation


class ThumbnailSizeOption(str, Enum):
    """Thumbnail size options."""

    SMALL = "SMALL"  # 200px width
    MEDIUM = "MEDIUM"  # 800px width
    LARGE = "LARGE"  # 1600px width


class ErrorResponse(BaseModel):
    """Structured error response for tool failures."""

    error: bool = True
    error_type: str = Field(description="Type of error (e.g., SlideNotFound, ValidationError)")
    message: str = Field(description="Human-readable error message")
    details: Dict[str, Any] = Field(
        default_factory=dict, description="Additional context about the error"
    )


class SuccessResponse(BaseModel):
    """Success response for modification operations."""

    success: bool = True
    message: str = Field(description="Success message")
    details: Dict[str, Any] = Field(
        default_factory=dict, description="Additional details about the operation"
    )
