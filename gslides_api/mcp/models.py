"""Models for the gslides-api MCP server."""

from enum import Enum
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class OutputFormat(str, Enum):
    """Output format for presentation/slide/element data."""

    RAW = "raw"  # Raw Google Slides API JSON response
    DOMAIN = "domain"  # gslides-api domain object model_dump()
    OUTLINE = "outline"  # Bare-bones structure with names and markdown content


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


class ElementOutline(BaseModel):
    """Outline representation of a page element."""

    element_name: Optional[str] = Field(None, description="Element name from alt-text title")
    element_id: str = Field(description="Element object ID")
    type: str = Field(description="Element type (shape, image, table, etc.)")
    alt_description: Optional[str] = Field(None, description="Alt-text description if present")
    content_markdown: Optional[str] = Field(
        None, description="Markdown content for text elements"
    )


class SlideOutline(BaseModel):
    """Outline representation of a slide."""

    slide_name: Optional[str] = Field(None, description="Slide name from speaker notes")
    slide_id: str = Field(description="Slide object ID")
    elements: List[ElementOutline] = Field(default_factory=list)


class PresentationOutline(BaseModel):
    """Outline representation of a presentation."""

    presentation_id: str = Field(description="Presentation ID")
    title: str = Field(description="Presentation title")
    slides: List[SlideOutline] = Field(default_factory=list)


class SuccessResponse(BaseModel):
    """Success response for modification operations."""

    success: bool = True
    message: str = Field(description="Success message")
    details: Dict[str, Any] = Field(
        default_factory=dict, description="Additional details about the operation"
    )
