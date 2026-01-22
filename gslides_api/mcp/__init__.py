"""MCP server for gslides-api.

This module provides an MCP (Model Context Protocol) server that exposes
Google Slides operations as tools for AI assistants.

Usage:
    python -m gslides_api.mcp.server --credential-path /path/to/credentials

Or set the GSLIDES_CREDENTIALS_PATH environment variable.
"""

from .models import (
    ElementOutline,
    ErrorResponse,
    OutputFormat,
    PresentationOutline,
    SlideOutline,
    SuccessResponse,
    ThumbnailSizeOption,
)
from .server import initialize_server, main, mcp
from .utils import (
    find_element_by_name,
    find_slide_by_name,
    get_element_name,
    get_slide_name,
    parse_presentation_id,
)

__all__ = [
    # Server
    "mcp",
    "main",
    "initialize_server",
    # Models
    "OutputFormat",
    "ThumbnailSizeOption",
    "ErrorResponse",
    "SuccessResponse",
    "ElementOutline",
    "SlideOutline",
    "PresentationOutline",
    # Utils
    "parse_presentation_id",
    "get_slide_name",
    "get_element_name",
    "find_slide_by_name",
    "find_element_by_name",
]
