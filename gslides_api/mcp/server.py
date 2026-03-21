"""MCP server for gslides-api.

This module provides an MCP server that exposes Google Slides operations as tools.
"""

import argparse
import json
import logging
import os
import re
import sys
import tempfile
import traceback
import uuid
from typing import Any, Dict, List, Optional, Union

from mcp.server import FastMCP
from mcp.server.fastmcp.utilities.types import Image

from gslides_api.adapters.abstract_slides import (
    AbstractPresentation,
    AbstractThumbnailSize,
)
from gslides_api.adapters.add_names import name_slides
from gslides_api.adapters.gslides_adapter import GSlidesAPIClient
from gslides_api.agnostic.element import MarkdownTableElement
from gslides_api.client import GoogleAPIClient
from gslides_api.domain.domain import (
    Color,
    DashStyle,
    Outline,
    OutlineFill,
    RgbColor,
    SolidFill,
    ThumbnailSize,
    Weight,
)
from gslides_api.element.base import ElementKind
from gslides_api.element.element import ImageElement
from gslides_api.element.shape import ShapeElement
from gslides_api.element.table import TableElement
from gslides_api.presentation import Presentation
from gslides_api.request.request import UpdateShapePropertiesRequest

from .models import (
    ErrorResponse,
    OutputFormat,
    SuccessResponse,
    ThumbnailSizeOption,
)
from .utils import (
    element_not_found_error,
    find_abstract_slide_by_name,
    find_element_by_name,
    find_slide_by_name,
    get_abstract_slide_names,
    get_available_element_names,
    get_available_slide_names,
    get_slide_name,
    parse_presentation_id,
    presentation_error,
    slide_not_found_error,
    validation_error,
)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Global API client factory - initialized with auto_flush=True for the factory itself
# Child clients created for each request will have auto_flush=False
_api_client_factory: Optional[GoogleAPIClient] = None

# Default output format - can be overridden via CLI arg
DEFAULT_OUTPUT_FORMAT: OutputFormat = OutputFormat.RAW


def get_api_client() -> GoogleAPIClient:
    """Get a new API client for the current request.

    Creates a child client with isolated batch state that shares the
    initialized Google API services from the factory client. This allows
    concurrent tool invocations without corrupting shared batch state.
    """
    if _api_client_factory is None:
        raise RuntimeError("API client not initialized. Call initialize_server() first.")
    return _api_client_factory.create_child_client(auto_flush=False)


def initialize_server(credential_path: str, default_format: OutputFormat = OutputFormat.RAW):
    """Initialize the MCP server with credentials.

    Args:
        credential_path: Path to the Google API credentials directory
        default_format: Default output format for tools
    """
    global _api_client_factory, DEFAULT_OUTPUT_FORMAT

    # Create factory client with auto_flush=True (default behavior for non-MCP use)
    _api_client_factory = GoogleAPIClient(auto_flush=True)

    # Initialize credentials on the factory client
    _api_client_factory.initialize_credentials(credential_path)

    # Set the global api_client in the gslides_api.client module for backward compatibility
    import gslides_api.client

    gslides_api.client.api_client = _api_client_factory

    DEFAULT_OUTPUT_FORMAT = default_format
    logger.info(f"MCP server initialized with credentials from {credential_path}")
    logger.info(f"Default output format: {default_format.value}")


# Create the MCP server
mcp = FastMCP("gslides-api")


def _get_effective_format(how: Optional[str]) -> OutputFormat:
    """Get the effective output format, using default if not specified."""
    if how is None:
        return DEFAULT_OUTPUT_FORMAT
    try:
        return OutputFormat(how)
    except ValueError:
        return DEFAULT_OUTPUT_FORMAT


def _format_response(data: Any, error: Optional[ErrorResponse] = None) -> str:
    """Format a response as JSON string."""
    if error is not None:
        return json.dumps(error.model_dump(), indent=2)
    if hasattr(data, "model_dump"):
        return json.dumps(data.model_dump(), indent=2)
    return json.dumps(data, indent=2, default=str)


# =============================================================================
# QUERY TOOLS
# =============================================================================


@mcp.tool()
def get_presentation(
    presentation_id_or_url: str,
    how: str = None,
) -> str:
    """Get a full presentation by URL or deck ID.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        how: Output format - 'raw' (Google API JSON), 'domain' (model_dump), or 'markdown' (slide layout markdown)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    format_type = _get_effective_format(how)
    client = get_api_client()

    try:
        if format_type == OutputFormat.RAW:
            # Get raw JSON from Google API
            result = client.get_presentation_json(pres_id)
            client.flush_batch_update()
            return _format_response(result)

        elif format_type == OutputFormat.MARKDOWN:
            gslides_client = GSlidesAPIClient(gslides_client=client)
            abs_pres = AbstractPresentation.from_id(
                api_client=gslides_client, presentation_id=pres_id
            )
            parts = []
            for i, slide in enumerate(abs_pres.slides):
                parts.append(f"## Slide {i}\n\n{slide.markdown()}")
            client.flush_batch_update()
            return "\n\n---\n\n".join(parts)

        else:  # DOMAIN
            presentation = Presentation.from_id(pres_id, api_client=client)
            client.flush_batch_update()
            return _format_response(presentation.model_dump())

    except Exception as e:
        logger.error(f"Error getting presentation: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def get_slide(
    presentation_id_or_url: str,
    slide_name: str = None,
    slide_index: int = None,
    how: str = None,
    include_thumbnail: bool = True,
) -> Union[str, List]:
    """Get a single slide by name or index.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes). Mutually exclusive with slide_index.
        slide_index: Zero-based slide index. Mutually exclusive with slide_name.
        how: Output format - 'markdown' (default), 'raw' (Google API JSON), or 'domain' (model_dump).
        include_thumbnail: Include slide thumbnail as image payload. Default True.
    """
    # Validate slide_name/slide_index
    if slide_name is not None and slide_index is not None:
        return _format_response(
            None,
            validation_error(
                "slide_name/slide_index",
                "slide_name and slide_index are mutually exclusive",
                f"slide_name={slide_name}, slide_index={slide_index}",
            ),
        )
    if slide_name is None and slide_index is None:
        return _format_response(
            None,
            validation_error(
                "slide_name/slide_index",
                "Either slide_name or slide_index must be provided",
                "both are None",
            ),
        )

    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    format_type = _get_effective_format(how)
    client = get_api_client()

    try:
        # Always load via AbstractPresentation for unified slide lookup
        gslides_client = GSlidesAPIClient(gslides_client=client)
        abs_pres = AbstractPresentation.from_id(
            api_client=gslides_client, presentation_id=pres_id
        )

        # Find slide
        if slide_name is not None:
            abs_slide = find_abstract_slide_by_name(abs_pres, slide_name)
            if abs_slide is None:
                names = get_abstract_slide_names(abs_pres)
                client.flush_batch_update()
                return _format_response(None, slide_not_found_error(pres_id, slide_name, names))
        else:
            if slide_index < 0 or slide_index >= len(abs_pres.slides):
                client.flush_batch_update()
                return _format_response(
                    None,
                    validation_error(
                        "slide_index",
                        f"Slide index {slide_index} out of range (0-{len(abs_pres.slides) - 1})",
                        str(slide_index),
                    ),
                )
            abs_slide = abs_pres.slides[slide_index]

        # Format output based on `how`
        if format_type == OutputFormat.MARKDOWN:
            result = abs_slide.markdown()
        elif format_type == OutputFormat.RAW:
            result = _format_response(
                client.get_slide_json(pres_id, abs_slide.objectId)
            )
        else:  # DOMAIN
            result = _format_response(abs_slide._gslides_slide.model_dump())

        client.flush_batch_update()

        # Optionally attach thumbnail
        if include_thumbnail:
            thumb = abs_slide.thumbnail(
                api_client=gslides_client,
                size=AbstractThumbnailSize.MEDIUM,
                include_data=True,
            )
            return [result, Image(data=thumb.content, format="png")]
        else:
            return result

    except Exception as e:
        logger.error(f"Error getting slide: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def get_element(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str,
    how: str = None,
) -> str:
    """Get a single element by slide name and element name (alt-title).

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (from alt-text title, stripped)
        how: Output format - 'raw' (Google API JSON) or 'domain' (model_dump)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    format_type = _get_effective_format(how)
    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        element = find_element_by_name(slide, element_name)

        if element is None:
            available = get_available_element_names(slide)
            client.flush_batch_update()
            return _format_response(None, element_not_found_error(pres_id, slide_name, element_name, available))

        client.flush_batch_update()

        if format_type == OutputFormat.RAW:
            # For raw, we return the element's API format
            return _format_response(element.to_api_format() if hasattr(element, "to_api_format") else element.model_dump())

        else:  # DOMAIN (also handles MARKDOWN since element-level markdown is not distinct)
            return _format_response(element.model_dump())

    except Exception as e:
        logger.error(f"Error getting element: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def get_slide_thumbnail(
    presentation_id_or_url: str,
    slide_name: str,
    add_text_box_borders: bool = False,
    size: str = "LARGE",
) -> str:
    """Get a slide thumbnail image, optionally with black borders around text boxes.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        add_text_box_borders: Add 1pt black outlines to all text boxes
        size: Thumbnail size - 'SMALL' (200px), 'MEDIUM' (800px), or 'LARGE' (1600px)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    # Validate size
    try:
        thumbnail_size = ThumbnailSize[size.upper()]
    except KeyError:
        return _format_response(
            None,
            validation_error("size", f"Invalid size '{size}'. Must be SMALL, MEDIUM, or LARGE", size),
        )

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        if add_text_box_borders:
            # Create a temporary copy, add borders, get thumbnail, delete copy
            copy_result = client.copy_presentation(pres_id, f"_temp_thumbnail_{pres_id}")
            temp_pres_id = copy_result["id"]

            try:
                # Load the temp presentation
                temp_presentation = Presentation.from_id(temp_pres_id, api_client=client)

                # Find the same slide in the copy
                temp_slide = find_slide_by_name(temp_presentation, slide_name)
                if temp_slide is None:
                    # Fall back to finding by index
                    slide_index = presentation.slides.index(slide)
                    temp_slide = temp_presentation.slides[slide_index]

                # Add black borders to all shape elements
                black_outline = Outline(
                    outlineFill=OutlineFill(
                        solidFill=SolidFill(
                            color=Color(rgbColor=RgbColor(red=0.0, green=0.0, blue=0.0)),
                            alpha=1.0,
                        )
                    ),
                    weight=Weight(magnitude=1.0, unit="PT"),
                    dashStyle=DashStyle.SOLID,
                )

                for element in temp_slide.page_elements_flat:
                    if element.type == ElementKind.SHAPE:
                        from gslides_api.domain.text import ShapeProperties

                        update_request = UpdateShapePropertiesRequest(
                            objectId=element.objectId,
                            shapeProperties=ShapeProperties(outline=black_outline),
                            fields="outline",
                        )
                        client.batch_update([update_request], temp_pres_id)

                client.flush_batch_update()

                # Get thumbnail from the temp slide
                thumbnail = temp_slide.thumbnail(size=thumbnail_size, api_client=client)
                client.flush_batch_update()

            finally:
                # Always clean up the temp presentation
                try:
                    client.delete_file(temp_pres_id)
                except Exception as cleanup_error:
                    logger.warning(f"Failed to delete temp presentation: {cleanup_error}")
        else:
            # Just get the thumbnail directly
            thumbnail = slide.thumbnail(size=thumbnail_size, api_client=client)
            client.flush_batch_update()

        # Save thumbnail to temp file with unique name to avoid concurrent collisions
        image_data = thumbnail.payload

        # Sanitize slide_name for filename (replace unsafe chars with underscore)
        safe_slide_name = re.sub(r'[^\w\-]', '_', slide_name)
        unique_suffix = uuid.uuid4().hex[:8]
        filename = f"{pres_id}_{safe_slide_name}_{unique_suffix}_thumbnail.png"
        file_path = os.path.join(tempfile.gettempdir(), filename)

        with open(file_path, 'wb') as f:
            f.write(image_data)

        result = {
            "success": True,
            "file_path": file_path,
            "slide_name": slide_name,
            "slide_id": slide.objectId,
            "width": thumbnail.width,
            "height": thumbnail.height,
            "mime_type": thumbnail.mime_type,
        }
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error getting thumbnail: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


# =============================================================================
# MARKDOWN TOOLS
# =============================================================================


@mcp.tool()
def read_element_markdown(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str,
) -> str:
    """Read the text content of a shape element as markdown.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (text box alt-title)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        element = find_element_by_name(slide, element_name)

        if element is None:
            available = get_available_element_names(slide)
            client.flush_batch_update()
            return _format_response(None, element_not_found_error(pres_id, slide_name, element_name, available))

        # Check if it's a shape element
        if not isinstance(element, ShapeElement):
            client.flush_batch_update()
            return _format_response(
                None,
                validation_error(
                    "element_name",
                    f"Element '{element_name}' is not a text element (type: {element.type.value})",
                    element_name,
                ),
            )

        markdown_content = element.read_text(as_markdown=True)
        client.flush_batch_update()

        result = {
            "success": True,
            "element_name": element_name,
            "element_id": element.objectId,
            "markdown": markdown_content,
        }
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error reading element markdown: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def write_element_markdown(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str,
    markdown: str,
) -> str:
    """Write markdown content to a shape element (text box).

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (text box alt-title)
        markdown: Markdown content to write
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        element = find_element_by_name(slide, element_name)

        if element is None:
            available = get_available_element_names(slide)
            client.flush_batch_update()
            return _format_response(None, element_not_found_error(pres_id, slide_name, element_name, available))

        # Check if it's a shape element
        if not isinstance(element, ShapeElement):
            client.flush_batch_update()
            return _format_response(
                None,
                validation_error(
                    "element_name",
                    f"Element '{element_name}' is not a text element (type: {element.type.value})",
                    element_name,
                ),
            )

        # Write the markdown content
        element.write_text(markdown, as_markdown=True, api_client=client)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully wrote markdown to element '{element_name}'",
            details={
                "element_id": element.objectId,
                "slide_name": slide_name,
                "content_length": len(markdown),
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error writing element markdown: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def write_table_markdown(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str,
    markdown_table: str,
) -> str:
    """Write a markdown-formatted table to a table element, resizing if needed.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (table alt-title)
        markdown_table: Markdown table string (with | delimiters and --- separator)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        element = find_element_by_name(slide, element_name)

        if element is None:
            available = get_available_element_names(slide)
            client.flush_batch_update()
            return _format_response(None, element_not_found_error(pres_id, slide_name, element_name, available))

        # Check if it's a table element
        if not isinstance(element, TableElement):
            client.flush_batch_update()
            return _format_response(
                None,
                validation_error(
                    "element_name",
                    f"Element '{element_name}' is not a table element (type: {element.type.value})",
                    element_name,
                ),
            )

        # Parse the markdown table
        markdown_elem = MarkdownTableElement.from_markdown(element_name, markdown_table)

        # Compare shapes and resize if needed
        current_shape = (element.table.rows, element.table.columns)
        target_shape = markdown_elem.shape
        font_scale_factor = 1.0

        if current_shape != target_shape:
            font_scale_factor = element.resize(
                target_shape[0], target_shape[1], api_client=client
            )
            client.flush_batch_update()

            # Re-fetch presentation to get updated table structure after resize
            presentation = Presentation.from_id(pres_id, api_client=client)
            slide = find_slide_by_name(presentation, slide_name)
            element = find_element_by_name(slide, element_name)

        # Generate and execute content update requests
        requests = element.content_update_requests(
            markdown_elem, check_shape=False, font_scale_factor=font_scale_factor
        )
        client.batch_update(requests, pres_id)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully wrote table to element '{element_name}'",
            details={
                "element_id": element.objectId,
                "slide_name": slide_name,
                "table_shape": list(target_shape),
                "resized": current_shape != target_shape,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error writing table markdown: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def bulk_write_element_markdown(
    presentation_id_or_url: str,
    writes: str,
) -> str:
    """Write markdown content to multiple shape elements in a single batch operation.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        writes: JSON string containing a list of write operations.
                Each entry: {"slide_name": str, "element_name": str, "markdown": str}
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    # Parse writes JSON
    try:
        write_list = json.loads(writes)
    except json.JSONDecodeError as e:
        return _format_response(
            None,
            validation_error("writes", f"Invalid JSON: {e}", writes[:200]),
        )

    if not isinstance(write_list, list):
        return _format_response(
            None,
            validation_error("writes", "Expected a JSON array of write operations", type(write_list).__name__),
        )

    # Validate each entry has required keys
    required_keys = {"slide_name", "element_name", "markdown"}
    for i, entry in enumerate(write_list):
        if not isinstance(entry, dict):
            return _format_response(
                None,
                validation_error("writes", f"Entry {i} is not an object", str(entry)[:200]),
            )
        missing = required_keys - set(entry.keys())
        if missing:
            return _format_response(
                None,
                validation_error("writes", f"Entry {i} missing keys: {missing}", str(entry)[:200]),
            )

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)

        # Cache slides by name for efficient lookup
        slides_by_name = {}
        for slide in presentation.slides:
            name = get_slide_name(slide)
            if name:
                slides_by_name[name] = slide

        successes = []
        failures = []

        for entry in write_list:
            slide_name = entry["slide_name"]
            element_name = entry["element_name"]
            markdown = entry["markdown"]

            try:
                slide = slides_by_name.get(slide_name)
                if slide is None:
                    failures.append({
                        "slide_name": slide_name,
                        "element_name": element_name,
                        "error": f"Slide '{slide_name}' not found",
                    })
                    continue

                element = find_element_by_name(slide, element_name)
                if element is None:
                    failures.append({
                        "slide_name": slide_name,
                        "element_name": element_name,
                        "error": f"Element '{element_name}' not found",
                    })
                    continue

                if not isinstance(element, ShapeElement):
                    failures.append({
                        "slide_name": slide_name,
                        "element_name": element_name,
                        "error": f"Element '{element_name}' is not a text element (type: {element.type.value})",
                    })
                    continue

                element.write_text(markdown, as_markdown=True, api_client=client)
                successes.append({
                    "slide_name": slide_name,
                    "element_name": element_name,
                    "element_id": element.objectId,
                })
            except Exception as entry_error:
                failures.append({
                    "slide_name": slide_name,
                    "element_name": element_name,
                    "error": str(entry_error),
                })

        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Bulk write completed: {len(successes)} succeeded, {len(failures)} failed",
            details={
                "total": len(write_list),
                "succeeded": len(successes),
                "failed": len(failures),
                "successes": successes,
                "failures": failures,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error in bulk write: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


# =============================================================================
# IMAGE TOOLS
# =============================================================================


@mcp.tool()
def replace_element_image(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str = None,
    image_source: str = "",
    element_id: str = None,
) -> str:
    """Replace an image element with a new image from a URL or local file path.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (image alt-title). Either this or element_id must be provided.
        image_source: URL (http/https) or local file path of the new image
        element_id: Element object ID (alternative to element_name, for unnamed elements)
    """
    if element_name is None and element_id is None:
        return _format_response(
            None,
            validation_error("element_name", "Either element_name or element_id must be provided", None),
        )

    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        # Find element by name or by ID
        element = None
        if element_id is not None:
            for el in slide.page_elements_flat:
                if el.objectId == element_id:
                    element = el
                    break
            if element is None:
                available = get_available_element_names(slide)
                client.flush_batch_update()
                return _format_response(
                    None,
                    validation_error(
                        "element_id",
                        f"No element found with ID '{element_id}' on slide '{slide_name}'",
                        element_id,
                    ),
                )
        else:
            element = find_element_by_name(slide, element_name)
            if element is None:
                available = get_available_element_names(slide)
                client.flush_batch_update()
                return _format_response(None, element_not_found_error(pres_id, slide_name, element_name, available))

        display_name = element_name or element_id

        # Check if it's an image element
        if not isinstance(element, ImageElement):
            client.flush_batch_update()
            return _format_response(
                None,
                validation_error(
                    "element_name",
                    f"Element '{display_name}' is not an image element (type: {element.type.value})",
                    display_name,
                ),
            )

        # Replace the image - route to url= or file= based on source
        if image_source.startswith(("http://", "https://")):
            element.replace_image(url=image_source, api_client=client)
        else:
            element.replace_image(file=image_source, api_client=client)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully replaced image in element '{display_name}'",
            details={
                "element_id": element.objectId,
                "slide_name": slide_name,
                "image_source": image_source,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error replacing image: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


# =============================================================================
# SLIDE MANIPULATION TOOLS
# =============================================================================


@mcp.tool()
def copy_slide(
    presentation_id_or_url: str,
    slide_name: str,
    insertion_index: int = None,
) -> str:
    """Duplicate a slide within the presentation.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name to copy
        insertion_index: Position for new slide (None = after original)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        # Duplicate the slide
        new_slide = slide.duplicate(api_client=client)

        # Move to specified position if provided
        if insertion_index is not None:
            new_slide.move(insertion_index, api_client=client)

        client.flush_batch_update()

        # Get the name of the new slide (will be same speaker notes initially)
        new_slide_name = get_slide_name(new_slide)

        result = SuccessResponse(
            message=f"Successfully copied slide '{slide_name}'",
            details={
                "original_slide_id": slide.objectId,
                "new_slide_id": new_slide.objectId,
                "new_slide_name": new_slide_name,
                "insertion_index": insertion_index,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error copying slide: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def move_slide(
    presentation_id_or_url: str,
    slide_name: str,
    insertion_index: int,
) -> str:
    """Move a slide to a new position in the presentation.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name to move
        insertion_index: New position (0-indexed)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        # Get current index for reporting
        current_index = presentation.slides.index(slide)

        # Move the slide
        slide.move(insertion_index, api_client=client)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully moved slide '{slide_name}' to position {insertion_index}",
            details={
                "slide_id": slide.objectId,
                "previous_index": current_index,
                "new_index": insertion_index,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error moving slide: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def delete_slide(
    presentation_id_or_url: str,
    slide_name: str,
) -> str:
    """Delete a slide from the presentation.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name to delete
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        presentation = Presentation.from_id(pres_id, api_client=client)
        slide = find_slide_by_name(presentation, slide_name)

        if slide is None:
            available = get_available_slide_names(presentation)
            client.flush_batch_update()
            return _format_response(None, slide_not_found_error(pres_id, slide_name, available))

        slide_id = slide.objectId

        # Delete the slide
        slide.delete(api_client=client)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully deleted slide '{slide_name}'",
            details={
                "deleted_slide_id": slide_id,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error deleting slide: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


# =============================================================================
# PRESENTATION MANIPULATION TOOLS
# =============================================================================


@mcp.tool()
def copy_presentation(
    presentation_id_or_url: str,
    copy_title: str = None,
    folder_id: str = None,
) -> str:
    """Copy an entire presentation to create a new one.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        copy_title: Title for the copy (defaults to "Copy of {original title}")
        folder_id: Google Drive folder ID to place the copy in (optional)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        # Load presentation to get its title for the default copy name
        presentation = Presentation.from_id(pres_id, api_client=client)
        original_title = presentation.title or "Untitled"

        if copy_title is None:
            copy_title = f"Copy of {original_title}"

        # Copy the presentation
        copy_result = client.copy_presentation(pres_id, copy_title, folder_id)
        new_pres_id = copy_result["id"]

        result = SuccessResponse(
            message=f"Successfully copied presentation '{original_title}'",
            details={
                "original_presentation_id": pres_id,
                "new_presentation_id": new_pres_id,
                "new_presentation_url": f"https://docs.google.com/presentation/d/{new_pres_id}/edit",
                "new_title": copy_title,
            },
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error copying presentation: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def add_element_names(
    presentation_id_or_url: str,
    skip_empty_text_boxes: bool = False,
    min_image_size_cm: float = 4.0,
) -> str:
    """Name all slides and elements in a presentation.

    Names slides based on their speaker notes (first line).
    Names elements (text boxes, images, charts, tables) with descriptive alt-text titles.
    The topmost text box becomes "Title", others become "Text_1", "Text_2", etc.
    Images and charts are named "Image_1", "Chart_1", etc.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        skip_empty_text_boxes: Skip text boxes that contain only whitespace
        min_image_size_cm: Minimum image dimension (cm) to include (smaller images are skipped)
    """
    try:
        pres_id = parse_presentation_id(presentation_id_or_url)
    except ValueError as e:
        return _format_response(None, validation_error("presentation_id_or_url", str(e), presentation_id_or_url))

    client = get_api_client()

    try:
        gslides_client = GSlidesAPIClient(gslides_client=client)
        slide_names = name_slides(
            pres_id,
            name_elements=True,
            api_client=gslides_client,
            skip_empty_text_boxes=skip_empty_text_boxes,
            min_image_size_cm=min_image_size_cm,
        )
        client.flush_batch_update()

        # Convert SlideElementNames dataclass to serializable dict
        names_dict = {}
        for slide_name, element_names in slide_names.items():
            names_dict[slide_name] = {
                "text_names": element_names.text_names,
                "image_names": element_names.image_names,
                "chart_names": element_names.chart_names,
                "table_names": element_names.table_names,
            }

        result = SuccessResponse(
            message=f"Successfully named {len(slide_names)} slides and their elements",
            details={"slide_element_names": names_dict},
        )
        return _format_response(result)

    except Exception as e:
        logger.error(f"Error naming elements: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================


def main():
    """Main entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="gslides-api MCP Server")
    parser.add_argument(
        "--credential-path",
        type=str,
        default=os.environ.get("GSLIDES_CREDENTIALS_PATH"),
        help="Path to Google API credentials directory (or set GSLIDES_CREDENTIALS_PATH env var)",
    )
    parser.add_argument(
        "--default-format",
        type=str,
        choices=["raw", "domain", "markdown"],
        default="markdown",
        help="Default output format for tools (default: markdown)",
    )

    args = parser.parse_args()

    if not args.credential_path:
        print(
            "Error: Credential path required. Use --credential-path or set GSLIDES_CREDENTIALS_PATH",
            file=sys.stderr,
        )
        sys.exit(1)

    # Initialize the server
    default_format = OutputFormat(args.default_format)
    initialize_server(args.credential_path, default_format)

    # Run the MCP server
    mcp.run()


if __name__ == "__main__":
    main()
