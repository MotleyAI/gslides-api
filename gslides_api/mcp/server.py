"""MCP server for gslides-api.

This module provides an MCP server that exposes Google Slides operations as tools.
"""

import argparse
import base64
import json
import logging
import os
import re
import sys
import tempfile
import traceback
from typing import Any, Dict, Optional

from mcp.server import FastMCP

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
from gslides_api.presentation import Presentation
from gslides_api.request.request import UpdateShapePropertiesRequest

from .models import (
    ErrorResponse,
    OutputFormat,
    PresentationOutline,
    SlideOutline,
    SuccessResponse,
    ThumbnailSizeOption,
)
from .utils import (
    build_element_outline,
    build_presentation_outline,
    build_slide_outline,
    element_not_found_error,
    find_element_by_name,
    find_slide_by_name,
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

# Global API client - initialized with auto_flush=False as per requirements
api_client: Optional[GoogleAPIClient] = None

# Default output format - can be overridden via CLI arg
DEFAULT_OUTPUT_FORMAT: OutputFormat = OutputFormat.RAW


def get_api_client() -> GoogleAPIClient:
    """Get the initialized API client."""
    if api_client is None:
        raise RuntimeError("API client not initialized. Call initialize_server() first.")
    return api_client


def initialize_server(credential_path: str, default_format: OutputFormat = OutputFormat.RAW):
    """Initialize the MCP server with credentials.

    Args:
        credential_path: Path to the Google API credentials directory
        default_format: Default output format for tools
    """
    global api_client, DEFAULT_OUTPUT_FORMAT

    # Create client with auto_flush=False
    api_client = GoogleAPIClient(auto_flush=False)

    # Initialize credentials on the api_client instance directly
    api_client.initialize_credentials(credential_path)

    # Set the global api_client in the gslides_api.client module
    import gslides_api.client

    gslides_api.client.api_client = api_client

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
        how: Output format - 'raw' (Google API JSON), 'domain' (model_dump), or 'outline' (condensed)
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

        elif format_type == OutputFormat.DOMAIN:
            # Get domain object and dump
            presentation = Presentation.from_id(pres_id, api_client=client)
            client.flush_batch_update()
            return _format_response(presentation.model_dump())

        else:  # OUTLINE
            presentation = Presentation.from_id(pres_id, api_client=client)
            client.flush_batch_update()
            outline = build_presentation_outline(presentation)
            return _format_response(outline)

    except Exception as e:
        logger.error(f"Error getting presentation: {e}\n{traceback.format_exc()}")
        return _format_response(None, presentation_error(pres_id, e))


@mcp.tool()
def get_slide(
    presentation_id_or_url: str,
    slide_name: str,
    how: str = None,
) -> str:
    """Get a single slide by name (first line of speaker notes).

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes, stripped)
        how: Output format - 'raw' (Google API JSON), 'domain' (model_dump), or 'outline' (condensed)
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

        if format_type == OutputFormat.RAW:
            result = client.get_slide_json(pres_id, slide.objectId)
            client.flush_batch_update()
            return _format_response(result)

        elif format_type == OutputFormat.DOMAIN:
            client.flush_batch_update()
            return _format_response(slide.model_dump())

        else:  # OUTLINE
            client.flush_batch_update()
            outline = build_slide_outline(slide)
            return _format_response(outline)

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
        how: Output format - 'raw' (Google API JSON), 'domain' (model_dump), or 'outline' (condensed)
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

        elif format_type == OutputFormat.DOMAIN:
            return _format_response(element.model_dump())

        else:  # OUTLINE
            outline = build_element_outline(element)
            return _format_response(outline)

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

        # Save thumbnail to temp file
        image_data = thumbnail.payload

        # Sanitize slide_name for filename (replace unsafe chars with underscore)
        safe_slide_name = re.sub(r'[^\w\-]', '_', slide_name)
        filename = f"{pres_id}_{safe_slide_name}_thumbnail.png"
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


# =============================================================================
# IMAGE TOOLS
# =============================================================================


@mcp.tool()
def replace_element_image(
    presentation_id_or_url: str,
    slide_name: str,
    element_name: str,
    image_url: str,
) -> str:
    """Replace an image element with a new image from URL.

    Args:
        presentation_id_or_url: Google Slides URL or presentation ID
        slide_name: Slide name (first line of speaker notes)
        element_name: Element name (image alt-title)
        image_url: URL of new image
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

        # Check if it's an image element
        if not isinstance(element, ImageElement):
            client.flush_batch_update()
            return _format_response(
                None,
                validation_error(
                    "element_name",
                    f"Element '{element_name}' is not an image element (type: {element.type.value})",
                    element_name,
                ),
            )

        # Replace the image
        element.replace_image(url=image_url, api_client=client)
        client.flush_batch_update()

        result = SuccessResponse(
            message=f"Successfully replaced image in element '{element_name}'",
            details={
                "element_id": element.objectId,
                "slide_name": slide_name,
                "new_image_url": image_url,
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
        choices=["raw", "domain", "outline"],
        default="raw",
        help="Default output format for tools (default: raw)",
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
