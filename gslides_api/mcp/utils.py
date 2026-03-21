"""Utility functions for the gslides-api MCP server."""

import re
from typing import List, Optional

from gslides_api.adapters.abstract_slides import AbstractPresentation, AbstractSlide
from gslides_api.element.base import ElementKind, PageElementBase
from gslides_api.element.element import PageElement
from gslides_api.page.slide import Slide
from gslides_api.presentation import Presentation

from .models import ErrorResponse

# Pattern to match Google Slides URLs and extract the presentation ID
# Matches: https://docs.google.com/presentation/d/{ID}/edit
#          https://docs.google.com/presentation/d/{ID}
#          https://docs.google.com/presentation/d/{ID}/edit#slide=id.p
GOOGLE_SLIDES_URL_PATTERN = re.compile(
    r"^https?://docs\.google\.com/presentation/d/([a-zA-Z0-9_-]+)(?:/[^?#]*)?(?:\?[^#]*)?(?:#.*)?$"
)


def parse_presentation_id(url_or_id: str) -> str:
    """Extract presentation ID from a Google Slides URL or return the ID as-is.

    Args:
        url_or_id: Either a full Google Slides URL or a presentation ID

    Returns:
        The presentation ID

    Raises:
        ValueError: If the input looks like a URL but doesn't match the expected pattern
    """
    url_or_id = url_or_id.strip()

    # Check if it looks like a URL
    if url_or_id.startswith("http://") or url_or_id.startswith("https://"):
        match = GOOGLE_SLIDES_URL_PATTERN.match(url_or_id)
        if match:
            return match.group(1)
        else:
            raise ValueError(
                f"Invalid Google Slides URL format: {url_or_id}. "
                "Expected format: https://docs.google.com/presentation/d/{ID}/edit"
            )

    # Assume it's already a presentation ID
    return url_or_id


def get_slide_name(slide: Slide) -> Optional[str]:
    """Get the name of a slide from its speaker notes.

    The slide name is the first line of the speaker notes, stripped.

    Args:
        slide: The slide to get the name from

    Returns:
        The slide name, or None if no speaker notes
    """
    try:
        speaker_notes = slide.speaker_notes
        if speaker_notes is None:
            return None
        text = speaker_notes.read_text()
        if not text:
            return None
        text = text.strip()
        if not text:
            return None
        # First line only
        first_line = text.split("\n")[0].strip()
        return first_line if first_line else None
    except Exception:
        return None


def get_element_name(element: PageElementBase) -> Optional[str]:
    """Get the name of an element from its alt-text title.

    Args:
        element: The element to get the name from

    Returns:
        The element name, or None if no alt-text title
    """
    if hasattr(element, "title") and element.title:
        return element.title.strip() or None
    return None


def find_slide_by_name(presentation: Presentation, slide_name: str) -> Optional[Slide]:
    """Find a slide by its name (first line of speaker notes).

    Args:
        presentation: The presentation to search in
        slide_name: The slide name to find

    Returns:
        The slide if found, None otherwise
    """
    for slide in presentation.slides:
        name = get_slide_name(slide)
        if name == slide_name:
            return slide
    return None


def find_element_by_name(
    slide: Slide, element_name: str
) -> Optional[PageElement]:
    """Find an element on a slide by its name (alt-text title).

    Args:
        slide: The slide to search in
        element_name: The element name to find

    Returns:
        The element if found, None otherwise
    """
    for element in slide.page_elements_flat:
        name = get_element_name(element)
        if name == element_name:
            return element
    return None


def get_available_slide_names(presentation: Presentation) -> List[str]:
    """Get a list of all slide names in a presentation.

    Args:
        presentation: The presentation to get slide names from

    Returns:
        List of slide names (or slide IDs for unnamed slides)
    """
    names = []
    for i, slide in enumerate(presentation.slides):
        name = get_slide_name(slide)
        if name:
            names.append(name)
        else:
            names.append(f"(unnamed slide at index {i}, id: {slide.objectId})")
    return names


def get_available_element_names(slide: Slide) -> List[str]:
    """Get a list of all element names on a slide.

    Args:
        slide: The slide to get element names from

    Returns:
        List of element names (or element IDs for unnamed elements)
    """
    names = []
    for element in slide.page_elements_flat:
        name = get_element_name(element)
        if name:
            names.append(name)
        else:
            element_type = element.type.value if hasattr(element, "type") else "unknown"
            names.append(f"(unnamed {element_type}, id: {element.objectId})")
    return names


def get_element_type_string(element: PageElement) -> str:
    """Get a string representation of the element type.

    Args:
        element: The element to get the type from

    Returns:
        String representation of the element type
    """
    if hasattr(element, "type") and isinstance(element.type, ElementKind):
        return element.type.value
    return "unknown"


def create_error_response(
    error_type: str,
    message: str,
    **details,
) -> ErrorResponse:
    """Create a standardized error response.

    Args:
        error_type: The type of error
        message: Human-readable error message
        **details: Additional context to include

    Returns:
        ErrorResponse instance
    """
    return ErrorResponse(
        error_type=error_type,
        message=message,
        details=details,
    )


def slide_not_found_error(
    presentation_id: str, slide_name: str, available_slides: List[str]
) -> ErrorResponse:
    """Create a slide not found error response.

    Args:
        presentation_id: The presentation ID
        slide_name: The slide name that was not found
        available_slides: List of available slide names

    Returns:
        ErrorResponse for slide not found
    """
    return create_error_response(
        error_type="SlideNotFound",
        message=f"No slide found with name '{slide_name}'",
        presentation_id=presentation_id,
        searched_slide_name=slide_name,
        available_slides=available_slides,
    )


def element_not_found_error(
    presentation_id: str,
    slide_name: str,
    element_name: str,
    available_elements: List[str],
) -> ErrorResponse:
    """Create an element not found error response.

    Args:
        presentation_id: The presentation ID
        slide_name: The slide name
        element_name: The element name that was not found
        available_elements: List of available element names

    Returns:
        ErrorResponse for element not found
    """
    return create_error_response(
        error_type="ElementNotFound",
        message=f"No element found with name '{element_name}' on slide '{slide_name}'",
        presentation_id=presentation_id,
        slide_name=slide_name,
        searched_element_name=element_name,
        available_elements=available_elements,
    )


def presentation_error(presentation_id: str, error: Exception) -> ErrorResponse:
    """Create a presentation access error response.

    Args:
        presentation_id: The presentation ID
        error: The exception that occurred

    Returns:
        ErrorResponse for presentation access error
    """
    error_type = type(error).__name__
    return create_error_response(
        error_type=f"PresentationError:{error_type}",
        message=f"Failed to access presentation: {str(error)}",
        presentation_id=presentation_id,
        exception_type=error_type,
        exception_message=str(error),
    )


def validation_error(field: str, message: str, value: str = None) -> ErrorResponse:
    """Create a validation error response.

    Args:
        field: The field that failed validation
        message: Description of the validation failure
        value: The invalid value (optional)

    Returns:
        ErrorResponse for validation error
    """
    details = {"field": field}
    if value is not None:
        details["invalid_value"] = value
    return create_error_response(
        error_type="ValidationError",
        message=message,
        **details,
    )


def find_abstract_slide_by_name(
    presentation: AbstractPresentation, slide_name: str
) -> Optional[AbstractSlide]:
    """Find an AbstractSlide by speaker notes name.

    Args:
        presentation: The abstract presentation to search in
        slide_name: The slide name to find (first line of speaker notes)

    Returns:
        The abstract slide if found, None otherwise
    """
    for slide in presentation.slides:
        if slide.speaker_notes:
            text = slide.speaker_notes.read_text()
            if text:
                first_line = text.strip().split("\n")[0].strip()
                if first_line == slide_name:
                    return slide
    return None


def get_abstract_slide_names(presentation: AbstractPresentation) -> List[str]:
    """Get slide names from an AbstractPresentation for error messages.

    Args:
        presentation: The abstract presentation to get slide names from

    Returns:
        List of slide names (or placeholder for unnamed slides)
    """
    names = []
    for i, slide in enumerate(presentation.slides):
        name = None
        if slide.speaker_notes:
            text = slide.speaker_notes.read_text()
            if text:
                name = text.strip().split("\n")[0].strip() or None
        names.append(name or f"(unnamed slide at index {i})")
    return names
