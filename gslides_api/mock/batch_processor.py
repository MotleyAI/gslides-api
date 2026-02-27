"""Batch request processor for MockGoogleAPIClient.

Processes Google Slides API batch request dicts and applies mutations
to in-memory presentation state. Returns reply dicts matching the
real Google API response format.
"""

import copy
import logging
from typing import Any, Callable, Dict, List, Optional

logger = logging.getLogger(__name__)

# Request types that create objects and return objectId in the reply
_CREATE_REQUEST_TYPES = {
    "createSlide",
    "createShape",
    "createImage",
    "createTable",
    "createLine",
    "createVideo",
    "createSheetsChart",
}


def process_batch_requests(
    requests: List[List[Dict[str, Any]]],
    presentation_id: str,
    presentations: Dict[str, dict],
    generate_id: Callable[[], str],
) -> Dict[str, Any]:
    """Process a list of batch request dicts against in-memory state.

    Args:
        requests: List of request lists from GSlidesAPIRequest.to_request().
            Each inner list typically contains one dict like [{"createSlide": {...}}].
        presentation_id: The target presentation ID.
        presentations: The shared in-memory presentations store.
        generate_id: Callable that returns a new unique ID string.

    Returns:
        A dict matching Google Slides API batchUpdate response format:
        {"presentationId": "...", "replies": [{...}, ...]}
    """
    presentation = presentations.get(presentation_id)
    if presentation is None:
        raise KeyError(
            f"Presentation '{presentation_id}' not found in mock store"
        )

    replies = []
    for request_group in requests:
        for request_dict in request_group:
            request_type = list(request_dict.keys())[0]
            request_body = request_dict[request_type]
            reply = _process_single_request(
                request_type, request_body, presentation, generate_id
            )
            replies.append(reply)

    return {"presentationId": presentation_id, "replies": replies}


def _process_single_request(
    request_type: str,
    body: Dict[str, Any],
    presentation: dict,
    generate_id: Callable[[], str],
) -> Dict[str, Any]:
    """Process a single request and return the reply dict."""
    handler = _HANDLERS.get(request_type)
    if handler is not None:
        return handler(body, presentation, generate_id)

    # Passthrough: record but don't mutate state
    return {}


def _handle_create_slide(
    body: Dict[str, Any], presentation: dict, generate_id: Callable[[], str]
) -> Dict[str, Any]:
    object_id = body.get("objectId") or generate_id()
    notes_id = generate_id()
    speaker_notes_id = generate_id()

    slide = {
        "objectId": object_id,
        "pageElements": [],
        "slideProperties": {
            "layoutObjectId": body.get("slideLayoutReference", {}).get("layoutId"),
            "masterObjectId": None,
            "notesPage": {
                "objectId": notes_id,
                "pageElements": [],
                "notesProperties": {"speakerNotesObjectId": speaker_notes_id},
                "pageType": "NOTES",
            },
        },
        "pageType": "SLIDE",
    }

    slides = presentation.setdefault("slides", [])
    insertion_index = body.get("insertionIndex")
    if insertion_index is not None and insertion_index < len(slides):
        slides.insert(insertion_index, slide)
    else:
        slides.append(slide)

    return {"createSlide": {"objectId": object_id}}


def _handle_create_element(
    element_type: str,
    body: Dict[str, Any],
    presentation: dict,
    generate_id: Callable[[], str],
) -> Dict[str, Any]:
    """Generic handler for element creation requests (shape, image, table, etc.)."""
    object_id = body.get("objectId") or generate_id()

    element = {
        "objectId": object_id,
        "transform": body.get("elementProperties", {}).get("transform", {}),
        "size": body.get("elementProperties", {}).get("size", {}),
    }

    # Find the target page and add the element
    page_id = body.get("elementProperties", {}).get("pageObjectId")
    if page_id:
        page = _find_page(presentation, page_id)
        if page is not None:
            page.setdefault("pageElements", []).append(element)

    # The API key in the reply matches the request type
    return {element_type: {"objectId": object_id}}


def _handle_create_shape(body, presentation, generate_id):
    return _handle_create_element("createShape", body, presentation, generate_id)


def _handle_create_image(body, presentation, generate_id):
    return _handle_create_element("createImage", body, presentation, generate_id)


def _handle_create_table(body, presentation, generate_id):
    return _handle_create_element("createTable", body, presentation, generate_id)


def _handle_create_line(body, presentation, generate_id):
    return _handle_create_element("createLine", body, presentation, generate_id)


def _handle_create_video(body, presentation, generate_id):
    return _handle_create_element("createVideo", body, presentation, generate_id)


def _handle_duplicate_object(
    body: Dict[str, Any], presentation: dict, generate_id: Callable[[], str]
) -> Dict[str, Any]:
    source_id = body["objectId"]
    id_mapping = body.get("objectIds") or {}

    # Try to find as a slide first
    slides = presentation.get("slides", [])
    source_slide = next((s for s in slides if s.get("objectId") == source_id), None)

    if source_slide is not None:
        new_slide = copy.deepcopy(source_slide)
        new_id = id_mapping.get(source_id) or generate_id()
        new_slide["objectId"] = new_id

        # Remap element IDs within the duplicated slide
        for element in new_slide.get("pageElements", []):
            old_element_id = element.get("objectId")
            if old_element_id and old_element_id in id_mapping:
                element["objectId"] = id_mapping[old_element_id]
            elif old_element_id:
                element["objectId"] = generate_id()

        # Insert after the source slide
        source_index = slides.index(source_slide)
        slides.insert(source_index + 1, new_slide)
        return {"duplicateObject": {"objectId": new_id}}

    # Try to find as a page element
    for slide in slides:
        for element in slide.get("pageElements", []):
            if element.get("objectId") == source_id:
                new_element = copy.deepcopy(element)
                new_id = id_mapping.get(source_id) or generate_id()
                new_element["objectId"] = new_id
                slide["pageElements"].append(new_element)
                return {"duplicateObject": {"objectId": new_id}}

    # Object not found — still return a reply with a generated ID
    new_id = id_mapping.get(source_id) or generate_id()
    return {"duplicateObject": {"objectId": new_id}}


def _handle_delete_object(
    body: Dict[str, Any], presentation: dict, generate_id: Callable[[], str]
) -> Dict[str, Any]:
    target_id = body["objectId"]

    # Try removing as a slide
    slides = presentation.get("slides", [])
    presentation["slides"] = [s for s in slides if s.get("objectId") != target_id]

    # Try removing as a page element from all slides
    for slide in presentation.get("slides", []):
        elements = slide.get("pageElements", [])
        slide["pageElements"] = [
            e for e in elements if e.get("objectId") != target_id
        ]

    return {}


def _handle_update_slides_position(
    body: Dict[str, Any], presentation: dict, generate_id: Callable[[], str]
) -> Dict[str, Any]:
    slide_ids = set(body.get("slideObjectIds", []))
    insertion_index = body.get("insertionIndex", 0)

    slides = presentation.get("slides", [])
    moving = [s for s in slides if s.get("objectId") in slide_ids]
    remaining = [s for s in slides if s.get("objectId") not in slide_ids]

    # Clamp insertion index
    insertion_index = min(insertion_index, len(remaining))
    presentation["slides"] = (
        remaining[:insertion_index] + moving + remaining[insertion_index:]
    )
    return {}


def _find_page(presentation: dict, page_id: str) -> Optional[dict]:
    """Find a page (slide, layout, or master) by its objectId."""
    for page_list_key in ("slides", "layouts", "masters"):
        for page in presentation.get(page_list_key, []):
            if page.get("objectId") == page_id:
                return page
    return None


# Handler registry
_HANDLERS = {
    "createSlide": _handle_create_slide,
    "createShape": _handle_create_shape,
    "createImage": _handle_create_image,
    "createTable": _handle_create_table,
    "createLine": _handle_create_line,
    "createVideo": _handle_create_video,
    "duplicateObject": _handle_duplicate_object,
    "deleteObject": _handle_delete_object,
    "updateSlidesPosition": _handle_update_slides_position,
}
