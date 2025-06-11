from typing import Dict, Any, List

from gslides_api.execute import batch_update


def duplicate_object(object_id: str, presentation_id: str, id_map: Dict[str, str] = None) -> str:
    """Duplicates an object in a Google Slides presentation.
    When duplicating a slide, the duplicate slide will be created immediately following the specified slide.
    When duplicating a page element, the duplicate will be placed on the same page at the same position
    as the original.

    Args:
        object_id: The ID of the object to duplicate.
        presentation_id: The ID of the presentation containing the object.
        id_map: A dictionary mapping the IDs of the original objects to the IDs of the duplicated objects.

    Returns:
        The ID of the duplicated object.
    """
    # https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#DuplicateObjectRequest

    request = {"duplicateObject": {"objectId": object_id}}
    if id_map:
        request["duplicateObject"]["objectIds"] = id_map
    out = batch_update([request], presentation_id)
    new_object_id = out["replies"][0]["duplicateObject"]["objectId"]
    return new_object_id


def delete_object(object_id: str, presentation_id: str) -> None:
    """Deletes an object in a Google Slides presentation.

    Args:
        object_id: The ID of the object to delete.
        presentation_id: The ID of the presentation containing the object.
    """
    request = {"deleteObject": {"objectId": object_id}}
    batch_update([request], presentation_id)


def dict_to_dot_separated_field_list(x: Dict[str, Any]) -> List[str]:
    """Convert a dictionary to a list of dot-separated fields."""
    out = []
    for k, v in x.items():
        if isinstance(v, dict):
            out += [f"{k}.{i}" for i in dict_to_dot_separated_field_list(v)]
        else:
            out.append(k)
    return out


def image_url_is_valid(url: str) -> bool:
    """
    Validate that an image URL is accessible and valid.

    Args:
        url: Image URL to validate

    Returns:
        True if URL appears to be valid and accessible
    """
    import urllib.request
    import urllib.error

    if not url or not url.startswith(("http://", "https://")):
        return False

    # Check for common image extensions
    valid_extensions = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp")
    url_lower = url.lower()

    # Allow URLs with parameters that might contain image extensions
    if not any(ext in url_lower for ext in valid_extensions):
        # If no obvious image extension, try a quick HEAD request
        try:
            req = urllib.request.Request(url, method="HEAD")
            req.add_header("User-Agent", "Mozilla/5.0 (compatible; Google-Slides-Templater/1.0)")

            with urllib.request.urlopen(req, timeout=5) as response:
                content_type = response.headers.get("Content-Type", "")
                return content_type.startswith("image/")

        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError):
            logger.warning(f"Could not validate image URL: {url}")
            return False

    return True
