from typing import Dict, Any, List

from gslides_api.execute import slides_batch_update


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
    out = slides_batch_update([request], presentation_id)
    new_object_id = out["replies"][0]["duplicateObject"]["objectId"]
    return new_object_id


def delete_object(object_id: str, presentation_id: str) -> None:
    """Deletes an object in a Google Slides presentation.

    Args:
        object_id: The ID of the object to delete.
        presentation_id: The ID of the presentation containing the object.
    """
    request = {"deleteObject": {"objectId": object_id}}
    slides_batch_update([request], presentation_id)


def dict_to_dot_separated_field_list(x: Dict[str, Any]) -> List[str]:
    """Convert a dictionary to a list of dot-separated fields."""
    out = []
    for k, v in x.items():
        if isinstance(v, dict):
            out += [f"{k}.{i}" for i in dict_to_dot_separated_field_list(v)]
        else:
            out.append(k)
    return out
