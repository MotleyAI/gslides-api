from typing import Dict, Any


import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from gslides_api.credentials import creds

# The functions in this file are the only interaction with the raw gslides API in this library


def batch_update(requests: list, presentation_id: str) -> Dict[str, Any]:
    return (
        creds.slide_service.presentations()
        .batchUpdate(presentationId=presentation_id, body={"requests": requests})
        .execute()
    )


def create_presentation(config: dict) -> str:
    # https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/create
    out = creds.slide_service.presentations().create(body=config).execute()
    return out["presentationId"]


def get_slide_json(presentation_id: str, slide_id: str) -> Dict[str, Any]:
    return (
        creds.slide_service.presentations()
        .pages()
        .get(presentationId=presentation_id, pageObjectId=slide_id)
        .execute()
    )


def get_presentation_json(presentation_id: str) -> Dict[str, Any]:
    return creds.slide_service.presentations().get(presentationId=presentation_id).execute()


# TODO: test this out and adjust the credentials readme (Drive API scope, anything else?)
# https://developers.google.com/workspace/slides/api/guides/presentations#python
def copy_presentation(presentation_id, copy_title):
    """
    Creates the copy Presentation the user has access to.
    Load pre-authorized user credentials from the environment.
    TODO(developer) - See https://developers.google.com/identity
    for guides on implementing OAuth2 for the application.
    """

    return (
        creds.drive_service.files()
        .copy(fileId=presentation_id, body={"name": copy_title})
        .execute()
    )


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
