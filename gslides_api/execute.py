from typing import Dict, Any


import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from gslides_api.credentials import creds

# The functions in this file are the only interaction with the raw gslides API in this library


def slides_batch_update(requests: list, presentation_id: str) -> Dict[str, Any]:
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

    creds, _ = google.auth.default()
    # pylint: disable=maybe-no-member
    try:
        drive_service = build("drive", "v3", credentials=creds)
        body = {"name": copy_title}
        drive_response = drive_service.files().copy(fileId=presentation_id, body=body).execute()
        presentation_copy_id = drive_response.get("id")

    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Presentations  not copied")
        return error

    return presentation_copy_id
