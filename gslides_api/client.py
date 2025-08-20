import logging
import os
import time
from functools import wraps
from typing import Any, Callable, Dict, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import Resource, build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from typeguard import typechecked

from gslides_api.domain import ThumbnailProperties
from gslides_api.request.request import (
    DeleteObjectRequest,
    DuplicateObjectRequest,
    GSlidesAPIRequest,
)
from gslides_api.response import ImageThumbnail


# The functions in this file are the only interaction with the raw gslides API in this library
@typechecked
class GoogleAPIClient:
    # Initial version from the gslides package
    """The credentials object to build the connections to the APIs"""

    def __init__(
        self, auto_flush: bool = True, initial_wait_s: int = 60, n_backoffs: int = 4
    ) -> None:
        """Constructor method

        Args:
            auto_flush: Whether to automatically flush batch requests
            initial_wait_s: Initial wait time in seconds for exponential backoff
            n_backoffs: Number of backoff attempts before giving up
        """
        self.crdtls: Optional[Credentials] = None
        self.sht_srvc: Optional[Resource] = None
        self.sld_srvc: Optional[Resource] = None
        self.drive_srvc: Optional[Resource] = None
        self.pending_batch_requests: list[GSlidesAPIRequest] = []
        self.pending_presentation_id: Optional[str] = None
        self.auto_flush = auto_flush
        self.initial_wait_s = initial_wait_s
        self.n_backoffs = n_backoffs

        # Create the exponential backoff decorator
        self._with_exponential_backoff = self._create_exponential_backoff_decorator()

    def _create_exponential_backoff_decorator(self) -> Callable:
        """Creates an exponential backoff decorator for API calls."""

        def decorator(func: Callable) -> Callable:
            @wraps(func)
            def wrapper(*args, **kwargs):
                last_exception = None
                for attempt in range(self.n_backoffs + 1):  # +1 for initial attempt
                    try:
                        return func(*args, **kwargs)
                    except HttpError as e:
                        # Check if it's a rate limit error (429) or server error (5xx)
                        if e.resp.status in [429, 500, 502, 503, 504]:
                            last_exception = e
                            if attempt < self.n_backoffs:  # Don't wait after the last attempt
                                wait_time = self.initial_wait_s * (2**attempt)
                                logger.warning(
                                    f"Rate limit/server error encountered (status {e.resp.status}), "
                                    f"waiting {wait_time}s before retry {attempt + 1}/{self.n_backoffs}"
                                )
                                time.sleep(wait_time)
                            continue
                        else:
                            # For other HTTP errors, don't retry
                            raise e
                    except Exception as e:
                        # For non-HTTP errors, don't retry
                        raise e

                # If we get here, all retries failed
                logger.error(f"All {self.n_backoffs} retry attempts failed")
                raise last_exception

            return wrapper

        return decorator

    def set_credentials(self, credentials: Optional[Credentials]) -> None:
        """Sets the credentials

        :param credentials: :class:`google.oauth2.credentials.Credentials`
        :type credentials: :class:`google.oauth2.credentials.Credentials`

        """
        self.crdtls = credentials
        logger.info("Building sheets connection")
        self.sht_srvc = build("sheets", "v4", credentials=credentials)
        logger.info("Built sheets connection")
        logger.info("Building slides connection")
        self.sld_srvc = build("slides", "v1", credentials=credentials)
        logger.info("Built slides connection")
        logger.info("Building drive connection")
        self.drive_srvc = build("drive", "v3", credentials=credentials)
        logger.info("Built drive connection")

    @property
    def sheet_service(self) -> Resource:
        """Returns the connects to the sheets API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sht_srvc:
            return self.sht_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def slide_service(self) -> Resource:
        """Returns the connects to the slides API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.sld_srvc:
            return self.sld_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    @property
    def drive_service(self) -> Resource:
        """Returns the connects to the drive API

        :raises RuntimeError: Must run set_credentials before executing method
        :return: API connection
        :rtype: :class:`googleapiclient.discovery.Resource`
        """
        if self.drive_srvc:
            return self.drive_srvc
        else:
            raise RuntimeError("Must run set_credentials before executing method")

    def flush_batch_update(self) -> Dict[str, Any]:
        if not len(self.pending_batch_requests):
            return {}

        re_requests = [r.to_request() for r in self.pending_batch_requests]

        @self._with_exponential_backoff
        def _execute_batch_update():
            return (
                self.slide_service.presentations()
                .batchUpdate(
                    presentationId=self.pending_presentation_id, body={"requests": re_requests}
                )
                .execute()
            )

        try:
            out = _execute_batch_update()
            self.pending_batch_requests = []
            self.pending_presentation_id = None
            return out
        except Exception as e:
            logger.error(f"Failed to execute batch update: {e}")
            raise e

    def batch_update(
        self, requests: list, presentation_id: str, flush: bool = False
    ) -> Dict[str, Any]:
        assert all(isinstance(r, GSlidesAPIRequest) for r in requests)

        if self.pending_presentation_id is None:
            self.pending_presentation_id = presentation_id
        elif self.pending_presentation_id != presentation_id:
            self.flush_batch_update()
            self.pending_presentation_id = presentation_id

        self.pending_batch_requests.extend(requests)

        if self.auto_flush or flush:
            return self.flush_batch_update()
        else:
            return {}

    # methods that call batch_update under the hood
    def duplicate_object(
        self,
        object_id: str,
        presentation_id: str,
        id_map: Dict[str, str] | None = None,
    ) -> str:
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
        request = DuplicateObjectRequest(objectId=object_id, objectIds=id_map)

        if id_map is not None and object_id in id_map:
            # the result object ID is known in advance, no need to flush
            self.batch_update([request], presentation_id, flush=False)
            return id_map[object_id]

        # Here we need to flush,
        out = self.batch_update([request], presentation_id, flush=True)
        # The new object ID is always the last one in the replies because we force-flushed
        new_object_id = out["replies"][-1]["duplicateObject"]["objectId"]
        return new_object_id

    def delete_object(self, object_id: str, presentation_id: str) -> None:
        """Deletes an object in a Google Slides presentation.

        Args:
            object_id: The ID of the object to delete.
            presentation_id: The ID of the presentation containing the object.
        """
        request = DeleteObjectRequest(objectId=object_id)
        self.batch_update([request], presentation_id, flush=False)

    # All methods that don't call batchUpdate under the hood must first flush any pending
    # batchUpdate calls, to preserve the correct order of operations

    def create_presentation(self, config: dict) -> str:
        self.flush_batch_update()
        # https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/create

        @self._with_exponential_backoff
        def _create():
            return self.slide_service.presentations().create(body=config).execute()

        out = _create()
        return out["presentationId"]

    def get_slide_json(self, presentation_id: str, slide_id: str) -> Dict[str, Any]:
        self.flush_batch_update()

        @self._with_exponential_backoff
        def _get():
            return (
                self.slide_service.presentations()
                .pages()
                .get(presentationId=presentation_id, pageObjectId=slide_id)
                .execute()
            )

        return _get()

    def get_presentation_json(self, presentation_id: str) -> Dict[str, Any]:
        self.flush_batch_update()

        @self._with_exponential_backoff
        def _get():
            return self.slide_service.presentations().get(presentationId=presentation_id).execute()

        return _get()

    # TODO: test this out and adjust the credentials readme (Drive API scope, anything else?)
    # https://developers.google.com/workspace/slides/api/guides/presentations#python
    def copy_presentation(self, presentation_id, copy_title, folder_id=None):
        """
        Creates the copy Presentation the user has access to.
        Load pre-authorized user credentials from the environment.
        TODO(developer) - See https://developers.google.com/identity
        for guides on implementing OAuth2 for the application.

        :param presentation_id: The ID of the presentation to copy
        :param copy_title: The title for the copied presentation
        :param folder_id: Optional folder ID to place the copy in. If None, copies to the root directory.
        :return: The response from the Drive API copy operation
        """
        self.flush_batch_update()

        body = {"name": copy_title}
        if folder_id:
            body["parents"] = [folder_id]

        @self._with_exponential_backoff
        def _copy():
            return self.drive_service.files().copy(fileId=presentation_id, body=body).execute()

        return _copy()

    def create_folder(self, folder_name, parent_folder_id=None, ignore_existing=False):
        """
        Creates a new folder in Google Drive.

        :param folder_name: The name of the folder to create
        :param parent_folder_id: Optional parent folder ID. If None, creates folder in the root directory.
        :param ignore_existing: If True, returns existing folder if one with the same name exists in the parent directory
        :return: The response from the Drive API create operation containing the new folder's ID
        """
        self.flush_batch_update()

        if ignore_existing:
            # Search for existing folder with the same name in the parent directory
            query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
            if parent_folder_id:
                query += f" and '{parent_folder_id}' in parents"
            else:
                query += " and 'root' in parents"

            @self._with_exponential_backoff
            def _list_folders():
                return self.drive_service.files().list(q=query, fields="files(id,name)").execute()

            existing_folders = _list_folders()

            if existing_folders.get("files"):
                # Return the first matching folder
                return existing_folders["files"][0]

        body = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"}
        if parent_folder_id:
            body["parents"] = [parent_folder_id]

        @self._with_exponential_backoff
        def _create_folder():
            return self.drive_service.files().create(body=body, fields="id,name").execute()

        return _create_folder()

    def delete_file(self, file_id):
        """
        Deletes a file from Google Drive.

        This method can delete any type of file including presentations, folders, images, etc.
        The file is moved to the trash and can be restored from there unless permanently deleted.

        :param file_id: The ID of the file to delete
        :return: Empty response if successful
        :raises: Exception if the file doesn't exist or user doesn't have permission to delete
        """
        self.flush_batch_update()

        @self._with_exponential_backoff
        def _delete():
            return self.drive_service.files().delete(fileId=file_id).execute()

        return _delete()

    def upload_image_to_drive(self, image_path) -> str:
        """
        Uploads an image to Google Drive and returns the public URL.

        Supports PNG, JPEG, and GIF image formats. The image type is automatically
        detected from the file extension.

        :param image_path: Path to the image file
        :return: Public URL of the uploaded image
        :raises ValueError: If the image format is not supported (not PNG, JPEG, or GIF)
        """
        # Don't call flush_batch_update here, as image upload doesn't interact with the slide deck
        # Define supported image formats and their MIME types
        supported_formats = {
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
        }

        # Extract file extension and convert to lowercase
        file_extension = os.path.splitext(image_path)[1].lower()

        # Check if the format is supported
        if file_extension not in supported_formats:
            supported_exts = ", ".join(supported_formats.keys())
            raise ValueError(
                f"Unsupported image format '{file_extension}'. "
                f"Supported formats are: {supported_exts}"
            )

        # Get the appropriate MIME type
        mime_type = supported_formats[file_extension]

        file_metadata = {"name": os.path.basename(image_path), "mimeType": mime_type}
        media = MediaFileUpload(image_path, mimetype=mime_type)

        @self._with_exponential_backoff
        def _upload():
            return (
                self.drive_service.files()
                .create(body=file_metadata, media_body=media, fields="id")
                .execute()
            )

        uploaded = _upload()

        @self._with_exponential_backoff
        def _set_permissions():
            return (
                self.drive_service.permissions()
                .create(
                    fileId=uploaded["id"],
                    # TODO: do we need "anyone"?
                    body={"type": "anyone", "role": "reader"},
                )
                .execute()
            )

        _set_permissions()

        return f"https://drive.google.com/uc?id={uploaded['id']}"

    def slide_thumbnail(
        self, presentation_id: str, slide_id: str, properties: ThumbnailProperties
    ) -> ImageThumbnail:
        """Gets a thumbnail of the specified slide by calling
        https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail
        :param presentation_id: The ID of the presentation containing the slide
        :type presentation_id: str
        :param slide_id: The ID of the slide to get thumbnail for
        :type slide_id: str
        :param properties: Properties controlling thumbnail generation
        :type properties: ThumbnailProperties
        :return: Image response with thumbnail URL and dimensions
        :rtype: ImageResponse
        """
        self.flush_batch_update()

        @self._with_exponential_backoff
        def _get_thumbnail():
            return (
                self.slide_service.presentations()
                .pages()
                .getThumbnail(
                    presentationId=presentation_id,
                    pageObjectId=slide_id,
                    thumbnailProperties_mimeType=(
                        properties.mimeType.value if properties.mimeType else None
                    ),
                    thumbnailProperties_thumbnailSize=(
                        properties.thumbnailSize.value if properties.thumbnailSize else None
                    ),
                )
                .execute()
            )

        img_info = _get_thumbnail()
        return ImageThumbnail.model_validate(img_info)


api_client = GoogleAPIClient()

logger = logging.getLogger(__name__)


def initialize_credentials(credential_location: str):
    """

    :param credential_location:
    :return:
    """

    SCOPES = [
        "https://www.googleapis.com/auth/presentations",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    _creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(os.path.join(credential_location, "token.json")):
        _creds = Credentials.from_authorized_user_file(
            os.path.join(credential_location, "token.json"), SCOPES
        )
    # If there are no (valid) credentials available, let the user log in.
    if not _creds or not _creds.valid:
        if _creds and _creds.expired and _creds.refresh_token:
            _creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(credential_location, "credentials.json"), SCOPES
            )
            _creds = flow.run_local_server()
        # Save the credentials for the next run
        with open(os.path.join(credential_location, "token.json"), "w") as token:
            token.write(_creds.to_json())
    api_client.set_credentials(_creds)
