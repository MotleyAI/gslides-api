import json

from fastapi import HTTPException


class GoogleSlidesExportError(Exception):
    """Custom exception for Google Slides export errors."""

    pass


class GoogleSlidesAuthError(GoogleSlidesExportError):
    """Authentication-related errors during Google Slides export."""

    pass


class GoogleSlidesConnectionError(GoogleSlidesExportError, HTTPException):
    """Connection/integration not found errors."""

    def __init__(self, message: str, status_code: int = 400):
        super().__init__(status_code=status_code, detail=message)


class GoogleOAuthExpiredException(GoogleSlidesAuthError, HTTPException):
    """Raised when Google OAuth tokens have expired and user needs to re-authenticate."""

    def __init__(self, redirect_url: str | None = None):
        HTTPException.__init__(
            self,
            status_code=401,
            detail={
                "error": "google_oauth_expired",
                "message": "Your Google connection has expired. Please reconnect your Google account.",
                "redirect_url": redirect_url,
            },
        )


class GoogleDriveFileAccessDeniedException(GoogleSlidesExportError, HTTPException):
    """Raised when access to a specific Google Drive file is denied (appNotAuthorizedToFile).

    This typically happens when using drive.file scope and the user hasn't granted
    access to the specific file via the Google Picker.
    """

    def __init__(self, file_id: str, message: str | None = None):
        self.file_id = file_id
        HTTPException.__init__(
            self,
            status_code=403,
            detail={
                "error_type": "google_drive_file_access_denied",
                "file_id": file_id,
                "message": message or f"Access denied to Google Drive file: {file_id}",
                "requires_picker": True,
            },
        )


def detect_file_access_denied_error(error: Exception, file_id: str) -> None:
    """Check if an error is a Google Drive file access denied error and raise appropriate exception.

    This detects errors that occur when using drive.file scope and the app doesn't have access
    to a specific file. Google Drive API returns:
    - 404 "File not found" when the app can't see the file (most common with drive.file scope)
    - 403 "appNotAuthorizedToFile" in some cases

    Args:
        error: The exception to check
        file_id: The Google Drive file ID that was being accessed

    Raises:
        GoogleDriveFileAccessDeniedException: If the error indicates file access was denied
    """
    error_str = str(error).lower()

    # Check for common 404 file not found patterns (drive.file scope returns 404 for inaccessible files)
    if "404" in error_str and (
        "file not found" in error_str
        or "not found" in error_str
        or "requested entity was not found" in error_str
    ):
        raise GoogleDriveFileAccessDeniedException(
            file_id=file_id,
            message=f"File not accessible: {file_id}. Please grant access using the file picker.",
        )

    # Check for common 403 file access denied patterns
    if "403" in error_str and (
        "appnotauthorizedtofile" in error_str
        or "has not granted the app" in error_str
        or "access to the file" in error_str
    ):
        raise GoogleDriveFileAccessDeniedException(
            file_id=file_id,
            message=f"Access denied to file {file_id}. Please grant access using the file picker.",
        )

    # Handle googleapiclient HttpError objects
    if hasattr(error, "resp") and hasattr(error, "content"):
        resp_status = getattr(error.resp, "status", None)
        if resp_status in (403, 404):
            try:
                error_content = error.content
                if isinstance(error_content, bytes):
                    error_content = error_content.decode("utf-8")
                error_data = json.loads(error_content)
                errors = error_data.get("error", {}).get("errors", [])
                for err in errors:
                    reason = err.get("reason", "")
                    # 404 with notFound reason or 403 with appNotAuthorizedToFile
                    if reason in ("notFound", "appNotAuthorizedToFile"):
                        raise GoogleDriveFileAccessDeniedException(
                            file_id=file_id,
                            message=err.get("message"),
                        )
                # If 404 without specific reason, still treat as access denied
                if resp_status == 404:
                    raise GoogleDriveFileAccessDeniedException(
                        file_id=file_id,
                        message=f"File not accessible: {file_id}. Please grant access using the file picker.",
                    )
            except GoogleDriveFileAccessDeniedException:
                raise
            except (json.JSONDecodeError, KeyError, UnicodeDecodeError):
                # If we can't parse but got 404, treat as access denied
                if resp_status == 404:
                    raise GoogleDriveFileAccessDeniedException(
                        file_id=file_id,
                        message=f"File not accessible: {file_id}. Please grant access using the file picker.",
                    )

    # Handle httpx response errors
    if hasattr(error, "response"):
        response = getattr(error, "response", None)
        if response is not None and hasattr(response, "status_code"):
            if response.status_code in (403, 404):
                try:
                    error_data = response.json()
                    errors = error_data.get("error", {}).get("errors", [])
                    for err in errors:
                        reason = err.get("reason", "")
                        if reason in ("notFound", "appNotAuthorizedToFile"):
                            raise GoogleDriveFileAccessDeniedException(
                                file_id=file_id,
                                message=err.get("message"),
                            )
                    # If 404 without specific reason, still treat as access denied
                    if response.status_code == 404:
                        raise GoogleDriveFileAccessDeniedException(
                            file_id=file_id,
                            message=f"File not accessible: {file_id}. Please grant access using the file picker.",
                        )
                except GoogleDriveFileAccessDeniedException:
                    raise
                except (json.JSONDecodeError, KeyError, AttributeError):
                    # If we can't parse but got 404, treat as access denied
                    if response.status_code == 404:
                        raise GoogleDriveFileAccessDeniedException(
                            file_id=file_id,
                            message=f"File not accessible: {file_id}. Please grant access using the file picker.",
                        )
