"""Mock Google API client for integration testing without real credentials.

Provides MockGoogleAPIClient, a subclass of GoogleAPIClient that stores
presentations and Drive files in memory. Drop-in replacement wherever
GoogleAPIClient is accepted.

Usage:
    from gslides_api.mock import MockGoogleAPIClient
    from gslides_api import Presentation

    mock = MockGoogleAPIClient()
    pres = Presentation.create_blank("Test", api_client=mock)
    loaded = Presentation.from_id(pres.presentationId, api_client=mock)
"""

import copy
import logging
import os
from typing import Any, Dict, Optional

from gslides_api.client import GoogleAPIClient
from gslides_api.domain.domain import ThumbnailProperties
from gslides_api.mock.batch_processor import process_batch_requests
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.request import DeleteObjectRequest, DuplicateObjectRequest
from gslides_api.response import ImageThumbnail

logger = logging.getLogger(__name__)

# Default page size matching Google Slides default (widescreen 16:9)
_DEFAULT_PAGE_SIZE = {
    "width": {"magnitude": 9144000, "unit": "EMU"},
    "height": {"magnitude": 5143500, "unit": "EMU"},
}


class MockGoogleAPIClient(GoogleAPIClient):
    """In-memory mock of GoogleAPIClient for testing.

    Stores presentations and Drive files in memory. Supports all
    GoogleAPIClient methods without requiring Google credentials.

    The batch system (batch_update / flush_batch_update) works identically
    to the real client — requests are queued and flushed — but flush
    processes them against in-memory state instead of calling Google.

    Extra test helpers:
        seed_presentation(id, json_dict) — preload a presentation snapshot
        get_batch_log() — inspect all processed batch requests
    """

    def __init__(
        self,
        auto_flush: bool = True,
        *,
        _shared_state: Optional[dict] = None,
    ) -> None:
        # Initialize parent's batch state and attributes.
        # Pass n_backoffs=0 so the backoff decorator is a no-op if somehow invoked.
        super().__init__(auto_flush=auto_flush, initial_wait_s=0, n_backoffs=0)

        if _shared_state is not None:
            # Child client: share stores with parent
            self._state = _shared_state
        else:
            self._state = {
                "presentations": {},  # id -> presentation JSON dict
                "files": {},  # id -> file metadata dict
                "id_counter": 0,
                "batch_log": [],  # list of (presentation_id, request_dicts)
            }

    def _generate_id(self) -> str:
        self._state["id_counter"] += 1
        return f"mock_id_{self._state['id_counter']}"

    # ── Properties ──────────────────────────────────────────────────────

    @property
    def is_initialized(self) -> bool:
        return True

    # ── Credential methods (no-ops) ─────────────────────────────────────

    def set_credentials(self, credentials=None) -> None:
        pass

    def initialize_credentials(self, credential_location: str) -> None:
        pass

    # ── Child client ────────────────────────────────────────────────────

    def create_child_client(self, auto_flush: bool = False) -> "MockGoogleAPIClient":
        return MockGoogleAPIClient(
            auto_flush=auto_flush,
            _shared_state=self._state,
        )

    # ── Batch system ────────────────────────────────────────────────────
    # batch_update() is inherited from GoogleAPIClient — it queues
    # requests and calls flush_batch_update() when appropriate.

    def flush_batch_update(self) -> Dict[str, Any]:
        if not self.pending_batch_requests:
            return {}

        re_requests = [r.to_request() for r in self.pending_batch_requests]
        presentation_id = self.pending_presentation_id

        # Log for test assertions
        self._state["batch_log"].append(
            {"presentation_id": presentation_id, "requests": re_requests}
        )

        result = process_batch_requests(
            requests=re_requests,
            presentation_id=presentation_id,
            presentations=self._state["presentations"],
            generate_id=self._generate_id,
        )

        self.pending_batch_requests = []
        self.pending_presentation_id = None
        return result

    # ── Slides API methods ──────────────────────────────────────────────

    def create_presentation(self, config: dict) -> str:
        self.flush_batch_update()
        pres_id = self._generate_id()

        # Build a minimal but valid presentation JSON
        default_slide_id = self._generate_id()
        notes_id = self._generate_id()
        speaker_notes_id = self._generate_id()

        presentation = {
            "presentationId": pres_id,
            "title": config.get("title", "Untitled"),
            "pageSize": config.get("pageSize", copy.deepcopy(_DEFAULT_PAGE_SIZE)),
            "slides": config.get("slides", [
                {
                    "objectId": default_slide_id,
                    "pageElements": [],
                    "slideProperties": {
                        "layoutObjectId": None,
                        "masterObjectId": None,
                        "notesPage": {
                            "objectId": notes_id,
                            "pageElements": [],
                            "notesProperties": {
                                "speakerNotesObjectId": speaker_notes_id
                            },
                            "pageType": "NOTES",
                        },
                    },
                    "pageType": "SLIDE",
                }
            ]),
            "layouts": config.get("layouts", []),
            "masters": config.get("masters", []),
        }

        self._state["presentations"][pres_id] = presentation
        return pres_id

    def get_presentation_json(self, presentation_id: str) -> Dict[str, Any]:
        self.flush_batch_update()
        presentation = self._state["presentations"].get(presentation_id)
        if presentation is None:
            raise KeyError(
                f"Presentation '{presentation_id}' not found in mock store"
            )
        return copy.deepcopy(presentation)

    def get_slide_json(self, presentation_id: str, slide_id: str) -> Dict[str, Any]:
        self.flush_batch_update()
        presentation = self._state["presentations"].get(presentation_id)
        if presentation is None:
            raise KeyError(
                f"Presentation '{presentation_id}' not found in mock store"
            )

        for slide in presentation.get("slides", []):
            if slide.get("objectId") == slide_id:
                return copy.deepcopy(slide)

        raise KeyError(
            f"Slide '{slide_id}' not found in presentation '{presentation_id}'"
        )

    def duplicate_object(
        self,
        object_id: str,
        presentation_id: str,
        id_map: Dict[str, str] | None = None,
    ) -> str:
        request = DuplicateObjectRequest(objectId=object_id, objectIds=id_map)

        if id_map is not None and object_id in id_map:
            self.batch_update([request], presentation_id, flush=False)
            return id_map[object_id]

        out = self.batch_update([request], presentation_id, flush=True)
        new_object_id = out["replies"][-1]["duplicateObject"]["objectId"]
        return new_object_id

    def delete_object(self, object_id: str, presentation_id: str) -> None:
        request = DeleteObjectRequest(objectId=object_id)
        self.batch_update([request], presentation_id, flush=False)

    def slide_thumbnail(
        self, presentation_id: str, slide_id: str, properties: ThumbnailProperties
    ) -> ImageThumbnail:
        self.flush_batch_update()
        # Verify the presentation and slide exist
        self.get_slide_json(presentation_id, slide_id)
        return ImageThumbnail(
            contentUrl=f"https://mock.test/thumbnail/{presentation_id}/{slide_id}",
            width=1600,
            height=900,
        )

    # ── Drive API methods ───────────────────────────────────────────────

    def copy_presentation(self, presentation_id, copy_title, folder_id=None):
        self.flush_batch_update()
        source = self._state["presentations"].get(presentation_id)
        if source is None:
            raise KeyError(
                f"Presentation '{presentation_id}' not found in mock store"
            )

        new_id = self._generate_id()
        new_pres = copy.deepcopy(source)
        new_pres["presentationId"] = new_id
        new_pres["title"] = copy_title
        self._state["presentations"][new_id] = new_pres

        file_meta = {"id": new_id, "name": copy_title}
        if folder_id:
            file_meta["parents"] = [folder_id]
        self._state["files"][new_id] = file_meta

        return file_meta

    def create_folder(self, folder_name, parent_folder_id=None, ignore_existing=False):
        self.flush_batch_update()

        if ignore_existing:
            for file_meta in self._state["files"].values():
                if (
                    file_meta.get("name") == folder_name
                    and file_meta.get("mimeType") == "application/vnd.google-apps.folder"
                ):
                    parent_match = (
                        parent_folder_id is None
                        or parent_folder_id in file_meta.get("parents", [])
                    )
                    if parent_match:
                        return {"id": file_meta["id"], "name": file_meta["name"]}

        folder_id = self._generate_id()
        file_meta = {
            "id": folder_id,
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
        }
        if parent_folder_id:
            file_meta["parents"] = [parent_folder_id]
        self._state["files"][folder_id] = file_meta
        return {"id": folder_id, "name": folder_name}

    def delete_file(self, file_id):
        self.flush_batch_update()
        self._state["files"].pop(file_id, None)
        self._state["presentations"].pop(file_id, None)
        return {}

    def upload_image_to_drive(self, image_path, folder_id: str | None = None) -> str:
        supported_formats = {".png", ".jpg", ".jpeg", ".gif"}
        ext = os.path.splitext(image_path)[1].lower()
        if ext not in supported_formats:
            raise ValueError(
                f"Unsupported image format '{ext}'. "
                f"Supported formats are: {', '.join(supported_formats)}"
            )

        file_id = self._generate_id()
        file_meta = {
            "id": file_id,
            "name": os.path.basename(image_path),
            "mimeType": f"image/{ext.lstrip('.')}",
        }
        if folder_id:
            file_meta["parents"] = [folder_id]
        self._state["files"][file_id] = file_meta
        return f"https://drive.google.com/uc?id={file_id}"

    # ── Test helpers ────────────────────────────────────────────────────

    def seed_presentation(self, presentation_id: str, json_dict: dict) -> None:
        """Load a presentation JSON dict into the mock store.

        Useful for seeding the mock with a snapshot captured from a real
        Google Slides API response.

        Args:
            presentation_id: The ID to use for storage (overrides any
                presentationId in the dict).
            json_dict: A presentation JSON dict as returned by
                get_presentation_json() on a real client.
        """
        data = copy.deepcopy(json_dict)
        data["presentationId"] = presentation_id
        self._state["presentations"][presentation_id] = data

    def get_batch_log(self) -> list:
        """Return the log of all processed batch requests.

        Each entry is a dict with:
            - "presentation_id": str
            - "requests": the raw request dicts sent to batchUpdate

        Useful for asserting that the correct API requests were generated.
        """
        return list(self._state["batch_log"])
