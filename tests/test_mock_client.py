"""Tests for MockGoogleAPIClient."""

import json
import os
import tempfile

import pytest

from gslides_api.mock import (
    MockGoogleAPIClient,
    dump_presentation_snapshot,
    load_presentation_snapshot,
)
from gslides_api.mock.batch_processor import process_batch_requests
from gslides_api.presentation import Presentation
from gslides_api.request.request import (
    CreateSlideRequest,
    DeleteObjectRequest,
    DuplicateObjectRequest,
    InsertTextRequest,
    UpdateTextStyleRequest,
)


# ── Basic client operations ─────────────────────────────────────────────


class TestMockClientBasics:
    def test_is_initialized(self):
        mock = MockGoogleAPIClient()
        assert mock.is_initialized is True

    def test_set_credentials_noop(self):
        mock = MockGoogleAPIClient()
        mock.set_credentials(None)
        assert mock.is_initialized is True

    def test_initialize_credentials_noop(self):
        mock = MockGoogleAPIClient()
        mock.initialize_credentials("/nonexistent/path")
        assert mock.is_initialized is True

    def test_create_presentation(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test Presentation"})
        assert pres_id is not None
        assert isinstance(pres_id, str)

    def test_get_presentation_json(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "My Deck"})
        data = mock.get_presentation_json(pres_id)
        assert data["presentationId"] == pres_id
        assert data["title"] == "My Deck"
        assert len(data["slides"]) == 1  # default blank slide

    def test_get_presentation_not_found(self):
        mock = MockGoogleAPIClient()
        with pytest.raises(KeyError, match="not found"):
            mock.get_presentation_json("nonexistent")

    def test_get_slide_json(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        pres = mock.get_presentation_json(pres_id)
        slide_id = pres["slides"][0]["objectId"]

        slide = mock.get_slide_json(pres_id, slide_id)
        assert slide["objectId"] == slide_id

    def test_get_slide_not_found(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        with pytest.raises(KeyError, match="not found"):
            mock.get_slide_json(pres_id, "nonexistent_slide")

    def test_create_presentation_returns_unique_ids(self):
        mock = MockGoogleAPIClient()
        id1 = mock.create_presentation({"title": "A"})
        id2 = mock.create_presentation({"title": "B"})
        assert id1 != id2


# ── Drive operations ────────────────────────────────────────────────────


class TestMockDriveOperations:
    def test_copy_presentation(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Original"})
        result = mock.copy_presentation(pres_id, "Copy of Original")

        assert "id" in result
        assert result["name"] == "Copy of Original"

        # The copy should be loadable
        copy_data = mock.get_presentation_json(result["id"])
        assert copy_data["title"] == "Copy of Original"

    def test_copy_presentation_not_found(self):
        mock = MockGoogleAPIClient()
        with pytest.raises(KeyError):
            mock.copy_presentation("nonexistent", "copy")

    def test_create_folder(self):
        mock = MockGoogleAPIClient()
        result = mock.create_folder("My Folder")
        assert "id" in result
        assert result["name"] == "My Folder"

    def test_create_folder_ignore_existing(self):
        mock = MockGoogleAPIClient()
        result1 = mock.create_folder("Shared")
        result2 = mock.create_folder("Shared", ignore_existing=True)
        assert result1["id"] == result2["id"]

    def test_create_folder_no_ignore(self):
        mock = MockGoogleAPIClient()
        result1 = mock.create_folder("Shared")
        result2 = mock.create_folder("Shared", ignore_existing=False)
        assert result1["id"] != result2["id"]

    def test_delete_file(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "To Delete"})
        mock.delete_file(pres_id)
        with pytest.raises(KeyError):
            mock.get_presentation_json(pres_id)

    def test_upload_image_to_drive(self):
        mock = MockGoogleAPIClient()
        # Just checks the URL format and validation — doesn't need a real file
        url = mock.upload_image_to_drive("/fake/path/image.png")
        assert url.startswith("https://drive.google.com/uc?id=")

    def test_upload_image_unsupported_format(self):
        mock = MockGoogleAPIClient()
        with pytest.raises(ValueError, match="Unsupported"):
            mock.upload_image_to_drive("/fake/path/file.bmp")


# ── Batch operations ────────────────────────────────────────────────────


class TestMockBatchOperations:
    def test_create_slide_via_batch(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})

        request = CreateSlideRequest()
        result = mock.batch_update([request], pres_id)

        assert "replies" in result
        new_slide_id = result["replies"][0]["createSlide"]["objectId"]
        assert new_slide_id is not None

        # Verify the slide was added
        pres = mock.get_presentation_json(pres_id)
        slide_ids = [s["objectId"] for s in pres["slides"]]
        assert new_slide_id in slide_ids

    def test_duplicate_object(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        pres = mock.get_presentation_json(pres_id)
        slide_id = pres["slides"][0]["objectId"]

        new_id = mock.duplicate_object(slide_id, pres_id)
        assert new_id != slide_id

        # Should now have 2 slides
        pres = mock.get_presentation_json(pres_id)
        assert len(pres["slides"]) == 2

    def test_duplicate_object_with_id_map(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        pres = mock.get_presentation_json(pres_id)
        slide_id = pres["slides"][0]["objectId"]

        id_map = {slide_id: "my_custom_id"}
        new_id = mock.duplicate_object(slide_id, pres_id, id_map=id_map)
        assert new_id == "my_custom_id"

    def test_delete_object(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})

        # Add a second slide, then delete the first
        request = CreateSlideRequest()
        result = mock.batch_update([request], pres_id)
        new_slide_id = result["replies"][0]["createSlide"]["objectId"]

        pres = mock.get_presentation_json(pres_id)
        original_slide_id = pres["slides"][0]["objectId"]

        mock.delete_object(original_slide_id, pres_id)
        mock.flush_batch_update()

        pres = mock.get_presentation_json(pres_id)
        assert len(pres["slides"]) == 1
        assert pres["slides"][0]["objectId"] == new_slide_id

    def test_passthrough_requests_recorded_in_log(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})

        request = InsertTextRequest(
            objectId="some_element", text="Hello", insertionIndex=0
        )
        mock.batch_update([request], pres_id)

        log = mock.get_batch_log()
        assert len(log) > 0
        # Find the entry with our InsertText request
        last_entry = log[-1]
        assert last_entry["presentation_id"] == pres_id

    def test_auto_flush_false(self):
        mock = MockGoogleAPIClient(auto_flush=False)
        pres_id = mock.create_presentation({"title": "Test"})

        request = CreateSlideRequest()
        result = mock.batch_update([request], pres_id)
        # With auto_flush=False, batch_update returns empty dict
        assert result == {}
        assert len(mock.pending_batch_requests) == 1

        # Manual flush
        result = mock.flush_batch_update()
        assert "replies" in result
        assert len(mock.pending_batch_requests) == 0


# ── Child client ────────────────────────────────────────────────────────


class TestMockChildClient:
    def test_child_shares_state(self):
        parent = MockGoogleAPIClient()
        pres_id = parent.create_presentation({"title": "Shared"})

        child = parent.create_child_client(auto_flush=False)
        assert isinstance(child, MockGoogleAPIClient)

        # Child can see parent's presentations
        data = child.get_presentation_json(pres_id)
        assert data["title"] == "Shared"

    def test_child_has_isolated_batch_state(self):
        parent = MockGoogleAPIClient(auto_flush=False)
        pres_id = parent.create_presentation({"title": "Test"})

        child = parent.create_child_client(auto_flush=False)

        request = CreateSlideRequest()
        child.batch_update([request], pres_id)
        assert len(child.pending_batch_requests) == 1
        assert len(parent.pending_batch_requests) == 0

    def test_child_mutations_visible_to_parent(self):
        parent = MockGoogleAPIClient()
        pres_id = parent.create_presentation({"title": "Test"})

        child = parent.create_child_client(auto_flush=True)
        child.batch_update([CreateSlideRequest()], pres_id)

        # Parent should see the new slide
        pres = parent.get_presentation_json(pres_id)
        assert len(pres["slides"]) == 2


# ── Integration with Presentation model ─────────────────────────────────


class TestMockWithPresentation:
    def test_create_blank_presentation(self):
        mock = MockGoogleAPIClient()
        pres = Presentation.create_blank("Integration Test", api_client=mock)
        assert pres.presentationId is not None
        assert pres.title == "Integration Test"

    def test_from_id_round_trip(self):
        mock = MockGoogleAPIClient()
        pres = Presentation.create_blank("Round Trip", api_client=mock)
        loaded = Presentation.from_id(pres.presentationId, api_client=mock)
        assert loaded.presentationId == pres.presentationId
        assert loaded.title == "Round Trip"

    def test_copy_via_drive(self):
        mock = MockGoogleAPIClient()
        pres = Presentation.create_blank("Original", api_client=mock)
        copy_pres = pres.copy_via_drive("The Copy", api_client=mock)
        assert copy_pres.title == "The Copy"
        assert copy_pres.presentationId != pres.presentationId

    def test_delete_slide(self):
        mock = MockGoogleAPIClient()
        pres = Presentation.create_blank("Delete Test", api_client=mock)

        # Add a second slide
        request = CreateSlideRequest()
        result = mock.batch_update([request], pres.presentationId)
        new_slide_id = result["replies"][0]["createSlide"]["objectId"]

        original_slide_id = pres.slides[0].objectId
        pres.delete_slide(original_slide_id, api_client=mock)
        mock.flush_batch_update()

        reloaded = Presentation.from_id(pres.presentationId, api_client=mock)
        assert len(reloaded.slides) == 1
        assert reloaded.slides[0].objectId == new_slide_id


# ── Snapshot utilities ──────────────────────────────────────────────────


class TestSnapshots:
    def test_seed_presentation(self):
        mock = MockGoogleAPIClient()
        snapshot = {
            "presentationId": "original_id",
            "title": "Snapshot Test",
            "pageSize": {
                "width": {"magnitude": 9144000, "unit": "EMU"},
                "height": {"magnitude": 5143500, "unit": "EMU"},
            },
            "slides": [
                {
                    "objectId": "slide_1",
                    "pageElements": [],
                    "slideProperties": {
                        "notesPage": {
                            "objectId": "notes_1",
                            "pageElements": [],
                            "notesProperties": {
                                "speakerNotesObjectId": "speaker_1"
                            },
                            "pageType": "NOTES",
                        }
                    },
                }
            ],
            "layouts": [],
            "masters": [],
        }

        mock.seed_presentation("my_pres", snapshot)
        data = mock.get_presentation_json("my_pres")
        assert data["title"] == "Snapshot Test"
        assert data["presentationId"] == "my_pres"
        assert len(data["slides"]) == 1

    def test_load_and_dump_snapshot(self):
        # Create a mock with data, "dump" it, then load into a fresh mock
        source_mock = MockGoogleAPIClient()
        pres_id = source_mock.create_presentation({"title": "Snapshot Source"})

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False
        ) as f:
            tmp_path = f.name

        try:
            # dump_presentation_snapshot works with any client that has
            # get_presentation_json — including MockGoogleAPIClient
            dump_presentation_snapshot(source_mock, pres_id, tmp_path)

            # Verify JSON was written
            with open(tmp_path) as f:
                saved = json.load(f)
            assert saved["title"] == "Snapshot Source"

            # Load into a fresh mock
            target_mock = MockGoogleAPIClient()
            loaded_id = load_presentation_snapshot(target_mock, tmp_path)
            assert loaded_id == pres_id

            data = target_mock.get_presentation_json(loaded_id)
            assert data["title"] == "Snapshot Source"
        finally:
            os.unlink(tmp_path)

    def test_load_snapshot_with_custom_id(self):
        source_mock = MockGoogleAPIClient()
        pres_id = source_mock.create_presentation({"title": "Custom ID"})

        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False
        ) as f:
            tmp_path = f.name

        try:
            dump_presentation_snapshot(source_mock, pres_id, tmp_path)

            target_mock = MockGoogleAPIClient()
            loaded_id = load_presentation_snapshot(
                target_mock, tmp_path, presentation_id="custom_id"
            )
            assert loaded_id == "custom_id"
            data = target_mock.get_presentation_json("custom_id")
            assert data["presentationId"] == "custom_id"
        finally:
            os.unlink(tmp_path)


# ── Batch processor edge cases ──────────────────────────────────────────


class TestBatchProcessorEdgeCases:
    def test_delete_nonexistent_object(self):
        """Deleting a nonexistent object should not raise."""
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        mock.delete_object("nonexistent", pres_id)
        mock.flush_batch_update()  # should not raise

    def test_update_slides_position(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        pres = mock.get_presentation_json(pres_id)
        first_slide_id = pres["slides"][0]["objectId"]

        # Add a second slide
        result = mock.batch_update([CreateSlideRequest()], pres_id)
        second_slide_id = result["replies"][0]["createSlide"]["objectId"]

        # Move second slide to position 0
        from gslides_api.request.request import UpdateSlidesPositionRequest

        request = UpdateSlidesPositionRequest(
            slideObjectIds=[second_slide_id], insertionIndex=0
        )
        mock.batch_update([request], pres_id)

        pres = mock.get_presentation_json(pres_id)
        assert pres["slides"][0]["objectId"] == second_slide_id
        assert pres["slides"][1]["objectId"] == first_slide_id

    def test_slide_thumbnail_stub(self):
        mock = MockGoogleAPIClient()
        pres_id = mock.create_presentation({"title": "Test"})
        pres = mock.get_presentation_json(pres_id)
        slide_id = pres["slides"][0]["objectId"]

        from gslides_api.domain.domain import ThumbnailProperties

        thumb = mock.slide_thumbnail(pres_id, slide_id, ThumbnailProperties())
        assert thumb.contentUrl.startswith("https://mock.test/thumbnail/")
        assert thumb.width > 0
        assert thumb.height > 0
