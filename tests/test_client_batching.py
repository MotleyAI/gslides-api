"""
Tests for the GoogleAPIClient batching functionality.

This module tests the auto_flush parameter and batching behavior of the GoogleAPIClient,
including the batch_update and flush_batch_update methods.
"""

from typing import Any, Dict
from unittest.mock import MagicMock, Mock, patch

import pytest

from gslides_api.client import GoogleAPIClient
from gslides_api.request.request import (CreateShapeRequest,
                                         DeleteObjectRequest,
                                         DuplicateObjectRequest,
                                         GSlidesAPIRequest)


class MockRequest(GSlidesAPIRequest):
    """Mock request class for testing."""

    request_id: str

    def to_request(self):
        return [{"mockRequest": {"id": self.request_id}}]


class TestGoogleAPIClientBatching:
    """Test cases for GoogleAPIClient batching functionality."""

    def setup_method(self):
        """Set up test fixtures before each test method."""
        # Create mock services
        self.mock_slide_service = Mock()
        self.mock_presentations = Mock()
        self.mock_batch_update = Mock()

        # Set up the mock chain
        self.mock_slide_service.presentations.return_value = self.mock_presentations
        self.mock_presentations.batchUpdate.return_value = self.mock_batch_update
        self.mock_batch_update.execute.return_value = {
            "replies": [{"duplicateObject": {"objectId": "new_object_id"}}]
        }

    def test_auto_flush_true_by_default(self):
        """Test that auto_flush is True by default."""
        client = GoogleAPIClient()
        assert client.auto_flush is True
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None

    def test_auto_flush_false_initialization(self):
        """Test that auto_flush can be set to False during initialization."""
        client = GoogleAPIClient(auto_flush=False)
        assert client.auto_flush is False
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None

    def test_batch_update_with_auto_flush_true(self):
        """Test batch_update with auto_flush=True immediately executes requests."""
        client = GoogleAPIClient(auto_flush=True)
        client.sld_srvc = self.mock_slide_service

        requests = [MockRequest(request_id="test1"), MockRequest(request_id="test2")]
        presentation_id = "test_presentation"

        result = client.batch_update(requests, presentation_id)

        # Should execute immediately
        self.mock_presentations.batchUpdate.assert_called_once()
        call_args = self.mock_presentations.batchUpdate.call_args
        assert call_args[1]["presentationId"] == presentation_id
        assert len(call_args[1]["body"]["requests"]) == 2

        # Should clear pending requests
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None
        assert result == {
            "replies": [{"duplicateObject": {"objectId": "new_object_id"}}]
        }

    def test_batch_update_with_auto_flush_false(self):
        """Test batch_update with auto_flush=False accumulates requests."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        requests = [MockRequest(request_id="test1"), MockRequest(request_id="test2")]
        presentation_id = "test_presentation"

        result = client.batch_update(requests, presentation_id)

        # Should not execute immediately
        self.mock_presentations.batchUpdate.assert_not_called()

        # Should accumulate requests
        assert len(client.pending_batch_requests) == 2
        assert client.pending_presentation_id == presentation_id
        assert result == {}

    def test_batch_update_with_flush_parameter_override(self):
        """Test that flush parameter overrides auto_flush setting."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        requests = [MockRequest(request_id="test1")]
        presentation_id = "test_presentation"

        # Force flush despite auto_flush=False
        result = client.batch_update(requests, presentation_id, flush=True)

        # Should execute immediately
        self.mock_presentations.batchUpdate.assert_called_once()
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None

    def test_batch_update_different_presentation_flushes_previous(self):
        """Test that changing presentation ID flushes previous requests."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # First batch
        requests1 = [MockRequest(request_id="test1")]
        client.batch_update(requests1, "presentation1")

        # Verify first batch is pending
        assert len(client.pending_batch_requests) == 1
        assert client.pending_presentation_id == "presentation1"

        # Second batch with different presentation ID
        requests2 = [MockRequest(request_id="test2")]
        client.batch_update(requests2, "presentation2")

        # Should have flushed first batch and started new one
        self.mock_presentations.batchUpdate.assert_called_once()
        assert len(client.pending_batch_requests) == 1
        assert client.pending_presentation_id == "presentation2"

    def test_flush_batch_update_empty_requests(self):
        """Test flush_batch_update with no pending requests."""
        client = GoogleAPIClient()
        client.sld_srvc = self.mock_slide_service

        result = client.flush_batch_update()

        # Should return empty dict and not call API
        assert result == {}
        self.mock_presentations.batchUpdate.assert_not_called()

    def test_flush_batch_update_with_pending_requests(self):
        """Test flush_batch_update with pending requests."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # Add some pending requests
        requests = [MockRequest(request_id="test1"), MockRequest(request_id="test2")]
        client.batch_update(requests, "test_presentation")

        # Manually flush
        result = client.flush_batch_update()

        # Should execute and clear pending requests
        self.mock_presentations.batchUpdate.assert_called_once()
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None
        assert result == {
            "replies": [{"duplicateObject": {"objectId": "new_object_id"}}]
        }

    def test_duplicate_object_with_auto_flush_false(self):
        """Test duplicate_object method with auto_flush=False."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # duplicate_object should force flush by default since caller needs the ID
        object_id = client.duplicate_object("source_id", "presentation_id")

        # Should have executed immediately despite auto_flush=False
        self.mock_presentations.batchUpdate.assert_called_once()
        assert object_id == "new_object_id"
        assert client.pending_batch_requests == []

    def test_duplicate_object_with_auto_flush_false_and_id_map(self):
        """Test duplicate_object method with auto_flush=False."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # duplicate_object should force flush by default since caller needs the ID
        object_id = client.duplicate_object(
            "source_id", "presentation_id", id_map={"source_id": "new_object_id"}
        )

        assert object_id == "new_object_id"
        assert isinstance(client.pending_batch_requests[-1], DuplicateObjectRequest)
        assert client.pending_batch_requests[-1].objectId == "source_id"
        assert client.pending_batch_requests[-1].objectIds == {
            "source_id": "new_object_id"
        }

        # Manually flush
        result = client.flush_batch_update()

        # Should execute and clear pending requests
        self.mock_presentations.batchUpdate.assert_called_once()
        assert client.pending_batch_requests == []
        assert client.pending_presentation_id is None
        assert result == {
            "replies": [{"duplicateObject": {"objectId": "new_object_id"}}]
        }

    def test_delete_object_with_auto_flush_false(self):
        """Test delete_object method with auto_flush=False."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        client.delete_object("object_id", "presentation_id")

        # Should accumulate request without flushing
        assert len(client.pending_batch_requests) == 1
        assert client.pending_presentation_id == "presentation_id"
        self.mock_presentations.batchUpdate.assert_not_called()

    @patch("gslides_api.client.GoogleAPIClient.flush_batch_update")
    def test_non_batch_methods_flush_pending_requests(self, mock_flush):
        """Test that non-batchUpdate methods flush pending requests first."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # Set initial presentation ID to avoid flush on presentation change
        client.pending_presentation_id = "test_presentation"

        # Add some pending requests
        requests = [MockRequest(request_id="test1")]
        client.batch_update(requests, "test_presentation", flush=False)

        # Mock the get method to avoid actual API call
        mock_get = Mock()
        mock_get.execute.return_value = {"test": "data"}
        client.sld_srvc.presentations.return_value.get.return_value = mock_get

        # Call a non-batch method
        client.get_presentation_json("test_presentation")

        # Should have flushed pending requests first
        mock_flush.assert_called_once()

    def test_batch_update_assertion_error_for_invalid_requests(self):
        """Test that batch_update raises assertion error for invalid request types."""
        client = GoogleAPIClient()

        # Pass non-GSlidesAPIRequest objects
        invalid_requests = ["not_a_request", {"also": "not_a_request"}]

        with pytest.raises(AssertionError):
            client.batch_update(invalid_requests, "presentation_id")

    def test_batch_update_exception_handling(self):
        """Test that batch_update properly handles and re-raises exceptions."""
        client = GoogleAPIClient()
        client.sld_srvc = self.mock_slide_service

        # Make the execute method raise an exception
        self.mock_batch_update.execute.side_effect = Exception("API Error")

        requests = [MockRequest(request_id="test1")]

        with pytest.raises(Exception, match="API Error"):
            client.batch_update(requests, "test_presentation")

    def test_accumulate_multiple_batches_same_presentation(self):
        """Test accumulating multiple batches for the same presentation."""
        client = GoogleAPIClient(auto_flush=False)
        client.sld_srvc = self.mock_slide_service

        # First batch
        requests1 = [MockRequest(request_id="test1"), MockRequest(request_id="test2")]
        client.batch_update(requests1, "presentation_id")

        # Second batch, same presentation
        requests2 = [MockRequest(request_id="test3")]
        client.batch_update(requests2, "presentation_id")

        # Should accumulate all requests
        assert len(client.pending_batch_requests) == 3
        assert client.pending_presentation_id == "presentation_id"
        self.mock_presentations.batchUpdate.assert_not_called()

        # Flush and verify all requests are sent
        client.flush_batch_update()

        call_args = self.mock_presentations.batchUpdate.call_args
        assert len(call_args[1]["body"]["requests"]) == 3
