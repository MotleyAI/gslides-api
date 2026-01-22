"""Tests for concurrency safety of the GoogleAPIClient child client pattern.

This module tests that child clients properly share Google API services while
maintaining isolated batch state to enable safe concurrent operations.
"""

import asyncio
from unittest.mock import Mock, patch

import pytest

from gslides_api.client import GoogleAPIClient
from gslides_api.request.parent import GSlidesAPIRequest


class MockRequest(GSlidesAPIRequest):
    """Mock request class for testing."""

    request_id: str

    def to_request(self):
        return [{"mockRequest": {"id": self.request_id}}]


class TestCreateChildClient:
    """Test cases for the create_child_client method."""

    def test_create_child_client_shares_services(self):
        """Test that child client shares service objects with parent."""
        parent = GoogleAPIClient()

        # Mock all required components
        mock_credentials = Mock()
        mock_sheet_service = Mock()
        mock_slide_service = Mock()
        mock_drive_service = Mock()

        parent.crdtls = mock_credentials
        parent.sht_srvc = mock_sheet_service
        parent.sld_srvc = mock_slide_service
        parent.drive_srvc = mock_drive_service

        child = parent.create_child_client()

        # Services should be the exact same objects (identity, not equality)
        assert child.crdtls is parent.crdtls
        assert child.sht_srvc is parent.sht_srvc
        assert child.sld_srvc is parent.sld_srvc
        assert child.drive_srvc is parent.drive_srvc

    def test_create_child_client_has_independent_batch_state(self):
        """Test that child client has its own independent batch state."""
        parent = GoogleAPIClient()

        # Initialize parent
        parent.crdtls = Mock()
        parent.sht_srvc = Mock()
        parent.sld_srvc = Mock()
        parent.drive_srvc = Mock()

        # Add some state to parent
        parent.pending_batch_requests.append(MockRequest(request_id="parent_request"))
        parent.pending_presentation_id = "parent_presentation"

        child = parent.create_child_client()

        # Child should have empty batch state
        assert child.pending_batch_requests == []
        assert child.pending_presentation_id is None

        # Modifying child should not affect parent
        child.pending_batch_requests.append(MockRequest(request_id="child_request"))
        child.pending_presentation_id = "child_presentation"

        assert len(parent.pending_batch_requests) == 1
        assert parent.pending_batch_requests[0].request_id == "parent_request"
        assert parent.pending_presentation_id == "parent_presentation"

    def test_create_child_client_fails_if_not_initialized(self):
        """Test that creating child from uninitialized parent raises RuntimeError."""
        parent = GoogleAPIClient()

        # Parent is not initialized
        assert not parent.is_initialized

        with pytest.raises(RuntimeError) as exc_info:
            parent.create_child_client()

        assert "Cannot create child client from uninitialized parent client" in str(exc_info.value)

    def test_child_client_inherits_backoff_settings(self):
        """Test that child client inherits backoff settings from parent."""
        parent = GoogleAPIClient(
            auto_flush=True,
            initial_wait_s=120,
            n_backoffs=8,
        )

        # Initialize parent
        parent.crdtls = Mock()
        parent.sht_srvc = Mock()
        parent.sld_srvc = Mock()
        parent.drive_srvc = Mock()

        child = parent.create_child_client(auto_flush=False)

        # Backoff settings should be inherited
        assert child.initial_wait_s == 120
        assert child.n_backoffs == 8
        # auto_flush should be what was passed to create_child_client
        assert child.auto_flush is False

    def test_child_client_auto_flush_defaults_to_false(self):
        """Test that child client auto_flush defaults to False."""
        parent = GoogleAPIClient(auto_flush=True)

        parent.crdtls = Mock()
        parent.sht_srvc = Mock()
        parent.sld_srvc = Mock()
        parent.drive_srvc = Mock()

        child = parent.create_child_client()

        assert child.auto_flush is False

    def test_child_client_can_set_auto_flush_true(self):
        """Test that child client can be created with auto_flush=True."""
        parent = GoogleAPIClient(auto_flush=False)

        parent.crdtls = Mock()
        parent.sht_srvc = Mock()
        parent.sld_srvc = Mock()
        parent.drive_srvc = Mock()

        child = parent.create_child_client(auto_flush=True)

        assert child.auto_flush is True


class TestConcurrentBatchUpdates:
    """Test that concurrent batch operations don't interfere with each other."""

    def setup_method(self):
        """Set up test fixtures before each test method."""
        # Create parent client with mocked services
        self.parent = GoogleAPIClient()
        self.parent.crdtls = Mock()

        # Set up mock slide service
        self.mock_slide_service = Mock()
        self.mock_presentations = Mock()
        self.mock_batch_update = Mock()

        self.mock_slide_service.presentations.return_value = self.mock_presentations
        self.mock_presentations.batchUpdate.return_value = self.mock_batch_update
        self.mock_batch_update.execute.return_value = {
            "replies": [{"duplicateObject": {"objectId": "new_id"}}]
        }

        self.parent.sht_srvc = Mock()
        self.parent.sld_srvc = self.mock_slide_service
        self.parent.drive_srvc = Mock()

    def test_concurrent_batch_updates_isolated(self):
        """Test that batch updates in different child clients don't interfere."""
        child1 = self.parent.create_child_client(auto_flush=False)
        child2 = self.parent.create_child_client(auto_flush=False)

        # Add requests to different children for different presentations
        child1.batch_update(
            [MockRequest(request_id="req1")],
            "presentation_1"
        )
        child2.batch_update(
            [MockRequest(request_id="req2")],
            "presentation_2"
        )

        # Each child should have its own state
        assert len(child1.pending_batch_requests) == 1
        assert child1.pending_batch_requests[0].request_id == "req1"
        assert child1.pending_presentation_id == "presentation_1"

        assert len(child2.pending_batch_requests) == 1
        assert child2.pending_batch_requests[0].request_id == "req2"
        assert child2.pending_presentation_id == "presentation_2"

        # Flushing one should not affect the other
        child1.flush_batch_update()

        assert len(child1.pending_batch_requests) == 0
        assert child1.pending_presentation_id is None

        # child2 should be unaffected
        assert len(child2.pending_batch_requests) == 1
        assert child2.pending_batch_requests[0].request_id == "req2"
        assert child2.pending_presentation_id == "presentation_2"

    def test_multiple_children_share_same_service(self):
        """Test that multiple children use the same underlying slide service."""
        child1 = self.parent.create_child_client()
        child2 = self.parent.create_child_client()

        # Both children should use the same slide service instance
        assert child1.sld_srvc is child2.sld_srvc
        assert child1.sld_srvc is self.parent.sld_srvc

    def test_child_client_is_initialized(self):
        """Test that child client reports as initialized."""
        child = self.parent.create_child_client()

        assert child.is_initialized is True


class TestSharedServicesParameter:
    """Test the _shared_services parameter directly."""

    def test_shared_services_parameter_copies_services(self):
        """Test that _shared_services parameter copies services from source."""
        source = GoogleAPIClient()
        source.crdtls = Mock()
        source.sht_srvc = Mock()
        source.sld_srvc = Mock()
        source.drive_srvc = Mock()

        target = GoogleAPIClient(_shared_services=source)

        assert target.crdtls is source.crdtls
        assert target.sht_srvc is source.sht_srvc
        assert target.sld_srvc is source.sld_srvc
        assert target.drive_srvc is source.drive_srvc

    def test_shared_services_none_initializes_to_none(self):
        """Test that without _shared_services, services start as None."""
        client = GoogleAPIClient(_shared_services=None)

        assert client.crdtls is None
        assert client.sht_srvc is None
        assert client.sld_srvc is None
        assert client.drive_srvc is None

    def test_shared_services_with_uninitialized_source(self):
        """Test that _shared_services can be used with uninitialized source."""
        source = GoogleAPIClient()
        # source is uninitialized (all services are None)

        target = GoogleAPIClient(_shared_services=source)

        # Target should also have None services
        assert target.crdtls is None
        assert target.sht_srvc is None
        assert target.sld_srvc is None
        assert target.drive_srvc is None
        assert not target.is_initialized
