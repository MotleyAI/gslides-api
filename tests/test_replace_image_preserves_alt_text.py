from unittest.mock import Mock, patch

from gslides_api.domain.domain import Dimension, Image, ImageReplaceMethod, Size, Transform, Unit
from gslides_api.element.base import ElementKind
from gslides_api.element.image import ImageElement
from gslides_api.request.request import ReplaceImageRequest, UpdatePageElementAltTextRequest


class TestReplaceImagePreservesAltText:
    """Test that replacing an image preserves the element's alt-text title and description."""

    def _make_image_element(self, title=None, description=None):
        image = Image(
            contentUrl="https://example.com/old.png",
            sourceUrl="https://example.com/old.png",
        )
        transform = Transform(translateX=0, translateY=0, scaleX=1, scaleY=1)
        size = Size(
            width=Dimension(magnitude=914400, unit=Unit.EMU),
            height=Dimension(magnitude=914400, unit=Unit.EMU),
        )
        return ImageElement(
            objectId="test-image-id",
            image=image,
            transform=transform,
            size=size,
            type=ElementKind.IMAGE,
            title=title,
            description=description,
            presentation_id="test-pres-id",
            slide_id="test-slide-id",
        )

    def test_replace_image_from_id_with_title_includes_alt_text_request(self):
        """replace_image_from_id should include UpdatePageElementAltTextRequest when title is given."""
        mock_client = Mock()
        mock_client.auto_flush = False

        # Capture the requests passed to batch_update
        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        ImageElement.replace_image_from_id(
            image_id="test-image-id",
            presentation_id="test-pres-id",
            url="https://example.com/new.png",
            api_client=mock_client,
            title="my_chart",
            description="Chart description",
        )

        assert len(captured_requests) == 2
        assert isinstance(captured_requests[0], ReplaceImageRequest)
        assert isinstance(captured_requests[1], UpdatePageElementAltTextRequest)

        alt_text_req = captured_requests[1]
        assert alt_text_req.objectId == "test-image-id"
        assert alt_text_req.title == "my_chart"
        assert alt_text_req.description == "Chart description"

    def test_replace_image_from_id_without_title_no_alt_text_request(self):
        """replace_image_from_id should NOT include alt-text request when no title/description."""
        mock_client = Mock()
        mock_client.auto_flush = False

        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        ImageElement.replace_image_from_id(
            image_id="test-image-id",
            presentation_id="test-pres-id",
            url="https://example.com/new.png",
            api_client=mock_client,
        )

        assert len(captured_requests) == 1
        assert isinstance(captured_requests[0], ReplaceImageRequest)

    def test_replace_image_from_id_with_title_only(self):
        """replace_image_from_id should handle title without description."""
        mock_client = Mock()
        mock_client.auto_flush = False

        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        ImageElement.replace_image_from_id(
            image_id="test-image-id",
            presentation_id="test-pres-id",
            url="https://example.com/new.png",
            api_client=mock_client,
            title="my_chart",
        )

        assert len(captured_requests) == 2
        alt_text_req = captured_requests[1]
        assert alt_text_req.title == "my_chart"
        assert alt_text_req.description is None

    def test_replace_image_from_id_with_file_upload_preserves_title(self):
        """replace_image_from_id should preserve title when using file upload."""
        mock_client = Mock()
        mock_client.auto_flush = False
        mock_client.upload_image_to_drive.return_value = "https://drive.google.com/uc?id=abc123"

        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        ImageElement.replace_image_from_id(
            image_id="test-image-id",
            presentation_id="test-pres-id",
            file="/path/to/chart.png",
            api_client=mock_client,
            title="chart_element",
        )

        mock_client.upload_image_to_drive.assert_called_once_with("/path/to/chart.png")
        assert len(captured_requests) == 2
        assert isinstance(captured_requests[1], UpdatePageElementAltTextRequest)
        assert captured_requests[1].title == "chart_element"

    def test_replace_image_instance_method_passes_title(self):
        """ImageElement.replace_image() should pass self.title and self.description to replace_image_from_id."""
        element = self._make_image_element(
            title="my_element_name",
            description="my description",
        )

        mock_client = Mock()
        mock_client.auto_flush = False

        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        element.replace_image(
            url="https://example.com/new.png",
            api_client=mock_client,
            enforce_size=False,
        )

        assert len(captured_requests) == 2
        alt_text_req = captured_requests[1]
        assert isinstance(alt_text_req, UpdatePageElementAltTextRequest)
        assert alt_text_req.objectId == "test-image-id"
        assert alt_text_req.title == "my_element_name"
        assert alt_text_req.description == "my description"

    def test_replace_image_instance_method_no_title(self):
        """ImageElement.replace_image() with no title/description should not add alt-text request."""
        element = self._make_image_element(title=None, description=None)

        mock_client = Mock()
        mock_client.auto_flush = False

        captured_requests = []
        mock_client.batch_update.side_effect = lambda reqs, pres_id: captured_requests.extend(reqs)

        element.replace_image(
            url="https://example.com/new.png",
            api_client=mock_client,
            enforce_size=False,
        )

        assert len(captured_requests) == 1
        assert isinstance(captured_requests[0], ReplaceImageRequest)
