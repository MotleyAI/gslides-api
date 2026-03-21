"""Tests for GSlidesShapeElement.write_text style extraction behavior."""

from unittest.mock import Mock, call

import pytest

from gslides_api.domain.text import TextStyle
from gslides_api.element.shape import ShapeElement

from gslides_api.adapters.gslides_adapter import GSlidesShapeElement


def _make_shape_element(styles_return=None):
    """Create a mock ShapeElement with configurable styles return value."""
    mock_shape = Mock(spec=ShapeElement)
    mock_shape.objectId = "shape_123"
    mock_shape.presentation_id = "pres_123"
    mock_shape.slide_id = "slide_123"
    alt_text_mock = Mock()
    alt_text_mock.title = "Title"
    alt_text_mock.description = None
    mock_shape.alt_text = alt_text_mock
    mock_shape.styles.return_value = styles_return
    mock_shape.write_text.return_value = None
    return mock_shape


def _make_api_client():
    """Create a mock GSlidesAPIClient."""
    mock_api_client = Mock()
    mock_api_client.gslides_client = Mock()
    return mock_api_client


class TestWriteTextSkipsWhitespaceStyles:
    """Test that write_text uses skip_whitespace=True to avoid invisible spacer styles."""

    def test_write_text_calls_styles_with_skip_whitespace_true(self):
        """write_text should call styles(skip_whitespace=True) to avoid picking up
        invisible spacer styles like white theme colors from whitespace-only runs."""
        mock_shape = _make_shape_element(styles_return=[Mock(spec=TextStyle)])
        element = GSlidesShapeElement.model_validate(mock_shape)
        api_client = _make_api_client()

        element.write_text(api_client=api_client, content="Hello")

        mock_shape.styles.assert_called_once_with(skip_whitespace=True)

    def test_write_text_passes_extracted_styles_to_underlying_write(self):
        """write_text should pass the extracted styles to the underlying gslides_element.write_text."""
        mock_style = Mock(spec=TextStyle)
        mock_shape = _make_shape_element(styles_return=[mock_style])
        element = GSlidesShapeElement.model_validate(mock_shape)
        api_client = _make_api_client()

        element.write_text(api_client=api_client, content="Hello", autoscale=True)

        mock_shape.write_text.assert_called_once_with(
            "Hello",
            autoscale=True,
            styles=[mock_style],
            api_client=api_client.gslides_client,
        )

    def test_write_text_handles_none_styles(self):
        """write_text should handle styles() returning None gracefully."""
        mock_shape = _make_shape_element(styles_return=None)
        element = GSlidesShapeElement.model_validate(mock_shape)
        api_client = _make_api_client()

        element.write_text(api_client=api_client, content="Hello")

        mock_shape.write_text.assert_called_once_with(
            "Hello",
            autoscale=False,
            styles=None,
            api_client=api_client.gslides_client,
        )

    def test_write_text_handles_empty_styles_list(self):
        """write_text should handle styles() returning an empty list."""
        mock_shape = _make_shape_element(styles_return=[])
        element = GSlidesShapeElement.model_validate(mock_shape)
        api_client = _make_api_client()

        element.write_text(api_client=api_client, content="Hello")

        mock_shape.write_text.assert_called_once_with(
            "Hello",
            autoscale=False,
            styles=[],
            api_client=api_client.gslides_client,
        )
