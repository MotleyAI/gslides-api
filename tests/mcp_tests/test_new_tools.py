"""Tests for new MCP tools: replace_element_image (file paths), write_table_markdown, bulk_write_element_markdown."""

import json
from unittest.mock import Mock, patch, call

import pytest

from gslides_api.element.base import ElementKind
from gslides_api.mcp.server import (
    bulk_write_element_markdown,
    replace_element_image,
    write_table_markdown,
)


# =============================================================================
# Fixtures
# =============================================================================


@pytest.fixture
def mock_api_client():
    """Create a mock API client."""
    client = Mock()
    client.flush_batch_update.return_value = None
    client.batch_update.return_value = None
    return client


@pytest.fixture
def mock_slide():
    """Create a mock slide with page_elements_flat."""
    slide = Mock()
    slide.objectId = "slide_001"
    return slide


@pytest.fixture
def mock_image_element():
    """Create a mock ImageElement."""
    from gslides_api.element.element import ImageElement

    element = Mock(spec=ImageElement)
    element.objectId = "img_001"
    element.type = ElementKind.IMAGE
    element.replace_image = Mock()
    return element


@pytest.fixture
def mock_shape_element():
    """Create a mock ShapeElement."""
    from gslides_api.element.shape import ShapeElement

    element = Mock(spec=ShapeElement)
    element.objectId = "shape_001"
    element.type = ElementKind.SHAPE
    element.write_text = Mock()
    return element


@pytest.fixture
def mock_table_element():
    """Create a mock TableElement."""
    from gslides_api.element.table import TableElement

    element = Mock(spec=TableElement)
    element.objectId = "table_001"
    element.type = ElementKind.TABLE
    element.table = Mock()
    element.table.rows = 3
    element.table.columns = 2
    element.resize = Mock(return_value=1.0)
    element.content_update_requests = Mock(return_value=[])
    return element


@pytest.fixture
def mock_presentation(mock_slide):
    """Create a mock presentation with one slide."""
    pres = Mock()
    pres.slides = [mock_slide]
    pres.presentationId = "pres_123"
    return pres


# =============================================================================
# Tests for replace_element_image (URL vs file path routing)
# =============================================================================


class TestReplaceElementImageRouting:
    """Test that replace_element_image routes URL vs file path correctly."""

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_url_routed_to_url_param(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_image_element,
    ):
        """Test that http/https URLs are passed as url= parameter."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_image_element

        result = json.loads(replace_element_image(
            "pres_123", "slide1", "my_image", "https://example.com/image.png"
        ))

        assert result["success"] is True
        mock_image_element.replace_image.assert_called_once_with(
            url="https://example.com/image.png", api_client=mock_api_client
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_http_url_routed_to_url_param(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_image_element,
    ):
        """Test that http:// URLs are also routed to url= parameter."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_image_element

        result = json.loads(replace_element_image(
            "pres_123", "slide1", "my_image", "http://example.com/image.png"
        ))

        assert result["success"] is True
        mock_image_element.replace_image.assert_called_once_with(
            url="http://example.com/image.png", api_client=mock_api_client
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_local_file_routed_to_file_param(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_image_element,
    ):
        """Test that local file paths are passed as file= parameter."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_image_element

        result = json.loads(replace_element_image(
            "pres_123", "slide1", "my_image", "/tmp/chart.png"
        ))

        assert result["success"] is True
        mock_image_element.replace_image.assert_called_once_with(
            file="/tmp/chart.png", api_client=mock_api_client
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_relative_file_routed_to_file_param(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_image_element,
    ):
        """Test that relative file paths are passed as file= parameter."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_image_element

        result = json.loads(replace_element_image(
            "pres_123", "slide1", "my_image", "images/chart.png"
        ))

        assert result["success"] is True
        mock_image_element.replace_image.assert_called_once_with(
            file="images/chart.png", api_client=mock_api_client
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_response_contains_image_source(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_image_element,
    ):
        """Test that the response includes image_source field."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_image_element

        result = json.loads(replace_element_image(
            "pres_123", "slide1", "my_image", "/tmp/chart.png"
        ))

        assert result["details"]["image_source"] == "/tmp/chart.png"


# =============================================================================
# Tests for write_table_markdown
# =============================================================================


class TestWriteTableMarkdown:
    """Tests for the write_table_markdown tool."""

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_write_table_same_shape(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_table_element,
    ):
        """Test writing a table that matches the existing table shape (no resize)."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_table_element

        # Table has 3 rows, 2 columns - markdown matches
        md_table = "| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |"

        with patch("gslides_api.mcp.server.MarkdownTableElement") as mock_mte:
            mock_md_elem = Mock()
            mock_md_elem.shape = (3, 2)
            mock_mte.from_markdown.return_value = mock_md_elem

            result = json.loads(write_table_markdown("pres_123", "slide1", "my_table", md_table))

        assert result["success"] is True
        assert result["details"]["table_shape"] == [3, 2]
        assert result["details"]["resized"] is False
        mock_table_element.resize.assert_not_called()
        mock_table_element.content_update_requests.assert_called_once_with(
            mock_md_elem, check_shape=False, font_scale_factor=1.0
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_write_table_with_resize(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_table_element,
    ):
        """Test writing a table that requires resizing."""
        mock_get_client.return_value = mock_api_client

        # After resize, re-fetch returns updated element
        resized_table_element = Mock()
        resized_table_element.objectId = "table_001"
        resized_table_element.content_update_requests = Mock(return_value=[])

        # First call returns original, second call returns resized
        mock_pres_class.from_id.side_effect = [
            Mock(slides=[mock_slide]),
            Mock(slides=[mock_slide]),
        ]
        mock_find_slide.return_value = mock_slide
        mock_find_element.side_effect = [mock_table_element, resized_table_element]

        # Table has 3 rows, 2 cols but markdown has 4 rows, 3 cols
        md_table = "| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |\n| 4 | 5 | 6 |\n| 7 | 8 | 9 |"
        mock_table_element.resize.return_value = 0.8

        with patch("gslides_api.mcp.server.MarkdownTableElement") as mock_mte:
            mock_md_elem = Mock()
            mock_md_elem.shape = (4, 3)
            mock_mte.from_markdown.return_value = mock_md_elem

            result = json.loads(write_table_markdown("pres_123", "slide1", "my_table", md_table))

        assert result["success"] is True
        assert result["details"]["resized"] is True
        assert result["details"]["table_shape"] == [4, 3]
        mock_table_element.resize.assert_called_once_with(4, 3, api_client=mock_api_client)
        resized_table_element.content_update_requests.assert_called_once_with(
            mock_md_elem, check_shape=False, font_scale_factor=0.8
        )

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.find_slide_by_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_write_table_not_a_table(
        self, mock_get_client, mock_pres_class, mock_find_slide, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test error when element is not a table."""
        mock_get_client.return_value = mock_api_client
        mock_pres_class.from_id.return_value = Mock(slides=[mock_slide])
        mock_find_slide.return_value = mock_slide
        mock_find_element.return_value = mock_shape_element

        result = json.loads(write_table_markdown("pres_123", "slide1", "not_table", "| A |\n|---|\n| 1 |"))

        assert result["error"] is True
        assert "not a table element" in result["message"]

    def test_write_table_invalid_presentation_url(self):
        """Test error with invalid presentation URL."""
        result = json.loads(write_table_markdown(
            "https://example.com/bad-url", "slide1", "table1", "| A |\n|---|\n| 1 |"
        ))
        assert result["error"] is True
        assert result["error_type"] == "ValidationError"


# =============================================================================
# Tests for bulk_write_element_markdown
# =============================================================================


class TestBulkWriteElementMarkdown:
    """Tests for the bulk_write_element_markdown tool."""

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_successful_bulk_write(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test successful bulk write to multiple elements."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"
        mock_find_element.return_value = mock_shape_element

        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "title", "markdown": "# Hello"},
            {"slide_name": "slide1", "element_name": "body", "markdown": "World"},
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 2
        assert result["details"]["failed"] == 0
        assert mock_shape_element.write_text.call_count == 2

    def test_invalid_json(self):
        """Test error with invalid JSON input."""
        result = json.loads(bulk_write_element_markdown("pres_123", "not valid json{"))

        assert result["error"] is True
        assert result["error_type"] == "ValidationError"
        assert "Invalid JSON" in result["message"]

    def test_json_not_array(self):
        """Test error when JSON is not an array."""
        result = json.loads(bulk_write_element_markdown("pres_123", '{"key": "value"}'))

        assert result["error"] is True
        assert "Expected a JSON array" in result["message"]

    def test_missing_keys(self):
        """Test error when entries are missing required keys."""
        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "title"},  # missing "markdown"
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["error"] is True
        assert "missing keys" in result["message"]

    def test_entry_not_object(self):
        """Test error when an entry is not an object."""
        writes = json.dumps(["not an object"])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["error"] is True
        assert "Entry 0 is not an object" in result["message"]

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_partial_failure_slide_not_found(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test that missing slides are reported as failures without blocking others."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"
        mock_find_element.return_value = mock_shape_element

        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "title", "markdown": "# Hello"},
            {"slide_name": "nonexistent", "element_name": "body", "markdown": "World"},
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 1
        assert result["details"]["failed"] == 1
        assert "not found" in result["details"]["failures"][0]["error"]

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_partial_failure_element_not_found(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test that missing elements are reported as failures without blocking others."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"

        # First call returns shape, second returns None (not found)
        mock_find_element.side_effect = [mock_shape_element, None]

        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "title", "markdown": "# Hello"},
            {"slide_name": "slide1", "element_name": "missing_elem", "markdown": "World"},
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 1
        assert result["details"]["failed"] == 1
        assert "not found" in result["details"]["failures"][0]["error"]

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_partial_failure_wrong_element_type(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_table_element, mock_shape_element,
    ):
        """Test that non-shape elements are reported as failures."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"

        # First returns table (wrong type), second returns shape (correct)
        mock_find_element.side_effect = [mock_table_element, mock_shape_element]

        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "my_table", "markdown": "# Hello"},
            {"slide_name": "slide1", "element_name": "title", "markdown": "World"},
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 1
        assert result["details"]["failed"] == 1
        assert "not a text element" in result["details"]["failures"][0]["error"]

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_write_text_exception_captured(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test that exceptions during write_text are captured per-element."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"

        # Create two separate shape elements
        from gslides_api.element.shape import ShapeElement
        good_element = Mock(spec=ShapeElement)
        good_element.objectId = "shape_good"
        good_element.type = ElementKind.SHAPE
        good_element.write_text = Mock()

        bad_element = Mock(spec=ShapeElement)
        bad_element.objectId = "shape_bad"
        bad_element.type = ElementKind.SHAPE
        bad_element.write_text = Mock(side_effect=RuntimeError("API error"))

        mock_find_element.side_effect = [bad_element, good_element]

        writes = json.dumps([
            {"slide_name": "slide1", "element_name": "bad", "markdown": "fail"},
            {"slide_name": "slide1", "element_name": "good", "markdown": "succeed"},
        ])

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 1
        assert result["details"]["failed"] == 1
        assert "API error" in result["details"]["failures"][0]["error"]

    def test_invalid_presentation_url(self):
        """Test error with invalid presentation URL."""
        writes = json.dumps([{"slide_name": "s", "element_name": "e", "markdown": "m"}])
        result = json.loads(bulk_write_element_markdown("https://example.com/bad", writes))

        assert result["error"] is True
        assert result["error_type"] == "ValidationError"

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_empty_writes_list(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client,
    ):
        """Test with an empty writes list."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = []
        mock_pres_class.from_id.return_value = mock_pres

        result = json.loads(bulk_write_element_markdown("pres_123", "[]"))

        assert result["success"] is True
        assert result["details"]["succeeded"] == 0
        assert result["details"]["failed"] == 0

    @patch("gslides_api.mcp.server.find_element_by_name")
    @patch("gslides_api.mcp.server.get_slide_name")
    @patch("gslides_api.mcp.server.Presentation")
    @patch("gslides_api.mcp.server.get_api_client")
    def test_markdown_with_escaped_newlines(
        self, mock_get_client, mock_pres_class, mock_get_slide_name, mock_find_element,
        mock_api_client, mock_slide, mock_shape_element,
    ):
        """Test that JSON-escaped newlines in markdown are handled correctly."""
        mock_get_client.return_value = mock_api_client
        mock_pres = Mock()
        mock_pres.slides = [mock_slide]
        mock_pres_class.from_id.return_value = mock_pres
        mock_get_slide_name.return_value = "slide1"
        mock_find_element.return_value = mock_shape_element

        # JSON with escaped newlines
        writes = '[{"slide_name": "slide1", "element_name": "title", "markdown": "line1\\nline2\\nline3"}]'

        result = json.loads(bulk_write_element_markdown("pres_123", writes))

        assert result["success"] is True
        # Verify the markdown was passed with actual newlines
        mock_shape_element.write_text.assert_called_once_with(
            "line1\nline2\nline3", as_markdown=True, api_client=mock_api_client
        )
