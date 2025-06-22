"""
Tests for the requests module, specifically CreateParagraphBulletsRequest and related classes.
"""

import pytest
from pydantic import ValidationError

from gslides_api.request.domain import Range, RangeType, TableCellLocation, ElementProperties
from gslides_api.request.request import (
    CreateParagraphBulletsRequest,
    InsertTextRequest,
    UpdateTextStyleRequest,
    DeleteTextRequest,
    CreateShapeRequest,
    UpdateShapePropertiesRequest,
    ReplaceImageRequest,
    CreateSlideRequest,
    UpdateSlidePropertiesRequest,
    UpdateSlidesPositionRequest,
    UpdatePagePropertiesRequest,
    DeleteObjectRequest,
    DuplicateObjectRequest,
)
from gslides_api.domain import (
    BulletGlyphPreset,
    TextStyle,
    ShapeType,
    LayoutReference,
    PredefinedLayout,
    ShapeProperties,
    Size,
    Transform
)


class TestRange:
    """Test cases for the Range class."""

    def test_all_range_valid(self):
        """Test that ALL range type works correctly."""
        range_all = Range(type=RangeType.ALL)
        assert range_all.type == RangeType.ALL
        assert range_all.startIndex is None
        assert range_all.endIndex is None

    def test_fixed_range_valid(self):
        """Test that FIXED_RANGE type works correctly."""
        range_fixed = Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=10)
        assert range_fixed.type == RangeType.FIXED_RANGE
        assert range_fixed.startIndex == 0
        assert range_fixed.endIndex == 10

    def test_from_start_index_valid(self):
        """Test that FROM_START_INDEX type works correctly."""
        range_from_start = Range(type=RangeType.FROM_START_INDEX, startIndex=5)
        assert range_from_start.type == RangeType.FROM_START_INDEX
        assert range_from_start.startIndex == 5
        assert range_from_start.endIndex is None

    def test_all_range_with_indexes_invalid(self):
        """Test that ALL range with indexes raises validation error."""
        with pytest.raises(
            ValidationError, match="startIndex and endIndex must be None when type is ALL"
        ):
            Range(type=RangeType.ALL, startIndex=0)

    def test_fixed_range_missing_end_index_invalid(self):
        """Test that FIXED_RANGE without endIndex raises validation error."""
        with pytest.raises(
            ValidationError,
            match="Both startIndex and endIndex must be provided when type is FIXED_RANGE",
        ):
            Range(type=RangeType.FIXED_RANGE, startIndex=0)

    def test_fixed_range_missing_start_index_invalid(self):
        """Test that FIXED_RANGE without startIndex raises validation error."""
        with pytest.raises(
            ValidationError,
            match="Both startIndex and endIndex must be provided when type is FIXED_RANGE",
        ):
            Range(type=RangeType.FIXED_RANGE, endIndex=10)

    def test_fixed_range_start_greater_than_end_invalid(self):
        """Test that FIXED_RANGE with startIndex >= endIndex raises validation error."""
        with pytest.raises(ValidationError, match="startIndex must be less than endIndex"):
            Range(type=RangeType.FIXED_RANGE, startIndex=5, endIndex=3)

        with pytest.raises(ValidationError, match="startIndex must be less than endIndex"):
            Range(type=RangeType.FIXED_RANGE, startIndex=5, endIndex=5)

    def test_from_start_index_missing_start_invalid(self):
        """Test that FROM_START_INDEX without startIndex raises validation error."""
        with pytest.raises(
            ValidationError, match="startIndex must be provided when type is FROM_START_INDEX"
        ):
            Range(type=RangeType.FROM_START_INDEX)

    def test_from_start_index_with_end_index_invalid(self):
        """Test that FROM_START_INDEX with endIndex raises validation error."""
        with pytest.raises(
            ValidationError, match="endIndex must be None when type is FROM_START_INDEX"
        ):
            Range(type=RangeType.FROM_START_INDEX, startIndex=5, endIndex=10)


class TestTableCellLocation:
    """Test cases for the TableCellLocation class."""

    def test_valid_cell_location(self):
        """Test that valid cell location works correctly."""
        cell = TableCellLocation(rowIndex=1, columnIndex=2)
        assert cell.rowIndex == 1
        assert cell.columnIndex == 2

    def test_zero_indexes_valid(self):
        """Test that zero indexes are valid."""
        cell = TableCellLocation(rowIndex=0, columnIndex=0)
        assert cell.rowIndex == 0
        assert cell.columnIndex == 0

    def test_negative_row_index_invalid(self):
        """Test that negative row index raises validation error."""
        with pytest.raises(ValidationError, match="rowIndex must be non-negative"):
            TableCellLocation(rowIndex=-1, columnIndex=0)

    def test_negative_column_index_invalid(self):
        """Test that negative column index raises validation error."""
        with pytest.raises(ValidationError, match="columnIndex must be non-negative"):
            TableCellLocation(rowIndex=0, columnIndex=-1)


class TestCreateParagraphBulletsRequest:
    """Test cases for the CreateParagraphBulletsRequest class."""

    def test_shape_request_valid(self):
        """Test that request for shape (without cell location) works correctly."""
        request = CreateParagraphBulletsRequest(
            objectId="shape_123",
            textRange=Range(type=RangeType.ALL),
            bulletPreset=BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE,
        )
        assert request.objectId == "shape_123"
        assert request.textRange.type == RangeType.ALL
        assert request.bulletPreset == BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE
        assert request.cellLocation is None

    def test_table_request_valid(self):
        """Test that request for table (with cell location) works correctly."""
        request = CreateParagraphBulletsRequest(
            objectId="table_456",
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=5),
            bulletPreset=BulletGlyphPreset.BULLET_ARROW_DIAMOND_DISC,
            cellLocation=TableCellLocation(rowIndex=0, columnIndex=1),
        )
        assert request.objectId == "table_456"
        assert request.textRange.type == RangeType.FIXED_RANGE
        assert request.textRange.startIndex == 0
        assert request.textRange.endIndex == 5
        assert request.bulletPreset == BulletGlyphPreset.BULLET_ARROW_DIAMOND_DISC
        assert request.cellLocation.rowIndex == 0
        assert request.cellLocation.columnIndex == 1

    def test_api_format_conversion(self):
        """Test that to_api_format() works correctly."""
        request = CreateParagraphBulletsRequest(
            objectId="shape_123",
            textRange=Range(type=RangeType.ALL),
            bulletPreset=BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE,
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "shape_123",
            "textRange": {"type": "ALL"},
            "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE",
        }
        assert api_format == expected

    def test_api_format_with_cell_location(self):
        """Test that to_api_format() includes cellLocation when present."""
        request = CreateParagraphBulletsRequest(
            objectId="table_456",
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=5),
            bulletPreset=BulletGlyphPreset.BULLET_ARROW_DIAMOND_DISC,
            cellLocation=TableCellLocation(rowIndex=1, columnIndex=2),
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "table_456",
            "textRange": {"type": "FIXED_RANGE", "startIndex": 0, "endIndex": 5},
            "bulletPreset": "BULLET_ARROW_DIAMOND_DISC",
            "cellLocation": {"rowIndex": 1, "columnIndex": 2},
        }
        assert api_format == expected

    def test_all_bullet_presets_valid(self):
        """Test that all bullet presets from the enum work correctly."""
        for preset in BulletGlyphPreset:
            request = CreateParagraphBulletsRequest(
                objectId="test_shape", textRange=Range(type=RangeType.ALL), bulletPreset=preset
            )
            assert request.bulletPreset == preset


class TestInsertTextRequest:
    """Test cases for the InsertTextRequest class."""

    def test_shape_request_valid(self):
        """Test that request for shape (without cell location) works correctly."""
        request = InsertTextRequest(
            objectId="shape_123",
            text="Hello, World!",
            insertionIndex=0,
        )
        assert request.objectId == "shape_123"
        assert request.text == "Hello, World!"
        assert request.insertionIndex == 0
        assert request.cellLocation is None

    def test_table_request_valid(self):
        """Test that request for table (with cell location) works correctly."""
        request = InsertTextRequest(
            objectId="table_456",
            text="Cell text",
            insertionIndex=5,
            cellLocation=TableCellLocation(rowIndex=1, columnIndex=2),
        )
        assert request.objectId == "table_456"
        assert request.text == "Cell text"
        assert request.insertionIndex == 5
        assert request.cellLocation.rowIndex == 1
        assert request.cellLocation.columnIndex == 2

    def test_api_format_conversion(self):
        """Test that to_api_format() works correctly."""
        request = InsertTextRequest(
            objectId="shape_123",
            text="Test text",
            insertionIndex=10,
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "shape_123",
            "text": "Test text",
            "insertionIndex": 10,
        }
        assert api_format == expected

    def test_api_format_with_cell_location(self):
        """Test that to_api_format() includes cellLocation when present."""
        request = InsertTextRequest(
            objectId="table_456",
            text="Table cell text",
            insertionIndex=0,
            cellLocation=TableCellLocation(rowIndex=0, columnIndex=1),
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "table_456",
            "text": "Table cell text",
            "insertionIndex": 0,
            "cellLocation": {"rowIndex": 0, "columnIndex": 1},
        }
        assert api_format == expected

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = InsertTextRequest(
            objectId="shape_123",
            text="Test",
            insertionIndex=0,
        )
        request_format = request.to_request()

        expected = [
            {
                "insertText": {
                    "objectId": "shape_123",
                    "text": "Test",
                    "insertionIndex": 0,
                }
            }
        ]
        assert request_format == expected


class TestUpdateTextStyleRequest:
    """Test cases for the UpdateTextStyleRequest class."""

    def test_shape_request_valid(self):
        """Test that request for shape (without cell location) works correctly."""
        style = TextStyle(bold=True, italic=False)
        request = UpdateTextStyleRequest(
            objectId="shape_123",
            style=style,
            textRange=Range(type=RangeType.ALL),
            fields="bold,italic",
        )
        assert request.objectId == "shape_123"
        assert request.style.bold is True
        assert request.style.italic is False
        assert request.textRange.type == RangeType.ALL
        assert request.fields == "bold,italic"
        assert request.cellLocation is None

    def test_table_request_valid(self):
        """Test that request for table (with cell location) works correctly."""
        style = TextStyle(underline=True, fontFamily="Arial")
        request = UpdateTextStyleRequest(
            objectId="table_456",
            style=style,
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=10),
            fields="underline,fontFamily",
            cellLocation=TableCellLocation(rowIndex=1, columnIndex=2),
        )
        assert request.objectId == "table_456"
        assert request.style.underline is True
        assert request.style.fontFamily == "Arial"
        assert request.textRange.type == RangeType.FIXED_RANGE
        assert request.textRange.startIndex == 0
        assert request.textRange.endIndex == 10
        assert request.fields == "underline,fontFamily"
        assert request.cellLocation.rowIndex == 1
        assert request.cellLocation.columnIndex == 2

    def test_api_format_conversion(self):
        """Test that to_api_format() works correctly."""
        style = TextStyle(bold=True)
        request = UpdateTextStyleRequest(
            objectId="shape_123",
            style=style,
            textRange=Range(type=RangeType.ALL),
            fields="bold",
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "shape_123",
            "style": {"bold": True},
            "textRange": {"type": "ALL"},
            "fields": "bold",
        }
        assert api_format == expected

    def test_api_format_with_cell_location(self):
        """Test that to_api_format() includes cellLocation when present."""
        style = TextStyle(italic=True, strikethrough=False)
        request = UpdateTextStyleRequest(
            objectId="table_456",
            style=style,
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=5, endIndex=15),
            fields="italic,strikethrough",
            cellLocation=TableCellLocation(rowIndex=0, columnIndex=1),
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "table_456",
            "style": {"italic": True, "strikethrough": False},
            "textRange": {"type": "FIXED_RANGE", "startIndex": 5, "endIndex": 15},
            "fields": "italic,strikethrough",
            "cellLocation": {"rowIndex": 0, "columnIndex": 1},
        }
        assert api_format == expected

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        style = TextStyle(bold=True)
        request = UpdateTextStyleRequest(
            objectId="shape_123",
            style=style,
            textRange=Range(type=RangeType.ALL),
            fields="bold",
        )
        request_format = request.to_request()

        expected = [
            {
                "updateTextStyle": {
                    "objectId": "shape_123",
                    "style": {"bold": True},
                    "textRange": {"type": "ALL"},
                    "fields": "bold",
                }
            }
        ]
        assert request_format == expected

    def test_wildcard_fields(self):
        """Test that wildcard fields work correctly."""
        style = TextStyle(bold=True, italic=True, underline=True)
        request = UpdateTextStyleRequest(
            objectId="shape_123",
            style=style,
            textRange=Range(type=RangeType.ALL),
            fields="*",
        )
        assert request.fields == "*"


class TestDeleteTextRequest:
    """Test cases for the DeleteTextRequest class."""

    def test_shape_request_valid(self):
        """Test that request for shape (without cell location) works correctly."""
        request = DeleteTextRequest(
            objectId="shape_123",
            textRange=Range(type=RangeType.ALL),
        )
        assert request.objectId == "shape_123"
        assert request.textRange.type == RangeType.ALL
        assert request.cellLocation is None

    def test_table_request_valid(self):
        """Test that request for table (with cell location) works correctly."""
        request = DeleteTextRequest(
            objectId="table_456",
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=10),
            cellLocation=TableCellLocation(rowIndex=1, columnIndex=2),
        )
        assert request.objectId == "table_456"
        assert request.textRange.type == RangeType.FIXED_RANGE
        assert request.textRange.startIndex == 0
        assert request.textRange.endIndex == 10
        assert request.cellLocation.rowIndex == 1
        assert request.cellLocation.columnIndex == 2

    def test_api_format_conversion(self):
        """Test that to_api_format() works correctly."""
        request = DeleteTextRequest(
            objectId="shape_123",
            textRange=Range(type=RangeType.ALL),
        )
        api_format = request.to_api_format()

        expected = {
            "objectId": "shape_123",
            "textRange": {"type": "ALL"},
        }
        assert api_format == expected

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = DeleteTextRequest(
            objectId="shape_123",
            textRange=Range(type=RangeType.ALL),
        )
        request_format = request.to_request()

        expected = [
            {
                "deleteText": {
                    "objectId": "shape_123",
                    "textRange": {"type": "ALL"},
                }
            }
        ]
        assert request_format == expected


class TestCreateShapeRequest:
    """Test cases for the CreateShapeRequest class."""

    def test_basic_shape_request_valid(self):
        """Test that basic shape request works correctly."""
        element_props = ElementProperties(
            pageObjectId="slide_123",
            size={"width": {"magnitude": 100, "unit": "PT"}, "height": {"magnitude": 50, "unit": "PT"}},
            transform={"scaleX": 1, "scaleY": 1, "translateX": 10, "translateY": 20, "unit": "PT"}
        )
        request = CreateShapeRequest(
            elementProperties=element_props,
            shapeType=ShapeType.TEXT_BOX,
        )
        assert request.elementProperties.pageObjectId == "slide_123"
        assert request.shapeType == ShapeType.TEXT_BOX
        assert request.objectId is None

    def test_shape_request_with_object_id(self):
        """Test that shape request with custom object ID works correctly."""
        element_props = ElementProperties(
            pageObjectId="slide_123",
            size={"width": {"magnitude": 100, "unit": "PT"}, "height": {"magnitude": 50, "unit": "PT"}},
            transform={"scaleX": 1, "scaleY": 1, "translateX": 10, "translateY": 20, "unit": "PT"}
        )
        request = CreateShapeRequest(
            objectId="custom_shape_id",
            elementProperties=element_props,
            shapeType=ShapeType.RECTANGLE,
        )
        assert request.objectId == "custom_shape_id"
        assert request.shapeType == ShapeType.RECTANGLE

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        element_props = ElementProperties(
            pageObjectId="slide_123",
            size={"width": {"magnitude": 100, "unit": "PT"}, "height": {"magnitude": 50, "unit": "PT"}},
            transform={"scaleX": 1, "scaleY": 1, "translateX": 10, "translateY": 20, "unit": "PT"}
        )
        request = CreateShapeRequest(
            objectId="shape_456",
            elementProperties=element_props,
            shapeType=ShapeType.ELLIPSE,
        )
        request_format = request.to_request()

        expected = [
            {
                "createShape": {
                    "objectId": "shape_456",
                    "elementProperties": {
                        "pageObjectId": "slide_123",
                        "size": {"width": {"magnitude": 100, "unit": "PT"}, "height": {"magnitude": 50, "unit": "PT"}},
                        "transform": {"scaleX": 1, "scaleY": 1, "translateX": 10, "translateY": 20, "unit": "PT"}
                    },
                    "shapeType": "ELLIPSE",
                }
            }
        ]
        assert request_format == expected


class TestReplaceImageRequest:
    """Test cases for the ReplaceImageRequest class."""

    def test_basic_replace_image_request(self):
        """Test that basic replace image request works correctly."""
        request = ReplaceImageRequest(
            imageObjectId="image_123",
            url="https://example.com/image.jpg",
        )
        assert request.imageObjectId == "image_123"
        assert request.url == "https://example.com/image.jpg"
        assert request.imageReplaceMethod == "CENTER_INSIDE"

    def test_replace_image_with_custom_method(self):
        """Test that replace image request with custom method works correctly."""
        request = ReplaceImageRequest(
            imageObjectId="image_456",
            url="https://example.com/image.png",
            imageReplaceMethod="CENTER_CROP",
        )
        assert request.imageObjectId == "image_456"
        assert request.url == "https://example.com/image.png"
        assert request.imageReplaceMethod == "CENTER_CROP"

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = ReplaceImageRequest(
            imageObjectId="image_789",
            url="https://example.com/image.gif",
        )
        request_format = request.to_request()

        expected = [
            {
                "replaceImage": {
                    "imageObjectId": "image_789",
                    "url": "https://example.com/image.gif",
                    "imageReplaceMethod": "CENTER_INSIDE",
                }
            }
        ]
        assert request_format == expected


class TestCreateSlideRequest:
    """Test cases for the CreateSlideRequest class."""

    def test_basic_slide_request(self):
        """Test that basic slide request works correctly."""
        request = CreateSlideRequest()
        assert request.objectId is None
        assert request.insertionIndex is None
        assert request.slideLayoutReference is None
        assert request.placeholderIdMappings is None

    def test_slide_request_with_layout_reference(self):
        """Test that slide request with layout reference works correctly."""
        layout_ref = LayoutReference(predefinedLayout=PredefinedLayout.TITLE_AND_BODY)
        request = CreateSlideRequest(
            objectId="slide_123",
            insertionIndex=2,
            slideLayoutReference=layout_ref,
        )
        assert request.objectId == "slide_123"
        assert request.insertionIndex == 2
        assert request.slideLayoutReference.predefinedLayout == PredefinedLayout.TITLE_AND_BODY

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        layout_ref = LayoutReference(predefinedLayout=PredefinedLayout.BLANK)
        request = CreateSlideRequest(
            objectId="slide_456",
            insertionIndex=0,
            slideLayoutReference=layout_ref,
        )
        request_format = request.to_request()

        expected = [
            {
                "createSlide": {
                    "objectId": "slide_456",
                    "insertionIndex": 0,
                    "slideLayoutReference": {"predefinedLayout": "BLANK"},
                }
            }
        ]
        assert request_format == expected


class TestUpdateSlidesPositionRequest:
    """Test cases for the UpdateSlidesPositionRequest class."""

    def test_basic_position_update_request(self):
        """Test that basic position update request works correctly."""
        request = UpdateSlidesPositionRequest(
            slideObjectIds=["slide_1", "slide_2"],
            insertionIndex=3,
        )
        assert request.slideObjectIds == ["slide_1", "slide_2"]
        assert request.insertionIndex == 3

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = UpdateSlidesPositionRequest(
            slideObjectIds=["slide_a", "slide_b", "slide_c"],
            insertionIndex=1,
        )
        request_format = request.to_request()

        expected = [
            {
                "updateSlidesPosition": {
                    "slideObjectIds": ["slide_a", "slide_b", "slide_c"],
                    "insertionIndex": 1,
                }
            }
        ]
        assert request_format == expected


class TestDeleteObjectRequest:
    """Test cases for the DeleteObjectRequest class."""

    def test_basic_delete_request(self):
        """Test that basic delete request works correctly."""
        request = DeleteObjectRequest(objectId="object_123")
        assert request.objectId == "object_123"

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = DeleteObjectRequest(objectId="slide_456")
        request_format = request.to_request()

        expected = [
            {
                "deleteObject": {
                    "objectId": "slide_456",
                }
            }
        ]
        assert request_format == expected


class TestDuplicateObjectRequest:
    """Test cases for the DuplicateObjectRequest class."""

    def test_basic_duplicate_request(self):
        """Test that basic duplicate request works correctly."""
        request = DuplicateObjectRequest(objectId="object_123")
        assert request.objectId == "object_123"
        assert request.objectIds is None

    def test_duplicate_request_with_id_mapping(self):
        """Test that duplicate request with ID mapping works correctly."""
        id_mapping = {"old_id_1": "new_id_1", "old_id_2": "new_id_2"}
        request = DuplicateObjectRequest(
            objectId="slide_456",
            objectIds=id_mapping,
        )
        assert request.objectId == "slide_456"
        assert request.objectIds == id_mapping

    def test_to_request_format(self):
        """Test that to_request() returns correct format."""
        request = DuplicateObjectRequest(objectId="slide_789")
        request_format = request.to_request()

        expected = [
            {
                "duplicateObject": {
                    "objectId": "slide_789",
                }
            }
        ]
        assert request_format == expected
