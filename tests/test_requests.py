"""
Tests for the requests module, specifically CreateParagraphBulletsRequest and related classes.
"""

import pytest
from pydantic import ValidationError

from gslides_api.requests import CreateParagraphBulletsRequest, Range, RangeType, TableCellLocation
from gslides_api.domain import BulletGlyphPreset


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
        with pytest.raises(ValidationError, match="startIndex and endIndex must be None when type is ALL"):
            Range(type=RangeType.ALL, startIndex=0)
    
    def test_fixed_range_missing_end_index_invalid(self):
        """Test that FIXED_RANGE without endIndex raises validation error."""
        with pytest.raises(ValidationError, match="Both startIndex and endIndex must be provided when type is FIXED_RANGE"):
            Range(type=RangeType.FIXED_RANGE, startIndex=0)
    
    def test_fixed_range_missing_start_index_invalid(self):
        """Test that FIXED_RANGE without startIndex raises validation error."""
        with pytest.raises(ValidationError, match="Both startIndex and endIndex must be provided when type is FIXED_RANGE"):
            Range(type=RangeType.FIXED_RANGE, endIndex=10)
    
    def test_fixed_range_start_greater_than_end_invalid(self):
        """Test that FIXED_RANGE with startIndex >= endIndex raises validation error."""
        with pytest.raises(ValidationError, match="startIndex must be less than endIndex"):
            Range(type=RangeType.FIXED_RANGE, startIndex=5, endIndex=3)
        
        with pytest.raises(ValidationError, match="startIndex must be less than endIndex"):
            Range(type=RangeType.FIXED_RANGE, startIndex=5, endIndex=5)
    
    def test_from_start_index_missing_start_invalid(self):
        """Test that FROM_START_INDEX without startIndex raises validation error."""
        with pytest.raises(ValidationError, match="startIndex must be provided when type is FROM_START_INDEX"):
            Range(type=RangeType.FROM_START_INDEX)
    
    def test_from_start_index_with_end_index_invalid(self):
        """Test that FROM_START_INDEX with endIndex raises validation error."""
        with pytest.raises(ValidationError, match="endIndex must be None when type is FROM_START_INDEX"):
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
            bulletPreset=BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE
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
            cellLocation=TableCellLocation(rowIndex=0, columnIndex=1)
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
            bulletPreset=BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE
        )
        api_format = request.to_api_format()
        
        expected = {
            "objectId": "shape_123",
            "textRange": {"type": "ALL"},
            "bulletPreset": "BULLET_DISC_CIRCLE_SQUARE"
        }
        assert api_format == expected
    
    def test_api_format_with_cell_location(self):
        """Test that to_api_format() includes cellLocation when present."""
        request = CreateParagraphBulletsRequest(
            objectId="table_456",
            textRange=Range(type=RangeType.FIXED_RANGE, startIndex=0, endIndex=5),
            bulletPreset=BulletGlyphPreset.BULLET_ARROW_DIAMOND_DISC,
            cellLocation=TableCellLocation(rowIndex=1, columnIndex=2)
        )
        api_format = request.to_api_format()
        
        expected = {
            "objectId": "table_456",
            "textRange": {
                "type": "FIXED_RANGE",
                "startIndex": 0,
                "endIndex": 5
            },
            "bulletPreset": "BULLET_ARROW_DIAMOND_DISC",
            "cellLocation": {
                "rowIndex": 1,
                "columnIndex": 2
            }
        }
        assert api_format == expected
    
    def test_all_bullet_presets_valid(self):
        """Test that all bullet presets from the enum work correctly."""
        for preset in BulletGlyphPreset:
            request = CreateParagraphBulletsRequest(
                objectId="test_shape",
                textRange=Range(type=RangeType.ALL),
                bulletPreset=preset
            )
            assert request.bulletPreset == preset
