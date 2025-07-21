"""Test AffineTransform and Unit classes."""

import pytest
from gslides_api.domain import AffineTransform, Unit


class TestUnit:
    """Test the Unit enum."""

    def test_unit_values(self):
        """Test that Unit enum has correct values."""
        assert Unit.UNIT_UNSPECIFIED.value == "UNIT_UNSPECIFIED"
        assert Unit.EMU.value == "EMU"
        assert Unit.PT.value == "PT"

    def test_unit_string_representation(self):
        """Test string representation of Unit enum values."""
        assert str(Unit.UNIT_UNSPECIFIED) == "Unit.UNIT_UNSPECIFIED"
        assert str(Unit.EMU) == "Unit.EMU"
        assert str(Unit.PT) == "Unit.PT"


class TestAffineTransform:
    """Test the AffineTransform class."""

    def test_affine_transform_creation_with_all_fields(self):
        """Test creating an AffineTransform with all fields."""
        transform = AffineTransform(
            scaleX=1.5,
            scaleY=2.0,
            shearX=0.1,
            shearY=0.2,
            translateX=100.0,
            translateY=200.0,
            unit=Unit.EMU
        )
        
        assert transform.scaleX == 1.5
        assert transform.scaleY == 2.0
        assert transform.shearX == 0.1
        assert transform.shearY == 0.2
        assert transform.translateX == 100.0
        assert transform.translateY == 200.0
        assert transform.unit == Unit.EMU

    def test_affine_transform_creation_without_unit(self):
        """Test creating an AffineTransform without unit (should default to None)."""
        transform = AffineTransform(
            scaleX=1.0,
            scaleY=1.0,
            shearX=0.0,
            shearY=0.0,
            translateX=0.0,
            translateY=0.0
        )
        
        assert transform.scaleX == 1.0
        assert transform.scaleY == 1.0
        assert transform.shearX == 0.0
        assert transform.shearY == 0.0
        assert transform.translateX == 0.0
        assert transform.translateY == 0.0
        assert transform.unit is None

    def test_affine_transform_api_format(self):
        """Test converting AffineTransform to API format."""
        transform = AffineTransform(
            scaleX=1.5,
            scaleY=2.0,
            shearX=0.1,
            shearY=0.2,
            translateX=100.0,
            translateY=200.0,
            unit=Unit.EMU
        )
        
        api_format = transform.to_api_format()
        
        expected = {
            "scaleX": 1.5,
            "scaleY": 2.0,
            "shearX": 0.1,
            "shearY": 0.2,
            "translateX": 100.0,
            "translateY": 200.0,
            "unit": "EMU"
        }
        
        assert api_format == expected

    def test_affine_transform_api_format_without_unit(self):
        """Test converting AffineTransform to API format when unit is None."""
        transform = AffineTransform(
            scaleX=1.0,
            scaleY=1.0,
            shearX=0.0,
            shearY=0.0,
            translateX=0.0,
            translateY=0.0
        )
        
        api_format = transform.to_api_format()
        
        expected = {
            "scaleX": 1.0,
            "scaleY": 1.0,
            "shearX": 0.0,
            "shearY": 0.0,
            "translateX": 0.0,
            "translateY": 0.0
        }
        
        # unit should be excluded when None due to exclude_none=True
        assert api_format == expected
        assert "unit" not in api_format

    def test_transformation_mathematics(self):
        """Test the mathematical transformation described in the docstring."""
        transform = AffineTransform(
            scaleX=2.0,
            scaleY=3.0,
            shearX=0.5,
            shearY=0.2,
            translateX=10.0,
            translateY=20.0
        )
        
        # Test point (5, 4)
        x, y = 5.0, 4.0
        
        # Apply transformation as described in docstring:
        # x' = scaleX * x + shearX * y + translateX
        # y' = scaleY * y + shearY * x + translateY
        x_prime = transform.scaleX * x + transform.shearX * y + transform.translateX
        y_prime = transform.scaleY * y + transform.shearY * x + transform.translateY
        
        # Expected: x' = 2.0 * 5 + 0.5 * 4 + 10 = 10 + 2 + 10 = 22
        # Expected: y' = 3.0 * 4 + 0.2 * 5 + 20 = 12 + 1 + 20 = 33
        expected_x = 22.0
        expected_y = 33.0
        
        assert abs(x_prime - expected_x) < 0.001
        assert abs(y_prime - expected_y) < 0.001

    def test_identity_transformation(self):
        """Test identity transformation (no change)."""
        transform = AffineTransform(
            scaleX=1.0,
            scaleY=1.0,
            shearX=0.0,
            shearY=0.0,
            translateX=0.0,
            translateY=0.0
        )
        
        # Test multiple points
        test_points = [(0, 0), (1, 1), (5, 3), (-2, 4)]
        
        for x, y in test_points:
            x_prime = transform.scaleX * x + transform.shearX * y + transform.translateX
            y_prime = transform.scaleY * y + transform.shearY * x + transform.translateY
            
            assert abs(x_prime - x) < 0.001
            assert abs(y_prime - y) < 0.001

    def test_required_fields(self):
        """Test that all required fields must be provided."""
        from pydantic import ValidationError

        # This should work
        AffineTransform(
            scaleX=1.0,
            scaleY=1.0,
            shearX=0.0,
            shearY=0.0,
            translateX=0.0,
            translateY=0.0
        )

        # Missing required fields should raise validation error
        with pytest.raises(ValidationError):
            AffineTransform()

        with pytest.raises(ValidationError):
            AffineTransform(scaleX=1.0)  # Missing other required fields
