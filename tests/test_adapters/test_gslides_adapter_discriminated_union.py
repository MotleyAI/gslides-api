"""Test discriminated union functionality for ConcreteElement in gslides_adapter."""

from unittest.mock import MagicMock, Mock

import pytest
from pydantic import TypeAdapter, ValidationError

from gslides_api.element.base import ElementKind
from gslides_api.element.element import PageElement
from gslides_api.element.image import ImageElement
from gslides_api.element.shape import ShapeElement
from gslides_api.element.table import TableElement

from gslides_api.adapters.gslides_adapter import (
    GSlidesElement,
    GSlidesElementParent,
    GSlidesImageElement,
    GSlidesShapeElement,
    GSlidesTableElement,
    concrete_element_discriminator,
)


class TestConcreteElementDiscriminator:
    """Test the discriminator function for ConcreteElement."""

    def test_discriminator_with_shape_element(self):
        """Test discriminator correctly identifies shape elements."""
        mock_element = Mock()
        mock_element.type = ElementKind.SHAPE

        result = concrete_element_discriminator(mock_element)
        assert result == "shape"

    def test_discriminator_with_image_element(self):
        """Test discriminator correctly identifies image elements."""
        mock_element = Mock()
        mock_element.type = ElementKind.IMAGE

        result = concrete_element_discriminator(mock_element)
        assert result == "image"

    def test_discriminator_with_table_element(self):
        """Test discriminator correctly identifies table elements."""
        mock_element = Mock()
        mock_element.type = ElementKind.TABLE

        result = concrete_element_discriminator(mock_element)
        assert result == "table"

    def test_discriminator_with_enum_value(self):
        """Test discriminator works with enum values that have .value attribute."""
        mock_element = Mock()
        mock_enum = Mock()
        mock_enum.value = "SHAPE"
        mock_element.type = mock_enum

        result = concrete_element_discriminator(mock_element)
        assert result == "shape"

    def test_discriminator_with_string_types(self):
        """Test discriminator works with string type values."""
        mock_element = Mock()
        mock_element.type = "shape"

        result = concrete_element_discriminator(mock_element)
        assert result == "shape"

        mock_element.type = "IMAGE"
        result = concrete_element_discriminator(mock_element)
        assert result == "image"

        mock_element.type = "table"
        result = concrete_element_discriminator(mock_element)
        assert result == "table"

    def test_discriminator_with_invalid_type(self):
        """Test discriminator returns 'generic' for unsupported types."""
        mock_element = Mock()
        mock_element.type = "unsupported_type"

        result = concrete_element_discriminator(mock_element)
        assert result == "generic"

    def test_discriminator_without_type_attribute(self):
        """Test discriminator raises error when element has no type attribute."""
        mock_element = Mock(spec=[])  # Mock without type attribute

        with pytest.raises(ValueError, match="Cannot determine element type"):
            concrete_element_discriminator(mock_element)


class TestConcreteElementValidation:
    """Test ConcreteElement discriminated union validation."""

    def setup_method(self):
        """Set up TypeAdapter for ConcreteElement."""
        self.adapter = TypeAdapter(GSlidesElement)

    def create_mock_shape_element(self):
        """Create a mock ShapeElement for testing."""
        mock_element = Mock(spec=ShapeElement)
        mock_element.type = ElementKind.SHAPE
        mock_element.objectId = "shape_123"
        mock_element.presentation_id = "pres_123"
        mock_element.slide_id = "slide_123"
        # Create alt_text with string title and description, not Mock
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Shape"
        alt_text_mock.description = None
        mock_element.alt_text = alt_text_mock
        return mock_element

    def create_mock_image_element(self):
        """Create a mock ImageElement for testing."""
        mock_element = Mock(spec=ImageElement)
        mock_element.type = ElementKind.IMAGE
        mock_element.objectId = "image_123"
        mock_element.presentation_id = "pres_123"
        mock_element.slide_id = "slide_123"
        # Create alt_text with string title and description, not Mock
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Image"
        alt_text_mock.description = None
        mock_element.alt_text = alt_text_mock
        return mock_element

    def create_mock_table_element(self):
        """Create a mock TableElement for testing."""
        mock_element = Mock(spec=TableElement)
        mock_element.type = ElementKind.TABLE
        mock_element.objectId = "table_123"
        mock_element.presentation_id = "pres_123"
        mock_element.slide_id = "slide_123"
        # Create alt_text with string title and description, not Mock
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Table"
        alt_text_mock.description = None
        mock_element.alt_text = alt_text_mock
        return mock_element

    def test_concrete_element_validates_shape(self):
        """Test that ConcreteElement validates and creates ConcreteShapeElement for shape elements."""
        mock_shape = self.create_mock_shape_element()

        result = self.adapter.validate_python(mock_shape)

        assert isinstance(result, GSlidesShapeElement)
        assert result.objectId == "shape_123"
        assert result.presentation_id == "pres_123"
        assert result.slide_id == "slide_123"
        assert result.alt_text.title == "Test Shape"

    def test_concrete_element_validates_image(self):
        """Test that ConcreteElement validates and creates ConcreteImageElement for image elements."""
        mock_image = self.create_mock_image_element()

        result = self.adapter.validate_python(mock_image)

        assert isinstance(result, GSlidesImageElement)
        assert result.objectId == "image_123"
        assert result.presentation_id == "pres_123"
        assert result.slide_id == "slide_123"
        assert result.alt_text.title == "Test Image"

    def test_concrete_element_validates_table(self):
        """Test that ConcreteElement validates and creates ConcreteTableElement for table elements."""
        mock_table = self.create_mock_table_element()

        result = self.adapter.validate_python(mock_table)

        assert isinstance(result, GSlidesTableElement)
        assert result.objectId == "table_123"
        assert result.presentation_id == "pres_123"
        assert result.slide_id == "slide_123"
        assert result.alt_text.title == "Test Table"

    def test_concrete_element_validation_unsupported_type(self):
        """Test that ConcreteElement validation creates ConcreteGenericElement for unsupported element types."""
        mock_element = Mock()
        mock_element.type = "unsupported_type"
        mock_element.objectId = "unsupported_123"
        mock_element.presentation_id = "pres_123"
        mock_element.slide_id = "slide_123"
        alt_text_mock = Mock()
        alt_text_mock.title = "Unsupported Element"
        alt_text_mock.description = None
        mock_element.alt_text = alt_text_mock

        result = self.adapter.validate_python(mock_element)

        assert isinstance(result, GSlidesElementParent)
        assert result.objectId == "unsupported_123"
        assert result.presentation_id == "pres_123"
        assert result.slide_id == "slide_123"
        assert result.alt_text.title == "Unsupported Element"

    def test_concrete_element_validation_error_no_type(self):
        """Test that ConcreteElement validation raises error when element has no type."""
        mock_element = Mock(spec=[])  # Mock without type attribute

        with pytest.raises(ValueError) as exc_info:
            self.adapter.validate_python(mock_element)

        # Verify the error is related to discriminator
        error_str = str(exc_info.value).lower()
        assert "cannot determine element type" in error_str


class TestConcreteShapeElementValidator:
    """Test ConcreteShapeElement validator functionality."""

    def test_validator_with_shape_element(self):
        """Test validator works correctly with ShapeElement."""
        mock_shape = Mock(spec=ShapeElement)
        mock_shape.objectId = "shape_123"
        mock_shape.presentation_id = "pres_123"
        mock_shape.slide_id = "slide_123"
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Shape"
        alt_text_mock.description = None
        mock_shape.alt_text = alt_text_mock

        # Test the validator method directly
        result_data = GSlidesShapeElement.convert_from_page_element(mock_shape)

        assert result_data["objectId"] == "shape_123"
        assert result_data["presentation_id"] == "pres_123"
        assert result_data["slide_id"] == "slide_123"
        assert result_data["gslides_element"] == mock_shape

    def test_validator_with_page_element_having_shape(self):
        """Test validator works with PageElement that has shape attribute."""
        mock_page_element = Mock()
        mock_page_element.shape = Mock()  # Has shape attribute
        mock_page_element.objectId = "shape_456"
        mock_page_element.presentation_id = "pres_456"
        mock_page_element.slide_id = "slide_456"
        alt_text_mock = Mock()
        alt_text_mock.title = "Page Element Shape"
        alt_text_mock.description = None
        mock_page_element.alt_text = alt_text_mock

        result_data = GSlidesShapeElement.convert_from_page_element(mock_page_element)

        assert result_data["objectId"] == "shape_456"
        assert result_data["gslides_element"] == mock_page_element

    def test_validator_error_invalid_element(self):
        """Test validator raises error for invalid element types."""
        # Create a mock without the 'shape' attribute using spec
        mock_invalid = Mock(spec=["objectId", "presentation_id", "slide_id", "alt_text"])
        alt_text_mock = Mock()
        alt_text_mock.title = "Invalid Element"
        alt_text_mock.description = None
        mock_invalid.alt_text = alt_text_mock
        mock_invalid.objectId = "invalid_123"
        mock_invalid.presentation_id = "pres_123"
        mock_invalid.slide_id = "slide_123"

        with pytest.raises(ValueError, match="Expected ShapeElement or PageElement with shape"):
            GSlidesShapeElement.convert_from_page_element(mock_invalid)


class TestConcreteImageElementValidator:
    """Test ConcreteImageElement validator functionality."""

    def test_validator_with_image_element(self):
        """Test validator works correctly with ImageElement."""
        mock_image = Mock(spec=ImageElement)
        mock_image.objectId = "image_123"
        mock_image.presentation_id = "pres_123"
        mock_image.slide_id = "slide_123"
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Image"
        alt_text_mock.description = None
        mock_image.alt_text = alt_text_mock

        result_data = GSlidesImageElement.convert_from_page_element(mock_image)

        assert result_data["objectId"] == "image_123"
        assert result_data["presentation_id"] == "pres_123"
        assert result_data["slide_id"] == "slide_123"
        assert result_data["gslides_element"] == mock_image

    def test_validator_with_page_element_having_image(self):
        """Test validator works with PageElement that has image attribute."""
        mock_page_element = Mock()
        mock_page_element.image = Mock()  # Has image attribute
        mock_page_element.objectId = "image_456"
        mock_page_element.presentation_id = "pres_456"
        mock_page_element.slide_id = "slide_456"
        alt_text_mock = Mock()
        alt_text_mock.title = "Page Element Image"
        alt_text_mock.description = None
        mock_page_element.alt_text = alt_text_mock

        result_data = GSlidesImageElement.convert_from_page_element(mock_page_element)

        assert result_data["objectId"] == "image_456"
        assert result_data["gslides_element"] == mock_page_element


class TestConcreteTableElementValidator:
    """Test ConcreteTableElement validator functionality."""

    def test_validator_with_table_element(self):
        """Test validator works correctly with TableElement."""
        mock_table = Mock(spec=TableElement)
        mock_table.objectId = "table_123"
        mock_table.presentation_id = "pres_123"
        mock_table.slide_id = "slide_123"
        alt_text_mock = Mock()
        alt_text_mock.title = "Test Table"
        alt_text_mock.description = None
        mock_table.alt_text = alt_text_mock

        result_data = GSlidesTableElement.convert_from_page_element(mock_table)

        assert result_data["objectId"] == "table_123"
        assert result_data["presentation_id"] == "pres_123"
        assert result_data["slide_id"] == "slide_123"
        assert result_data["gslides_element"] == mock_table

    def test_validator_with_page_element_having_table(self):
        """Test validator works with PageElement that has table attribute."""
        mock_page_element = Mock()
        mock_page_element.table = Mock()  # Has table attribute
        mock_page_element.objectId = "table_456"
        mock_page_element.presentation_id = "pres_456"
        mock_page_element.slide_id = "slide_456"
        alt_text_mock = Mock()
        alt_text_mock.title = "Page Element Table"
        alt_text_mock.description = None
        mock_page_element.alt_text = alt_text_mock

        result_data = GSlidesTableElement.convert_from_page_element(mock_page_element)

        assert result_data["objectId"] == "table_456"
        assert result_data["gslides_element"] == mock_page_element


class TestConcreteElementIntegration:
    """Integration tests for ConcreteElement discriminated union."""

    def setup_method(self):
        """Set up TypeAdapter for ConcreteElement."""
        self.adapter = TypeAdapter(GSlidesElement)

    def test_end_to_end_shape_validation(self):
        """Test complete validation flow for shape elements."""
        mock_shape = Mock(spec=ShapeElement)
        mock_shape.type = ElementKind.SHAPE
        mock_shape.objectId = "shape_end_to_end"
        mock_shape.presentation_id = "pres_end_to_end"
        mock_shape.slide_id = "slide_end_to_end"
        alt_text_mock = Mock()
        alt_text_mock.title = "End to End Shape"
        alt_text_mock.description = None
        mock_shape.alt_text = alt_text_mock

        # This tests the full discriminated union pipeline
        concrete_element = self.adapter.validate_python(mock_shape)

        # Verify it's the right type
        assert isinstance(concrete_element, GSlidesShapeElement)

        # Verify properties are correctly set
        assert concrete_element.objectId == "shape_end_to_end"
        assert concrete_element.presentation_id == "pres_end_to_end"
        assert concrete_element.slide_id == "slide_end_to_end"
        assert concrete_element.alt_text.title == "End to End Shape"

        # Verify the internal gslides element is preserved
        assert concrete_element.gslides_element == mock_shape

    def test_type_preservation_across_validation(self):
        """Test that the discriminated union preserves type information correctly."""
        # Create elements of different types
        mock_shape = Mock(spec=ShapeElement)
        mock_shape.type = ElementKind.SHAPE
        mock_shape.objectId = "shape_type_test"
        mock_shape.presentation_id = "pres_type_test"
        mock_shape.slide_id = "slide_type_test"
        alt_text_mock_shape = Mock()
        alt_text_mock_shape.title = "Shape Type Test"
        alt_text_mock_shape.description = None
        mock_shape.alt_text = alt_text_mock_shape

        mock_image = Mock(spec=ImageElement)
        mock_image.type = ElementKind.IMAGE
        mock_image.objectId = "image_type_test"
        mock_image.presentation_id = "pres_type_test"
        mock_image.slide_id = "slide_type_test"
        alt_text_mock_image = Mock()
        alt_text_mock_image.title = "Image Type Test"
        alt_text_mock_image.description = None
        mock_image.alt_text = alt_text_mock_image

        # Validate both
        shape_result = self.adapter.validate_python(mock_shape)
        image_result = self.adapter.validate_python(mock_image)

        # Verify correct types were created
        assert isinstance(shape_result, GSlidesShapeElement)
        assert isinstance(image_result, GSlidesImageElement)

        # Verify they maintain their type information
        assert shape_result.type == "SHAPE"  # From AbstractShapeElement
        assert image_result.type == "IMAGE"  # From AbstractImageElement
