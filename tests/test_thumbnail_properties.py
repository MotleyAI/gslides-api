"""Tests for ThumbnailProperties, ThumbnailSize, and MimeType classes."""

import pytest
from gslides_api.domain import ThumbnailProperties, ThumbnailSize, MimeType


class TestMimeType:
    """Test MimeType enum."""

    def test_mime_type_values(self):
        """Test that MimeType has the expected values."""
        assert MimeType.PNG.value == "PNG"
        
    def test_mime_type_enum_members(self):
        """Test that MimeType has exactly one member."""
        assert len(list(MimeType)) == 1
        assert MimeType.PNG in MimeType


class TestThumbnailSize:
    """Test ThumbnailSize enum."""

    def test_thumbnail_size_values(self):
        """Test that ThumbnailSize has the expected values."""
        assert ThumbnailSize.THUMBNAIL_SIZE_UNSPECIFIED.value == "THUMBNAIL_SIZE_UNSPECIFIED"
        assert ThumbnailSize.LARGE.value == "LARGE"
        assert ThumbnailSize.MEDIUM.value == "MEDIUM"
        assert ThumbnailSize.SMALL.value == "SMALL"
        
    def test_thumbnail_size_enum_members(self):
        """Test that ThumbnailSize has exactly four members."""
        assert len(list(ThumbnailSize)) == 4
        expected_members = {
            ThumbnailSize.THUMBNAIL_SIZE_UNSPECIFIED,
            ThumbnailSize.LARGE,
            ThumbnailSize.MEDIUM,
            ThumbnailSize.SMALL
        }
        assert set(ThumbnailSize) == expected_members


class TestThumbnailProperties:
    """Test ThumbnailProperties class."""

    def test_thumbnail_properties_default_creation(self):
        """Test creating ThumbnailProperties with default values."""
        props = ThumbnailProperties()
        assert props.mimeType is None
        assert props.thumbnailSize is None

    def test_thumbnail_properties_with_mime_type(self):
        """Test creating ThumbnailProperties with mimeType."""
        props = ThumbnailProperties(mimeType=MimeType.PNG)
        assert props.mimeType == MimeType.PNG
        assert props.thumbnailSize is None

    def test_thumbnail_properties_with_thumbnail_size(self):
        """Test creating ThumbnailProperties with thumbnailSize."""
        props = ThumbnailProperties(thumbnailSize=ThumbnailSize.LARGE)
        assert props.mimeType is None
        assert props.thumbnailSize == ThumbnailSize.LARGE

    def test_thumbnail_properties_with_both_fields(self):
        """Test creating ThumbnailProperties with both fields."""
        props = ThumbnailProperties(
            mimeType=MimeType.PNG,
            thumbnailSize=ThumbnailSize.MEDIUM
        )
        assert props.mimeType == MimeType.PNG
        assert props.thumbnailSize == ThumbnailSize.MEDIUM

    def test_thumbnail_properties_to_api_format_empty(self):
        """Test converting empty ThumbnailProperties to API format."""
        props = ThumbnailProperties()
        api_format = props.to_api_format()
        assert api_format == {}

    def test_thumbnail_properties_to_api_format_with_mime_type(self):
        """Test converting ThumbnailProperties with mimeType to API format."""
        props = ThumbnailProperties(mimeType=MimeType.PNG)
        api_format = props.to_api_format()
        assert api_format == {"mimeType": "PNG"}

    def test_thumbnail_properties_to_api_format_with_thumbnail_size(self):
        """Test converting ThumbnailProperties with thumbnailSize to API format."""
        props = ThumbnailProperties(thumbnailSize=ThumbnailSize.SMALL)
        api_format = props.to_api_format()
        assert api_format == {"thumbnailSize": "SMALL"}

    def test_thumbnail_properties_to_api_format_with_both_fields(self):
        """Test converting ThumbnailProperties with both fields to API format."""
        props = ThumbnailProperties(
            mimeType=MimeType.PNG,
            thumbnailSize=ThumbnailSize.LARGE
        )
        api_format = props.to_api_format()
        expected = {
            "mimeType": "PNG",
            "thumbnailSize": "LARGE"
        }
        assert api_format == expected

    def test_thumbnail_properties_all_thumbnail_sizes(self):
        """Test ThumbnailProperties with all possible thumbnail sizes."""
        sizes = [
            ThumbnailSize.THUMBNAIL_SIZE_UNSPECIFIED,
            ThumbnailSize.LARGE,
            ThumbnailSize.MEDIUM,
            ThumbnailSize.SMALL
        ]
        
        for size in sizes:
            props = ThumbnailProperties(thumbnailSize=size)
            api_format = props.to_api_format()
            assert api_format == {"thumbnailSize": size.value}

    def test_thumbnail_properties_inheritance(self):
        """Test that ThumbnailProperties inherits from GSlidesBaseModel."""
        from gslides_api.domain import GSlidesBaseModel
        props = ThumbnailProperties()
        assert isinstance(props, GSlidesBaseModel)
        assert hasattr(props, 'to_api_format')
