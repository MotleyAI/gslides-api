"""Test script to verify the discriminated union functionality."""

import pytest
from pydantic import TypeAdapter, ValidationError

from gslides_api.domain.domain import Image, Size, Transform, Video, VideoSourceType
from gslides_api.domain.table import Table
from gslides_api.element.element import PageElement, VideoElement
from gslides_api.element.image import ImageElement
from gslides_api.element.table import TableElement


def test_discriminated_union():
    """Test that the discriminated union correctly instantiates the right subclass."""

    # Create a TypeAdapter for the PageElement union
    page_element_adapter = TypeAdapter(PageElement)

    # Test data for different element types
    base_data = {
        "objectId": "test_id",
        "size": {"width": 100, "height": 100},
        "transform": {"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1},
        "title": "Test Element",
    }

    # Test TableElement
    table_data = {**base_data, "table": {"rows": 3, "columns": 4}}

    print("Testing TableElement...")
    table_element = page_element_adapter.validate_python(table_data)
    print(f"Type: {type(table_element)}")
    print(f"Is TableElement: {isinstance(table_element, TableElement)}")
    print(f"Table rows: {table_element.table.rows}")
    assert isinstance(table_element, TableElement)
    assert table_element.table.rows == 3
    print()

    # Test ImageElement
    image_data = {**base_data, "image": {"contentUrl": "https://example.com/image.jpg"}}

    print("Testing ImageElement...")
    image_element = page_element_adapter.validate_python(image_data)
    print(f"Type: {type(image_element)}")
    print(f"Is ImageElement: {isinstance(image_element, ImageElement)}")
    print(f"Image URL: {image_element.image.contentUrl}")
    assert isinstance(image_element, ImageElement)
    assert image_element.image.contentUrl == "https://example.com/image.jpg"
    print()

    # Test VideoElement
    video_data = {
        **base_data,
        "video": {
            "source": "YOUTUBE",
            "id": "dQw4w9WgXcQ",
            "url": "https://www.youtube.com/watch?v=dQw4w9WgXcQ",
        },
    }

    print("Testing VideoElement...")
    video_element = page_element_adapter.validate_python(video_data)
    print(f"Type: {type(video_element)}")
    print(f"Is VideoElement: {isinstance(video_element, VideoElement)}")
    print(f"Video source: {video_element.video.source}")
    print(f"Video ID: {video_element.video.id}")
    assert isinstance(video_element, VideoElement)
    assert video_element.video.id == "dQw4w9WgXcQ"
    print()

    print("\nAll tests completed!")


def test_table_element_instantiation():
    """Test direct TableElement instantiation."""
    table_element = TableElement(
        objectId="table_1",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        table=Table(rows=3, columns=4),
    )
    assert isinstance(table_element, TableElement)
    assert table_element.table.rows == 3
    assert table_element.table.columns == 4


def test_image_element_instantiation():
    """Test direct ImageElement instantiation."""
    image_element = ImageElement(
        objectId="image_1",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        image=Image(contentUrl="https://example.com/image.jpg"),
    )
    assert isinstance(image_element, ImageElement)
    assert image_element.image.contentUrl == "https://example.com/image.jpg"


def test_video_element_instantiation():
    """Test direct VideoElement instantiation."""
    video_element = VideoElement(
        objectId="video_1",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        video=Video(source=VideoSourceType.YOUTUBE, id="dQw4w9WgXcQ"),
    )
    assert isinstance(video_element, VideoElement)
    assert video_element.video.source == VideoSourceType.YOUTUBE
    assert video_element.video.id == "dQw4w9WgXcQ"


def test_discriminated_union_validation_error():
    """Test that validation fails when no element type is specified."""
    page_element_adapter = TypeAdapter(PageElement)

    base_data = {
        "objectId": "test_id",
        "size": {"width": 100, "height": 100},
        "transform": {"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1},
    }

    with pytest.raises(ValidationError) as exc_info:
        page_element_adapter.validate_python(base_data)

    # Check that the error is related to discriminator
    assert "discriminator" in str(exc_info.value).lower() or "tag" in str(exc_info.value).lower()
