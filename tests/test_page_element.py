import pytest
from pydantic import TypeAdapter
from gslides_api.element.element import (
    PageElement,
    LineElement,
    WordArtElement,
    SheetsChartElement,
    GroupElement,
    ImageElement,
)
from gslides_api.element.base import PageElementBase
from gslides_api.element.shape import ShapeElement
from gslides_api.domain import (
    Size,
    Transform,
    Line,
    LineProperties,
    WordArt,
    SheetsChart,
    SheetsChartProperties,
    SpeakerSpotlight,
    SpeakerSpotlightProperties,
    Group,
    Video,
    VideoProperties,
    Image,
    ImageProperties,
    Table,
    Dimension,
)
from gslides_api import ShapeProperties
from gslides_api.text import Shape, ShapeType


def test_page_element_base_fields():
    """Test that PageElementBase has all the required fields."""
    # Since PageElementBase is abstract, we'll test with a concrete subclass
    element = ShapeElement(
        objectId="test_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        title="Test Title",
        description="Test Description",
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    assert element.objectId == "test_id"
    assert element.size.width == 100
    assert element.size.height == 100
    assert element.transform.translateX == 0
    assert element.transform.translateY == 0
    assert element.transform.scaleX == 1
    assert element.transform.scaleY == 1
    assert element.title == "Test Title"
    assert element.description == "Test Description"
    assert element.shape is not None


def test_shape_element():
    """Test ShapeElement functionality."""
    element = ShapeElement(
        objectId="shape_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    assert element.shape is not None
    assert element.shape.shapeType == ShapeType.RECTANGLE

    # Create request should generate a valid request
    request = element.create_request("page_id")
    assert len(request) == 1
    # Now returns a domain object instead of a dictionary
    assert hasattr(request[0], "shapeType")
    assert request[0].shapeType == ShapeType.RECTANGLE


def test_line_element():
    """Test LineElement functionality."""
    element = LineElement(
        objectId="line_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        line=Line(lineType="STRAIGHT", lineProperties=LineProperties()),
    )

    assert element.line is not None
    assert element.line.lineType == "STRAIGHT"

    # Create request should generate a valid request
    request_objects = element.create_request("page_id")
    assert len(request_objects) == 1

    # Convert request objects to dictionaries to test the final format
    requests = []
    for req_obj in request_objects:
        requests.extend(req_obj.to_request())

    assert len(requests) == 1
    assert "createLine" in requests[0]
    assert requests[0]["createLine"]["lineCategory"] == "STRAIGHT"


# TODO: the correspoding API request seems to have been hallucinated.
# Should clean this up when this is next needed
# def test_word_art_element():
#     """Test WordArtElement functionality."""
#     element = WordArtElement(
#         objectId="wordart_id",
#         size=Size(width=100, height=100),
#         transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
#         wordArt=WordArt(renderedText="Test Word Art"),
#     )
#
#     assert element.wordArt is not None
#     assert element.wordArt.renderedText == "Test Word Art"
#
#     # Create request should generate a valid request
#     request_objects = element.create_request("page_id")
#     assert len(request_objects) == 1
#
#     # Convert request objects to dictionaries to test the final format
#     requests = []
#     for req_obj in request_objects:
#         requests.extend(req_obj.to_request())
#
#     assert len(requests) == 1
#     assert "createWordArt" in requests[0]
#     assert requests[0]["createWordArt"]["renderedText"] == "Test Word Art"


def test_sheets_chart_element():
    """Test SheetsChartElement functionality."""
    element = SheetsChartElement(
        objectId="chart_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        sheetsChart=SheetsChart(
            spreadsheetId="spreadsheet_id",
            chartId=123,
            sheetsChartProperties=SheetsChartProperties(),
        ),
    )

    assert element.sheetsChart is not None
    assert element.sheetsChart.spreadsheetId == "spreadsheet_id"
    assert element.sheetsChart.chartId == 123

    # Create request should generate a valid request
    request_objects = element.create_request("page_id")
    assert len(request_objects) == 1

    # Convert request objects to dictionaries to test the final format
    requests = []
    for req_obj in request_objects:
        requests.extend(req_obj.to_request())

    assert len(requests) == 1
    assert "createSheetsChart" in requests[0]
    assert requests[0]["createSheetsChart"]["spreadsheetId"] == "spreadsheet_id"
    assert requests[0]["createSheetsChart"]["chartId"] == 123


def test_update_request_with_title_description():
    """Test that update request includes title and description."""
    element = ShapeElement(
        objectId="element_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        title="Updated Title",
        description="Updated Description",
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    request_objects = element.element_to_update_request("element_id")
    assert len(request_objects) >= 1

    # Convert request objects to dictionaries to test the final format
    requests = []
    for req_obj in request_objects:
        requests.extend(req_obj.to_request())

    # Find the update request for page element alt text (title and description)
    update_request = None
    for req in requests:
        if "updatePageElementAltText" in req:
            update_request = req
            break

    assert update_request is not None
    assert update_request["updatePageElementAltText"]["objectId"] == "element_id"
    assert update_request["updatePageElementAltText"]["title"] == "Updated Title"
    assert update_request["updatePageElementAltText"]["description"] == "Updated Description"


def test_discriminated_union_with_type_adapter():
    """Test that the discriminated union works with TypeAdapter."""
    page_element_adapter = TypeAdapter(PageElement)

    # Test with shape data
    shape_data = {
        "objectId": "shape_id",
        "size": {"width": 100, "height": 100},
        "transform": {"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1},
        "shape": {"shapeType": "RECTANGLE", "shapeProperties": {}},
    }

    element = page_element_adapter.validate_python(shape_data)
    assert isinstance(element, ShapeElement)
    assert element.shape.shapeType == ShapeType.RECTANGLE


def test_absolute_size_with_dimension_objects():
    """Test absolute_size method with Dimension objects for size."""
    element = ShapeElement(
        objectId="test_id",
        size=Size(
            width=Dimension(magnitude=3000000, unit="EMU"),
            height=Dimension(magnitude=3000000, unit="EMU"),
        ),
        transform=Transform(translateX=0, translateY=0, scaleX=0.3, scaleY=0.12, unit="EMU"),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    # Test conversion to centimeters
    width_cm, height_cm = element.absolute_size("cm")

    # Expected calculation:
    # actual_width_emu = 3000000 * 0.3 = 900000
    # actual_height_emu = 3000000 * 0.12 = 360000
    # width_cm = 900000 / 360000 = 2.5
    # height_cm = 360000 / 360000 = 1.0
    assert abs(width_cm - 2.5) < 0.001
    assert abs(height_cm - 1.0) < 0.001

    # Test conversion to inches
    width_in, height_in = element.absolute_size("in")

    # Expected calculation:
    # width_in = 900000 / 914400 ≈ 0.984
    # height_in = 360000 / 914400 ≈ 0.394
    assert abs(width_in - 0.984375) < 0.001
    assert abs(height_in - 0.39375) < 0.001


def test_absolute_size_with_float_values():
    """Test absolute_size method with float values for size."""
    element = ShapeElement(
        objectId="test_id",
        size=Size(width=1000000, height=2000000),  # Direct float values in EMUs
        transform=Transform(translateX=0, translateY=0, scaleX=2.0, scaleY=0.5),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    # Test conversion to centimeters
    width_cm, height_cm = element.absolute_size("cm")

    # Expected calculation:
    # actual_width_emu = 1000000 * 2.0 = 2000000
    # actual_height_emu = 2000000 * 0.5 = 1000000
    # width_cm = 2000000 / 360000 ≈ 5.556
    # height_cm = 1000000 / 360000 ≈ 2.778
    assert abs(width_cm - 5.555555555555555) < 0.001
    assert abs(height_cm - 2.777777777777778) < 0.001


def test_absolute_size_invalid_units():
    """Test absolute_size method with invalid units."""
    element = ShapeElement(
        objectId="test_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    with pytest.raises(ValueError, match="Units must be 'cm' or 'in'"):
        element.absolute_size("px")


def test_absolute_size_no_size():
    """Test absolute_size method when size is None."""
    element = ShapeElement(
        objectId="test_id",
        size=None,
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties()),
    )

    with pytest.raises(ValueError, match="Element size is not available"):
        element.absolute_size("cm")


def test_recursive_group_structure():
    """Test that Group can contain PageElements and GroupElement works correctly."""

    # Create a simple image element
    image_data = {
        "objectId": "image1",
        "size": {"width": 100, "height": 100},
        "transform": {"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1},
        "image": {"contentUrl": "https://example.com/image.jpg"},
    }

    # Create a group element containing the image
    group_data = {
        "objectId": "group1",
        "size": {"width": 200, "height": 200},
        "transform": {"translateX": 10, "translateY": 10, "scaleX": 1, "scaleY": 1},
        "elementGroup": {"children": [image_data]},
    }

    # Test creating PageElement from group data
    page_element_adapter = TypeAdapter(PageElement)

    # This should create a GroupElement
    group_element = page_element_adapter.validate_python(group_data)
    assert isinstance(group_element, GroupElement)
    assert len(group_element.elementGroup.children) == 1

    # Test that the child is properly typed as an ImageElement
    child = group_element.elementGroup.children[0]
    assert isinstance(child, ImageElement)
    assert child.image.contentUrl == "https://example.com/image.jpg"


def test_nested_group_structure():
    """Test that Groups can contain other Groups (nested structure)."""

    # Create a simple image element
    image_data = {
        "objectId": "image1",
        "size": {"width": 100, "height": 100},
        "transform": {"translateX": 0, "translateY": 0, "scaleX": 1, "scaleY": 1},
        "image": {"contentUrl": "https://example.com/image.jpg"},
    }

    # Create an inner group containing the image
    inner_group_data = {
        "objectId": "inner_group",
        "size": {"width": 150, "height": 150},
        "transform": {"translateX": 5, "translateY": 5, "scaleX": 1, "scaleY": 1},
        "elementGroup": {"children": [image_data]},
    }

    # Create an outer group containing the inner group and another image
    outer_group_data = {
        "objectId": "outer_group",
        "size": {"width": 300, "height": 300},
        "transform": {"translateX": 20, "translateY": 20, "scaleX": 1, "scaleY": 1},
        "elementGroup": {"children": [inner_group_data, image_data]},
    }

    # Test creating PageElement from nested group data
    page_element_adapter = TypeAdapter(PageElement)
    outer_group = page_element_adapter.validate_python(outer_group_data)

    assert isinstance(outer_group, GroupElement)
    assert len(outer_group.elementGroup.children) == 2

    # First child should be a GroupElement
    inner_group = outer_group.elementGroup.children[0]
    assert isinstance(inner_group, GroupElement)
    assert len(inner_group.elementGroup.children) == 1
    assert isinstance(inner_group.elementGroup.children[0], ImageElement)

    # Second child should be an ImageElement
    second_child = outer_group.elementGroup.children[1]
    assert isinstance(second_child, ImageElement)
