import pytest
from pydantic import TypeAdapter
from gslides_api.element.element import (
    PageElement,
    LineElement,
    WordArtElement,
    SheetsChartElement,
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


def test_word_art_element():
    """Test WordArtElement functionality."""
    element = WordArtElement(
        objectId="wordart_id",
        size=Size(width=100, height=100),
        transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        wordArt=WordArt(renderedText="Test Word Art"),
    )

    assert element.wordArt is not None
    assert element.wordArt.renderedText == "Test Word Art"

    # Create request should generate a valid request
    request_objects = element.create_request("page_id")
    assert len(request_objects) == 1

    # Convert request objects to dictionaries to test the final format
    requests = []
    for req_obj in request_objects:
        requests.extend(req_obj.to_request())

    assert len(requests) == 1
    assert "createWordArt" in requests[0]
    assert requests[0]["createWordArt"]["renderedText"] == "Test Word Art"


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
