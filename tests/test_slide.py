import pytest
from gslides_api.page import Page
from gslides_api import SlidePageProperties
from gslides_api.element import ShapeElement
from gslides_api.domain import Size, Transform, Shape, ShapeType, ShapeProperties


def test_presentation_id_not_in_api_format():
    """Test that presentation_id is not included in the API format."""
    # Create a minimal slide
    slide = Page(
        objectId="test-slide-id",
        slideProperties=SlidePageProperties(
            layoutObjectId="test-layout-id", masterObjectId="test-master-id"
        ),
        pageProperties=SlidePageProperties(pageBackgroundFill={}),
        presentation_id="test-presentation-id",
    )

    # Convert to API format
    api_format = slide.to_api_format()

    # Check that presentation_id is not in the API format
    assert "presentation_id" not in api_format

    # Check that other fields are in the API format
    assert api_format["objectId"] == "test-slide-id"
    assert "slideProperties" in api_format
    assert "pageProperties" in api_format


def test_write_sets_presentation_id(monkeypatch):
    """Test that the write_copy method sets the presentation_id."""
    # Create a minimal slide
    slide = Page(
        objectId="test-slide-id",
        slideProperties=SlidePageProperties(
            layoutObjectId="test-layout-id", masterObjectId="test-master-id"
        ),
        pageProperties=SlidePageProperties(pageBackgroundFill={}),
    )

    # Mock the create_blank method to avoid API calls
    def mock_create_blank(
        self,
        presentation_id,
        insertion_index=None,
        slide_layout_reference=None,
        layoout_placeholder_id_mapping=None,
    ):
        # Return a Slide object instead of just a string
        mock_slide = Page(
            objectId="new-slide-id",
            slideProperties=SlidePageProperties(
                layoutObjectId="test-layout-id", masterObjectId="test-master-id"
            ),
            pageProperties=SlidePageProperties(pageBackgroundFill={}),
        )
        return mock_slide

    # Mock the batch_update function
    def mock_slides_batch_update(requests, presentation_id):
        return {"replies": [{"createSlide": {"objectId": "new-slide-id"}}]}

    # Mock the from_ids method to avoid API calls
    def mock_from_ids(cls, presentation_id, slide_id):
        return Page(
            objectId=slide_id,
            slideProperties=SlidePageProperties(
                layoutObjectId="test-layout-id", masterObjectId="test-master-id"
            ),
            pageProperties=SlidePageProperties(pageBackgroundFill={}),
            presentation_id=presentation_id,
        )

    # Apply the monkeypatches
    import gslides_api.page

    monkeypatch.setattr(Page, "create_blank", mock_create_blank)
    monkeypatch.setattr(Page, "from_ids", classmethod(mock_from_ids))
    monkeypatch.setattr(gslides_api.page, "batch_update", mock_slides_batch_update)

    # Call write_copy with a presentation_id
    result = slide.write_copy(presentation_id="test-presentation-id")

    # Check that presentation_id is set on the returned slide
    assert result.presentation_id == "test-presentation-id"


def test_duplicate_preserves_presentation_id(monkeypatch):
    """Test that the duplicate method preserves the presentation_id."""
    # Create a minimal slide
    slide = Page(
        objectId="test-slide-id",
        slideProperties=SlidePageProperties(
            layoutObjectId="test-layout-id", masterObjectId="test-master-id"
        ),
        pageProperties=SlidePageProperties(pageBackgroundFill={}),
        presentation_id="test-presentation-id",
    )

    # Mock the duplicate_object function to avoid API calls
    def mock_duplicate_object(object_id, presentation_id, id_map=None):
        return "new-slide-id"

    # Mock the from_ids method to avoid API calls
    def mock_from_ids(cls, presentation_id, slide_id):
        return Page(
            objectId=slide_id,
            slideProperties=SlidePageProperties(
                layoutObjectId="test-layout-id", masterObjectId="test-master-id"
            ),
            pageProperties=SlidePageProperties(pageBackgroundFill={}),
            presentation_id=presentation_id,
        )

    # Apply the monkeypatches
    import gslides_api.page

    monkeypatch.setattr(gslides_api.page, "duplicate_object", mock_duplicate_object)
    monkeypatch.setattr(Page, "from_ids", classmethod(mock_from_ids))

    # Call duplicate
    result = slide.duplicate()

    # Check that presentation_id is preserved on the returned slide
    assert result.presentation_id == "test-presentation-id"


def test_presentation_id_propagation_on_creation():
    """Test that presentation_id is automatically propagated to pageElements upon Page creation."""
    # Create some mock page elements
    element1 = ShapeElement(
        objectId="element1",
        size=Size(width=100, height=50),
        transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0, translateY=0),
        shape=Shape(shapeType=ShapeType.TEXT_BOX, shapeProperties=ShapeProperties())
    )

    element2 = ShapeElement(
        objectId="element2",
        size=Size(width=200, height=100),
        transform=Transform(scaleX=1.0, scaleY=1.0, translateX=100, translateY=100),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties())
    )

    # Create Page with presentation_id and pageElements
    page = Page(
        objectId="slide1",
        presentation_id="test-presentation-123",
        pageElements=[element1, element2]
    )

    # Verify that presentation_id was propagated to all elements
    assert page.presentation_id == "test-presentation-123"
    assert element1.presentation_id == "test-presentation-123"
    assert element2.presentation_id == "test-presentation-123"


def test_presentation_id_propagation_on_modification():
    """Test that presentation_id is propagated to pageElements when modified on existing Page."""
    # Create some mock page elements
    element1 = ShapeElement(
        objectId="element1",
        size=Size(width=100, height=50),
        transform=Transform(scaleX=1.0, scaleY=1.0, translateX=0, translateY=0),
        shape=Shape(shapeType=ShapeType.TEXT_BOX, shapeProperties=ShapeProperties())
    )

    element2 = ShapeElement(
        objectId="element2",
        size=Size(width=200, height=100),
        transform=Transform(scaleX=1.0, scaleY=1.0, translateX=100, translateY=100),
        shape=Shape(shapeType=ShapeType.RECTANGLE, shapeProperties=ShapeProperties())
    )

    # Create Page with initial presentation_id
    page = Page(
        objectId="slide1",
        presentation_id="initial-presentation-id",
        pageElements=[element1, element2]
    )

    # Verify initial state
    assert page.presentation_id == "initial-presentation-id"
    assert element1.presentation_id == "initial-presentation-id"
    assert element2.presentation_id == "initial-presentation-id"

    # Modify presentation_id
    page.presentation_id = "new-presentation-456"

    # Verify that presentation_id was propagated to all elements
    assert page.presentation_id == "new-presentation-456"
    assert element1.presentation_id == "new-presentation-456"
    assert element2.presentation_id == "new-presentation-456"


def test_presentation_id_propagation_with_no_elements():
    """Test that Page with no pageElements works correctly and doesn't crash."""
    # Create Page with no pageElements
    page = Page(
        objectId="slide2",
        presentation_id="test-presentation-789"
    )

    # Verify that presentation_id is set correctly
    assert page.presentation_id == "test-presentation-789"

    # Modify presentation_id (should not crash)
    page.presentation_id = "modified-presentation-id"
    assert page.presentation_id == "modified-presentation-id"


def test_presentation_id_propagation_with_empty_elements_list():
    """Test that Page with empty pageElements list works correctly."""
    # Create Page with empty pageElements list
    page = Page(
        objectId="slide3",
        presentation_id="test-presentation-empty",
        pageElements=[]
    )

    # Verify that presentation_id is set correctly
    assert page.presentation_id == "test-presentation-empty"

    # Modify presentation_id (should not crash)
    page.presentation_id = "modified-empty-presentation"
    assert page.presentation_id == "modified-empty-presentation"


def test_presentation_id_propagation_when_adding_elements_later():
    """Test that pageElements added after Page creation get correct presentation_id."""
    # Create Page without pageElements initially
    page = Page(
        objectId="slide4",
        presentation_id="test-presentation-later"
    )

    # Create and add pageElements later
    element = ShapeElement(
        objectId="element3",
        size=Size(width=150, height=75),
        transform=Transform(scaleX=1.0, scaleY=1.0, translateX=200, translateY=200),
        shape=Shape(shapeType=ShapeType.ELLIPSE, shapeProperties=ShapeProperties())
    )

    page.pageElements = [element]

    # Trigger propagation by setting presentation_id again
    page.presentation_id = page.presentation_id

    # Verify that the newly added element gets the correct presentation_id
    assert element.presentation_id == "test-presentation-later"
