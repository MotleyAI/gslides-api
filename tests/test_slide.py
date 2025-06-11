import pytest
from gslides_api.page import Page
from gslides_api import SlidePageProperties


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
    def mock_create_blank(self, presentation_id, insertion_index=None, slide_layout_reference=None, layoout_placeholder_id_mapping=None):
        # Return a Slide object instead of just a string
        mock_slide = Page(
            objectId="new-slide-id",
            slideProperties=SlidePageProperties(
                layoutObjectId="test-layout-id", masterObjectId="test-master-id"
            ),
            pageProperties=SlidePageProperties(pageBackgroundFill={}),
        )
        return mock_slide

    # Mock the slides_batch_update function
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
    monkeypatch.setattr(gslides_api.page, "slides_batch_update", mock_slides_batch_update)

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
