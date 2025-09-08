import pytest

from gslides_api.domain.domain import LayoutReference, PredefinedLayout


def test_layout_reference_validation():
    """Test that LayoutReference validates that exactly one field is set."""
    # Valid: only layoutId is set
    layout_ref = LayoutReference(layoutId="layout-123")
    assert layout_ref.layoutId == "layout-123"
    assert layout_ref.predefinedLayout is None

    # Valid: only predefinedLayout is set
    layout_ref = LayoutReference(predefinedLayout=PredefinedLayout.TITLE_AND_BODY)
    assert layout_ref.layoutId is None
    assert layout_ref.predefinedLayout == PredefinedLayout.TITLE_AND_BODY

    # Invalid: both fields are None
    with pytest.raises(ValueError, match="Exactly one of layoutId or predefinedLayout must be set"):
        LayoutReference()

    # Invalid: both fields are set
    with pytest.raises(ValueError, match="Exactly one of layoutId or predefinedLayout must be set"):
        LayoutReference(layoutId="layout-123", predefinedLayout=PredefinedLayout.TITLE_AND_BODY)
