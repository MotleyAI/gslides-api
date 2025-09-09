import pytest

from gslides_api.element.base import ElementKind


def test_element_kind_enum():
    """Test that the ElementKind enum has the expected values."""
    # Check that all expected values are present
    assert ElementKind.GROUP.value == "elementGroup"
    assert ElementKind.SHAPE.value == "shape"
    assert ElementKind.IMAGE.value == "image"
    assert ElementKind.VIDEO.value == "video"
    assert ElementKind.LINE.value == "line"
    assert ElementKind.TABLE.value == "table"
    assert ElementKind.WORD_ART.value == "wordArt"
    assert ElementKind.SHEETS_CHART.value == "sheetsChart"
    assert ElementKind.SPEAKER_SPOTLIGHT.value == "speakerSpotlight"

    # Check that we can convert from string to enum
    assert ElementKind("shape") == ElementKind.SHAPE
    assert ElementKind("table") == ElementKind.TABLE
