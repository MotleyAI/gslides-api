"""Tests for absolute_size() and absolute_position() returning None when size/transform is missing."""

import pytest

from gslides_api.agnostic.units import OutputUnit
from gslides_api.domain.domain import PageElementProperties, Size, Transform


class TestAbsoluteSizeNone:
    """Test that absolute_size returns None when size or transform is missing."""

    def test_returns_none_when_size_is_none(self):
        props = PageElementProperties(
            size=None,
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        )
        result = props.absolute_size(units=OutputUnit.CM)
        assert result is None

    def test_returns_none_when_transform_is_none(self):
        props = PageElementProperties(
            size=Size(width=914400, height=914400),
            transform=None,
        )
        result = props.absolute_size(units=OutputUnit.CM)
        assert result is None

    def test_returns_none_when_both_missing(self):
        props = PageElementProperties(
            size=None,
            transform=None,
        )
        result = props.absolute_size(units=OutputUnit.CM)
        assert result is None

    def test_returns_tuple_when_both_present(self):
        props = PageElementProperties(
            size=Size(width=914400, height=914400),
            transform=Transform(translateX=0, translateY=0, scaleX=1, scaleY=1),
        )
        result = props.absolute_size(units=OutputUnit.IN)
        assert result is not None
        assert isinstance(result, tuple)
        assert len(result) == 2
        assert result[0] == pytest.approx(1.0, abs=0.01)
        assert result[1] == pytest.approx(1.0, abs=0.01)


class TestAbsolutePositionNone:
    """Test that absolute_position returns None when transform is missing."""

    def test_returns_none_when_transform_is_none(self):
        props = PageElementProperties(
            size=Size(width=914400, height=914400),
            transform=None,
        )
        result = props.absolute_position(units=OutputUnit.CM)
        assert result is None

    def test_returns_tuple_when_transform_present(self):
        props = PageElementProperties(
            size=Size(width=914400, height=914400),
            transform=Transform(translateX=914400, translateY=457200, scaleX=1, scaleY=1),
        )
        result = props.absolute_position(units=OutputUnit.IN)
        assert result is not None
        assert isinstance(result, tuple)
        assert len(result) == 2
        assert result[0] == pytest.approx(1.0, abs=0.01)
        assert result[1] == pytest.approx(0.5, abs=0.01)
