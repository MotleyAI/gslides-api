import logging
from unittest.mock import Mock, patch

import pytest

from gslides_api.domain.domain import (Dimension, PageElementProperties, Size,
                                       Transform, Unit)
from gslides_api.domain.text import PlaceholderType, ShapeProperties
from gslides_api.element.base import ElementKind
from gslides_api.element.image import ImageElement
from gslides_api.element.shape import Placeholder, Shape, ShapeElement
from gslides_api.element.text_content import TextContent
from gslides_api.page.slide import Slide
from gslides_api.page.slide_properties import SlideProperties
from gslides_api.presentation import Presentation


class TestPlaceholderParentFetching:
    """Test the placeholder parent object resolution functionality."""

    def test_placeholder_parent_resolution_success(self):
        """Test successful resolution of a placeholder parent object."""
        # Create a parent shape element
        parent_shape = ShapeElement(
            objectId="parent-shape-1",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        # Create a child shape element with placeholder referencing the parent
        child_placeholder = Placeholder(
            type=PlaceholderType.BODY, parentObjectId="parent-shape-1"
        )
        child_shape = ShapeElement(
            objectId="child-shape-1",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=child_placeholder,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=50, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=150.0, translateY=150.0, unit="EMU"
            ),
        )

        # Create a slide with both elements
        slide = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[parent_shape, child_shape],
        )

        # Create presentation
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide],
        )

        # Verify that the parent object was resolved
        assert child_shape.shape.placeholder.parent_object is not None
        assert child_shape.shape.placeholder.parent_object == parent_shape

    def test_placeholder_parent_resolution_missing_parent(self, caplog):
        """Test logging when parent object is not found."""
        # Create a child shape element with placeholder referencing non-existent parent
        child_placeholder = Placeholder(
            type=PlaceholderType.BODY, parentObjectId="non-existent-parent"
        )
        child_shape = ShapeElement(
            objectId="child-shape-1",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=child_placeholder,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=50, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=150.0, translateY=150.0, unit="EMU"
            ),
        )

        # Create a slide with the child element
        slide = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[child_shape],
        )

        # Create presentation
        with caplog.at_level(logging.WARNING):
            presentation = Presentation(
                presentationId="test-presentation-id",
                pageSize=Size(
                    width=Dimension(magnitude=1280, unit=Unit.PT),
                    height=Dimension(magnitude=720, unit=Unit.PT),
                ),
                slides=[slide],
            )

        # Verify that parent object is None and warning was logged
        assert child_shape.shape.placeholder.parent_object is None
        assert (
            "Parent object non-existent-parent not found for placeholder in element child-shape-1"
            in caplog.text
        )

    def test_placeholder_parent_resolution_multiple_matches_uses_first(self):
        """Test that when multiple elements have same objectId, first match is used."""
        # Create two parent shape elements with same objectId
        parent_shape_1 = ShapeElement(
            objectId="duplicate-parent",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        parent_shape_2 = ShapeElement(
            objectId="duplicate-parent",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=200, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=200.0, translateY=200.0, unit="EMU"
            ),
        )

        # Create a child shape element with placeholder referencing the duplicate parent
        child_placeholder = Placeholder(
            type=PlaceholderType.BODY, parentObjectId="duplicate-parent"
        )
        child_shape = ShapeElement(
            objectId="child-shape-1",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=child_placeholder,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=50, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=150.0, translateY=150.0, unit="EMU"
            ),
        )

        # Create slides with the elements (first parent on first slide, second parent on second slide)
        slide1 = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[parent_shape_1, child_shape],
        )
        slide2 = Slide(
            objectId="slide-2",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout2", masterObjectId="test-master2"
            ),
            pageElements=[parent_shape_2],
        )

        # Create presentation
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide1, slide2],
        )

        # Verify that the first parent object was used
        assert child_shape.shape.placeholder.parent_object == parent_shape_1

    def test_only_processes_shape_elements_in_slides(self):
        """Test that only ShapeElements in slides are processed, not in layouts/masters."""
        # This test verifies the method only processes ShapeElements and ignores other types

        # Create a shape element without placeholder (should be ignored)
        shape_no_placeholder = ShapeElement(
            objectId="shape-no-placeholder",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        # Create a slide with the element
        slide = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[shape_no_placeholder],
        )

        # Create presentation - this should not crash and should ignore the shape without placeholder
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide],
        )

        # Verify that the shape was not processed (no parent_object set since it has no placeholder)
        # If we get here without errors, the test passes

    def test_handles_empty_slides_list(self):
        """Test that the method handles presentations with no slides gracefully."""
        # Create presentation with no slides
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[],
        )

        # Should not crash

    def test_handles_none_slides_list(self):
        """Test that the method handles presentations with None slides gracefully."""
        # Create presentation with None slides
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=None,
        )

        # Should not crash

    def test_shape_without_placeholder_ignored(self):
        """Test that shapes without placeholders are ignored."""
        # Create a shape element without a placeholder
        shape_without_placeholder = ShapeElement(
            objectId="shape-no-placeholder",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        # Create a slide with the element
        slide = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[shape_without_placeholder],
        )

        # Create presentation - should not crash
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide],
        )

    def test_shape_with_placeholder_but_no_parent_id_ignored(self):
        """Test that shapes with placeholders but no parentObjectId are ignored."""
        # Create a placeholder without parentObjectId
        placeholder_no_parent = Placeholder(
            type=PlaceholderType.BODY
            # no parentObjectId
        )

        # Create a shape element with placeholder but no parentObjectId
        shape_with_placeholder_no_parent = ShapeElement(
            objectId="shape-placeholder-no-parent",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=placeholder_no_parent,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        # Create a slide with the element
        slide = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[shape_with_placeholder_no_parent],
        )

        # Create presentation - should not crash
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide],
        )

    def test_multiple_placeholders_resolved_correctly(self):
        """Test that multiple placeholders in the same presentation are resolved correctly."""
        # Create parent shapes
        parent_shape_1 = ShapeElement(
            objectId="parent-1",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=300, unit=Unit.PT),
                height=Dimension(magnitude=100, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=100.0, translateY=100.0, unit="EMU"
            ),
        )

        parent_shape_2 = ShapeElement(
            objectId="parent-2",
            shape=Shape(
                shapeProperties=ShapeProperties(), text=TextContent(textElements=[])
            ),
            size=Size(
                width=Dimension(magnitude=400, unit=Unit.PT),
                height=Dimension(magnitude=150, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=200.0, translateY=200.0, unit="EMU"
            ),
        )

        # Create child shapes with placeholders referencing different parents
        child_placeholder_1 = Placeholder(
            type=PlaceholderType.BODY, parentObjectId="parent-1"
        )
        child_shape_1 = ShapeElement(
            objectId="child-1",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=child_placeholder_1,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=200, unit=Unit.PT),
                height=Dimension(magnitude=50, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=150.0, translateY=150.0, unit="EMU"
            ),
        )

        child_placeholder_2 = Placeholder(
            type=PlaceholderType.TITLE, parentObjectId="parent-2"
        )
        child_shape_2 = ShapeElement(
            objectId="child-2",
            shape=Shape(
                shapeProperties=ShapeProperties(),
                placeholder=child_placeholder_2,
                text=TextContent(textElements=[]),
            ),
            size=Size(
                width=Dimension(magnitude=250, unit=Unit.PT),
                height=Dimension(magnitude=75, unit=Unit.PT),
            ),
            transform=Transform(
                scaleX=1.0, scaleY=1.0, translateX=250.0, translateY=250.0, unit="EMU"
            ),
        )

        # Create slides with the elements
        slide1 = Slide(
            objectId="slide-1",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout", masterObjectId="test-master"
            ),
            pageElements=[parent_shape_1, child_shape_1],
        )
        slide2 = Slide(
            objectId="slide-2",
            slideProperties=SlideProperties(
                layoutObjectId="test-layout2", masterObjectId="test-master2"
            ),
            pageElements=[parent_shape_2, child_shape_2],
        )

        # Create presentation
        presentation = Presentation(
            presentationId="test-presentation-id",
            pageSize=Size(
                width=Dimension(magnitude=1280, unit=Unit.PT),
                height=Dimension(magnitude=720, unit=Unit.PT),
            ),
            slides=[slide1, slide2],
        )

        # Verify both placeholders were resolved correctly
        assert child_shape_1.shape.placeholder.parent_object == parent_shape_1
        assert child_shape_2.shape.placeholder.parent_object == parent_shape_2
