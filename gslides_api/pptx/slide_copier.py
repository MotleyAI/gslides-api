"""
Slide copier manager for PowerPoint presentations.

Main orchestration class that coordinates all aspects of slide copying
including ID management, XML manipulation, relationship copying, and shape copying.
"""

import logging
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.parts.slide import SlidePart
from pptx.slide import Slide

from .id_manager import IdManager
from .relationship_copier import RelationshipCopier
from .shape_copier import ShapeCopier
from .xml_utils import XmlUtils

logger = logging.getLogger(__name__)


def _remove_layout_placeholders(slide: Slide) -> None:
    """Remove placeholder shapes that come from slide layouts.

    When slides.add_slide(layout) is called, python-pptx automatically
    includes all placeholder shapes from that layout. These placeholders
    show text like "Click to add title" or "Click to add text" which we
    don't want in the exported presentation.
    """
    shapes_to_remove = []
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            shapes_to_remove.append(shape)

    for shape in shapes_to_remove:
        sp = shape.element
        sp.getparent().remove(sp)


class SlideCopierManager:
    """
    Main orchestration class for robust slide copying operations.

    Coordinates all aspects of slide copying to ensure no XML corruption,
    duplicate IDs, or broken relationships.
    """

    def __init__(self, target_presentation: Presentation):
        """
        Initialize the slide copier manager.

        Args:
            target_presentation: The presentation to copy slides into
        """
        self.target_presentation = target_presentation
        self.id_manager = IdManager(target_presentation)
        self.relationship_copier = RelationshipCopier()
        self.shape_copier = ShapeCopier(self.id_manager)

        self.copy_operations: List[Dict[str, Any]] = []
        self.errors: List[str] = []

    def copy_slide(
        self,
        source_slide: Slide,
        insertion_index: Optional[int] = None,
        copy_relationships: bool = True,
        regenerate_ids: bool = True,
        position_offset: Optional[Tuple[float, float]] = None,
        layout_matching: str = "auto"
    ) -> Optional[Slide]:
        """
        Copy a slide with all its content to the target presentation.

        Args:
            source_slide: The slide to copy
            insertion_index: Index where to insert the slide (None = append)
            copy_relationships: Whether to copy relationships (images, etc.)
            regenerate_ids: Whether to regenerate all IDs
            position_offset: Optional position offset for shapes
            layout_matching: How to match layouts ("auto", "blank", "match")

        Returns:
            The newly created slide, or None if copying failed
        """
        operation_id = len(self.copy_operations)
        operation = {
            'id': operation_id,
            'source_slide_id': getattr(source_slide, 'slide_id', 'unknown'),
            'status': 'started',
            'errors': [],
            'warnings': [],
        }
        self.copy_operations.append(operation)

        try:
            logger.info(f"Starting slide copy operation {operation_id}")

            # Step 1: Create target slide with appropriate layout
            target_slide = self._create_target_slide(source_slide, layout_matching)
            if not target_slide:
                operation['status'] = 'failed'
                operation['errors'].append('Failed to create target slide')
                return None

            operation['target_slide_id'] = target_slide.slide_id

            # Step 2: Copy relationships if requested
            relationship_mapping = {}
            if copy_relationships:
                try:
                    relationship_mapping = self.relationship_copier.copy_slide_relationships(
                        source_slide, target_slide
                    )
                    operation['relationships_copied'] = len(relationship_mapping)
                except Exception as e:
                    operation['warnings'].append(f'Relationship copying failed: {e}')
                    logger.warning(f"Relationship copying failed for operation {operation_id}: {e}")

            # Step 3: Copy all shapes
            copied_shapes = []
            shape_errors = []

            for shape_idx, source_shape in enumerate(source_slide.shapes):
                try:
                    copied_shape = self.shape_copier.copy_shape(
                        source_shape=source_shape,
                        target_slide=target_slide,
                        position_offset=position_offset,
                        relationship_mapping=relationship_mapping
                    )
                    if copied_shape:
                        copied_shapes.append({
                            'source_idx': shape_idx,
                            'target_shape': copied_shape,
                            'shape_type': source_shape.shape_type
                        })
                    else:
                        # Shape was copied via XML (GROUP, FREEFORM, etc.) or failed
                        # Not all shape types return a shape object
                        logger.debug(f'Shape {shape_idx} copied via XML or skipped (no return object)')

                except Exception as e:
                    error_msg = f'Error copying shape {shape_idx}: {e}'
                    shape_errors.append(error_msg)
                    logger.warning(error_msg)

            operation['shapes_copied'] = len(copied_shapes)
            operation['shape_errors'] = shape_errors

            # Step 4: Copy speaker notes
            try:
                notes_copied = self.relationship_copier.copy_notes_slide_relationships(
                    source_slide, target_slide
                )
                operation['notes_copied'] = notes_copied
            except Exception as e:
                operation['warnings'].append(f'Notes copying failed: {e}')

            # Step 5: Move slide to correct position if needed
            if insertion_index is not None:
                try:
                    self._move_slide_to_position(target_slide, insertion_index)
                    operation['moved_to_index'] = insertion_index
                except Exception as e:
                    operation['warnings'].append(f'Slide positioning failed: {e}')

            # Step 6: Validate the result
            validation_result = self._validate_copied_slide(target_slide)
            operation['validation'] = validation_result

            if validation_result['valid']:
                operation['status'] = 'completed'
                logger.info(f"Successfully completed slide copy operation {operation_id}")
                return target_slide
            else:
                operation['status'] = 'completed_with_warnings'
                operation['warnings'].extend(validation_result['issues'])
                logger.warning(f"Slide copy operation {operation_id} completed with warnings")
                return target_slide

        except Exception as e:
            operation['status'] = 'failed'
            operation['errors'].append(str(e))
            logger.error(f"Slide copy operation {operation_id} failed: {e}")
            return None

    def copy_slide_safe(
        self,
        source_slide: Slide,
        insertion_index: Optional[int] = None,
        fallback_to_layout_only: bool = True
    ) -> Optional[Slide]:
        """
        Copy a slide with automatic fallback strategies.

        This method tries different copying strategies if the full copy fails,
        ensuring that at least some version of the slide is created.

        Args:
            source_slide: The slide to copy
            insertion_index: Index where to insert the slide
            fallback_to_layout_only: Whether to fall back to layout-only copy

        Returns:
            The newly created slide, or None if all strategies failed
        """
        # Try full copying first
        result = self.copy_slide(
            source_slide,
            insertion_index=insertion_index,
            copy_relationships=True,
            regenerate_ids=True
        )

        if result:
            return result

        logger.warning("Full slide copy failed, trying without relationships")

        # Try without relationship copying
        result = self.copy_slide(
            source_slide,
            insertion_index=insertion_index,
            copy_relationships=False,
            regenerate_ids=True
        )

        if result:
            return result

        if fallback_to_layout_only:
            logger.warning("Shape copying failed, falling back to layout-only copy")

            # Final fallback: create slide with same layout only
            try:
                layout = source_slide.slide_layout
                target_slide = self.target_presentation.slides.add_slide(layout)
                _remove_layout_placeholders(target_slide)

                # Copy just the text content if possible
                if source_slide.has_notes_slide:
                    try:
                        notes_text = source_slide.notes_slide.notes_text_frame.text
                        if notes_text.strip():
                            target_slide.notes_slide.notes_text_frame.text = notes_text
                    except Exception:
                        pass

                return target_slide

            except Exception as e:
                logger.error(f"Even layout-only copy failed: {e}")

        return None

    def _create_target_slide(self, source_slide: Slide, layout_matching: str = "auto") -> Optional[Slide]:
        """Create a target slide with appropriate layout.

        When copying between presentations, we try to find a matching layout
        in the target presentation to avoid duplicating layouts/masters.
        """
        try:
            if layout_matching == "blank":
                # Use blank layout
                layout = self._get_blank_layout()
            elif layout_matching == "match":
                # Try to find matching layout in target, fall back to source layout
                layout = self._find_matching_layout(source_slide.slide_layout)
                if layout is None:
                    layout = source_slide.slide_layout
            else:  # auto
                # Try to find matching layout in target, fall back to blank
                layout = self._find_matching_layout(source_slide.slide_layout)
                if layout is None:
                    layout = self._get_blank_layout()

            slide = self.target_presentation.slides.add_slide(layout)
            _remove_layout_placeholders(slide)
            return slide

        except Exception as e:
            logger.error(f"Failed to create target slide: {e}")
            return None

    def _find_matching_layout(self, source_layout) -> Optional[Any]:
        """Find a layout in the target presentation that matches the source layout.

        Matching is done by layout name, which is typically descriptive
        (e.g., "Title Slide", "Title and Content", "Blank").
        """
        try:
            source_name = source_layout.name
            logger.info(f"Looking for layout matching: {source_name}")
            for layout in self.target_presentation.slide_layouts:
                if layout.name == source_name:
                    logger.info(f"Found matching layout in target: {source_name}")
                    return layout

            # Try partial matching if exact match fails
            for layout in self.target_presentation.slide_layouts:
                if source_name.lower() in layout.name.lower() or layout.name.lower() in source_name.lower():
                    logger.info(f"Found partial matching layout: {layout.name} for {source_name}")
                    return layout

            logger.warning(f"No matching layout found for: {source_name}")
            return None
        except Exception as e:
            logger.error(f"Error finding matching layout: {e}")
            return None

    def _get_blank_layout(self):
        """Get the blank slide layout from the presentation."""
        try:
            # Find the layout with the fewest placeholders (likely blank)
            layout_items_count = [len(layout.placeholders) for layout in self.target_presentation.slide_layouts]
            min_items = min(layout_items_count)
            blank_layout_id = layout_items_count.index(min_items)
            return self.target_presentation.slide_layouts[blank_layout_id]
        except Exception:
            # Fall back to first layout
            return self.target_presentation.slide_layouts[0]

    def _move_slide_to_position(self, slide: Slide, insertion_index: int):
        """Move a slide to a specific position (simplified implementation)."""
        # Note: python-pptx doesn't have built-in slide reordering
        # This would require complex XML manipulation
        # For now, we'll just log the intent
        logger.debug(f"Slide movement requested to index {insertion_index} (not implemented)")

    def _validate_copied_slide(self, slide: Slide) -> Dict[str, Any]:
        """Validate a copied slide for common issues."""
        validation = {
            'valid': True,
            'issues': [],
            'warnings': [],
            'shape_count': len(slide.shapes) if slide else 0,
        }

        try:
            if not slide:
                validation['valid'] = False
                validation['issues'].append('Slide is None')
                return validation

            # Check slide has valid ID
            slide_id = getattr(slide, 'slide_id', None)
            if not slide_id:
                validation['warnings'].append('Slide missing ID')

            # Check shapes for issues
            for idx, shape in enumerate(slide.shapes):
                try:
                    # Try to access basic properties
                    _ = shape.shape_type
                    _ = shape.left
                    _ = shape.top
                    _ = shape.width
                    _ = shape.height
                except Exception as e:
                    validation['warnings'].append(f'Shape {idx} has property access issues: {e}')

            # Check if slide can be saved (basic validation)
            if hasattr(slide, 'part'):
                try:
                    # This is a basic check that the slide's XML is well-formed
                    _ = slide.part.element
                except Exception as e:
                    validation['valid'] = False
                    validation['issues'].append(f'Slide XML structure invalid: {e}')

        except Exception as e:
            validation['valid'] = False
            validation['issues'].append(f'Validation error: {e}')

        return validation

    def batch_copy_slides(
        self,
        source_slides: List[Slide],
        copy_relationships: bool = True,
        fail_fast: bool = False
    ) -> List[Optional[Slide]]:
        """
        Copy multiple slides in batch.

        Args:
            source_slides: List of slides to copy
            copy_relationships: Whether to copy relationships
            fail_fast: Whether to stop on first error

        Returns:
            List of copied slides (None for failed copies)
        """
        results = []

        for idx, source_slide in enumerate(source_slides):
            try:
                logger.info(f"Batch copying slide {idx + 1}/{len(source_slides)}")

                result = self.copy_slide(
                    source_slide,
                    copy_relationships=copy_relationships
                )

                results.append(result)

                if not result and fail_fast:
                    logger.error(f"Batch copy failed at slide {idx}, stopping due to fail_fast")
                    break

            except Exception as e:
                logger.error(f"Error in batch copy for slide {idx}: {e}")
                results.append(None)

                if fail_fast:
                    break

        return results

    def get_copy_statistics(self) -> Dict[str, Any]:
        """Get detailed statistics about copy operations."""
        stats = {
            'total_operations': len(self.copy_operations),
            'successful_operations': sum(1 for op in self.copy_operations if op['status'] == 'completed'),
            'failed_operations': sum(1 for op in self.copy_operations if op['status'] == 'failed'),
            'operations_with_warnings': sum(1 for op in self.copy_operations if op['status'] == 'completed_with_warnings'),
            'total_shapes_copied': sum(op.get('shapes_copied', 0) for op in self.copy_operations),
            'total_relationships_copied': sum(op.get('relationships_copied', 0) for op in self.copy_operations),
            'id_manager_stats': self.id_manager.get_stats(),
            'relationship_copier_stats': self.relationship_copier.get_relationship_stats(),
            'shape_copier_stats': self.shape_copier.get_copy_stats(),
        }

        return stats

    def cleanup(self):
        """Clean up resources used during copying."""
        try:
            self.relationship_copier.cleanup()
            self.shape_copier.clear_stats()
            self.copy_operations.clear()
            self.errors.clear()
            logger.debug("Slide copier cleanup completed")
        except Exception as e:
            logger.warning(f"Error during cleanup: {e}")

    def __enter__(self):
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit with cleanup."""
        self.cleanup()