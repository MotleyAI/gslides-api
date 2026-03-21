"""
Slide deletion utilities for PowerPoint presentations.

This module provides robust slide deletion functionality that properly handles
XML manipulation and relationship cleanup to prevent presentation corruption.
"""

import logging
from typing import List, Dict, Any, Optional
from pptx import Presentation

from .xml_utils import XmlUtils

logger = logging.getLogger(__name__)


class SlideDeletionResult:
    """Result of a slide deletion operation."""

    def __init__(self, success: bool = True, error_message: str = "", slide_index: int = -1):
        self.success = success
        self.error_message = error_message
        self.slide_index = slide_index
        self.warnings: List[str] = []

    def add_warning(self, warning: str):
        """Add a warning to the result."""
        self.warnings.append(warning)
        logger.warning(f"Slide deletion warning: {warning}")


class SlideDeleter:
    """
    Safely delete slides from PowerPoint presentations.

    This class handles XML manipulation and relationship cleanup based on
    proven workarounds for python-pptx's lack of official slide deletion API.

    Implementation based on solutions from:
    - https://github.com/scanny/python-pptx/issues/67
    - https://github.com/pyhub-apps/pyhub-office-automation/issues/76
    """

    def __init__(self, presentation: Presentation):
        """
        Initialize the slide deleter.

        Args:
            presentation: The python-pptx Presentation object
        """
        self.presentation = presentation
        self.deleted_count = 0
        self.deletion_log: List[Dict[str, Any]] = []

    def validate_deletion(self, slide_index: int) -> SlideDeletionResult:
        """
        Check if a slide can be safely deleted.

        Args:
            slide_index: Index of the slide to validate

        Returns:
            SlideDeletionResult with validation status
        """
        result = SlideDeletionResult(slide_index=slide_index)

        try:
            # Check if index is valid
            slide_count = len(self.presentation.slides)
            if slide_index < 0 or slide_index >= slide_count:
                result.success = False
                result.error_message = f"Invalid slide index {slide_index}. Presentation has {slide_count} slides."
                return result

            # Check if this is the only slide
            if slide_count == 1:
                result.add_warning("Deleting the only slide in the presentation")

            # Check if slide has complex relationships that might cause issues
            slide = self.presentation.slides[slide_index]
            if hasattr(slide, 'notes_slide') and slide.notes_slide:
                logger.debug("Slide has notes that will be deleted")

            # Validate that we can access the slide ID list
            if not hasattr(self.presentation.slides, '_sldIdLst'):
                result.success = False
                result.error_message = "Cannot access internal slide ID list (_sldIdLst)"
                return result

            # Check that the slide has a relationship ID
            xml_slides = self.presentation.slides._sldIdLst
            slides = list(xml_slides)
            if slide_index >= len(slides):
                result.success = False
                result.error_message = f"Slide index {slide_index} not found in XML slide list"
                return result

            slide_element = slides[slide_index]
            if not hasattr(slide_element, 'rId'):
                result.success = False
                result.error_message = "Slide element missing relationship ID (rId)"
                return result

            logger.debug(f"Slide {slide_index} validation passed")

        except Exception as e:
            result.success = False
            result.error_message = f"Validation error: {str(e)}"
            logger.error(f"Slide validation failed: {e}")

        return result

    def delete_slide(self, slide_index: int) -> SlideDeletionResult:
        """
        Delete a single slide by index with proper cleanup.

        This method implements the proven workaround from python-pptx issues:
        1. Access the internal slide ID list (_sldIdLst)
        2. Get the relationship ID (rId)
        3. Drop the relationship first
        4. Remove from the XML slide list

        Args:
            slide_index: Index of the slide to delete (0-based)

        Returns:
            SlideDeletionResult indicating success/failure
        """
        logger.info(f"Attempting to delete slide at index {slide_index}")

        # Validate the deletion first
        validation_result = self.validate_deletion(slide_index)
        if not validation_result.success:
            return validation_result

        result = SlideDeletionResult(slide_index=slide_index)
        result.warnings.extend(validation_result.warnings)

        try:
            # Access the internal slide ID list
            xml_slides = self.presentation.slides._sldIdLst
            slides = list(xml_slides)
            slide_element = slides[slide_index]

            # Get the relationship ID
            rId = slide_element.rId
            logger.debug(f"Deleting slide with relationship ID: {rId}")

            # Step 1: Drop the relationship FIRST (critical for preventing corruption)
            self.presentation.part.drop_rel(rId)
            logger.debug(f"Dropped relationship {rId}")

            # Step 2: Remove from the XML slide list
            xml_slides.remove(slide_element)
            logger.debug(f"Removed slide element from XML list")

            # Log the successful operation
            self.deleted_count += 1
            deletion_record = {
                'index': slide_index,
                'rId': rId,
                'timestamp': logger.name,  # Simple timestamp placeholder
                'success': True
            }
            self.deletion_log.append(deletion_record)

            logger.info(f"Successfully deleted slide at index {slide_index}")

        except Exception as e:
            result.success = False
            result.error_message = f"Deletion failed: {str(e)}"
            logger.error(f"Failed to delete slide {slide_index}: {e}")

            # Log the failed operation
            deletion_record = {
                'index': slide_index,
                'error': str(e),
                'timestamp': logger.name,
                'success': False
            }
            self.deletion_log.append(deletion_record)

        return result

    def delete_slides(self, slide_indices: List[int]) -> List[SlideDeletionResult]:
        """
        Delete multiple slides in reverse order.

        Deleting in reverse order is critical to prevent index shifting issues
        when removing multiple slides from the same presentation.

        Args:
            slide_indices: List of slide indices to delete

        Returns:
            List of SlideDeletionResult objects, one for each attempted deletion
        """
        logger.info(f"Attempting to delete {len(slide_indices)} slides")

        # Sort indices in reverse order to prevent index shifting
        sorted_indices = sorted(slide_indices, reverse=True)
        results = []

        for slide_index in sorted_indices:
            result = self.delete_slide(slide_index)
            results.append(result)

            # If a deletion fails, log it but continue with others
            if not result.success:
                logger.warning(f"Slide deletion failed for index {slide_index}: {result.error_message}")

        successful_deletions = sum(1 for r in results if r.success)
        logger.info(f"Completed batch deletion: {successful_deletions}/{len(slide_indices)} slides deleted successfully")

        return results

    def get_deletion_stats(self) -> Dict[str, Any]:
        """
        Get statistics about deletion operations performed.

        Returns:
            Dictionary with deletion statistics
        """
        successful_deletions = sum(1 for record in self.deletion_log if record.get('success', False))
        failed_deletions = len(self.deletion_log) - successful_deletions

        return {
            'total_attempted': len(self.deletion_log),
            'successful': successful_deletions,
            'failed': failed_deletions,
            'current_slide_count': len(self.presentation.slides),
            'deletion_log': self.deletion_log.copy()
        }

    def cleanup_orphaned_parts(self) -> int:
        """
        Attempt to clean up orphaned parts after slide deletion.

        Note: This is experimental and may not catch all orphaned parts.
        PowerPoint's auto-repair is more comprehensive.

        Returns:
            Number of orphaned parts found (may not be cleanable)
        """
        logger.info("Attempting to identify orphaned parts")

        # This is a placeholder for future implementation
        # Full orphaned part cleanup is complex and beyond the scope
        # of the basic slide deletion functionality

        orphaned_count = 0

        try:
            # Future: implement orphaned part detection
            # This would involve checking for unreferenced relationships
            # and parts that are no longer needed
            pass

        except Exception as e:
            logger.warning(f"Orphaned part cleanup failed: {e}")

        logger.info(f"Orphaned part cleanup completed, found {orphaned_count} parts")
        return orphaned_count