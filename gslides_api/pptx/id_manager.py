"""
ID Manager for PowerPoint presentations.

Handles generation and tracking of unique IDs to prevent XML corruption
and PowerPoint repair prompts when copying slides and shapes.
"""

import uuid
import logging
from typing import Set, Dict, Optional
from pptx import Presentation
from pptx.slide import Slide

logger = logging.getLogger(__name__)


class IdManager:
    """
    Manages unique ID generation for PowerPoint elements.

    Tracks used IDs across a presentation to ensure no duplicates are created,
    which would cause PowerPoint to prompt for file repair.
    """

    def __init__(self, presentation: Presentation):
        """
        Initialize ID manager with existing presentation IDs.

        Args:
            presentation: The PowerPoint presentation to track IDs for
        """
        self.presentation = presentation
        self.used_slide_ids: Set[int] = set()
        self.used_shape_ids: Set[int] = set()
        self.used_creation_ids: Set[str] = set()
        self.next_slide_id = 256  # PowerPoint typically starts at 256
        self.next_shape_id = 1

        # Scan existing presentation for used IDs
        self._scan_existing_ids()

    def _scan_existing_ids(self):
        """Scan the presentation for existing IDs to avoid conflicts."""
        try:
            # Scan slide IDs
            for slide in self.presentation.slides:
                slide_id = slide.slide_id
                if slide_id:
                    self.used_slide_ids.add(slide_id)
                    self.next_slide_id = max(self.next_slide_id, slide_id + 1)

                # Scan shape IDs in each slide
                for shape in slide.shapes:
                    shape_id = getattr(shape, 'shape_id', None)
                    if shape_id:
                        self.used_shape_ids.add(shape_id)
                        self.next_shape_id = max(self.next_shape_id, shape_id + 1)

                    # Extract creation IDs from XML if present
                    creation_id = self._extract_creation_id(shape)
                    if creation_id:
                        self.used_creation_ids.add(creation_id)

        except Exception as e:
            logger.warning(f"Error scanning existing IDs: {e}")
            # Continue with default starting values

    def _extract_creation_id(self, shape) -> Optional[str]:
        """Extract a16:creationId from shape XML if present."""
        try:
            if hasattr(shape, '_element') and shape._element is not None:
                # Look for a16:creationId attribute in the XML
                creation_id_elem = shape._element.xpath('.//a16:creationId',
                                                      namespaces={'a16': 'http://schemas.microsoft.com/office/drawing/2013/main-command'})
                if creation_id_elem:
                    return creation_id_elem[0].get('id')
        except Exception:
            # Ignore XML parsing errors
            pass
        return None

    def generate_unique_slide_id(self) -> int:
        """
        Generate a unique slide ID for the presentation.

        Returns:
            A unique integer slide ID
        """
        while self.next_slide_id in self.used_slide_ids:
            self.next_slide_id += 1

        slide_id = self.next_slide_id
        self.used_slide_ids.add(slide_id)
        self.next_slide_id += 1

        logger.debug(f"Generated unique slide ID: {slide_id}")
        return slide_id

    def generate_unique_shape_id(self) -> int:
        """
        Generate a unique shape ID for the presentation.

        Returns:
            A unique integer shape ID
        """
        while self.next_shape_id in self.used_shape_ids:
            self.next_shape_id += 1

        shape_id = self.next_shape_id
        self.used_shape_ids.add(shape_id)
        self.next_shape_id += 1

        logger.debug(f"Generated unique shape ID: {shape_id}")
        return shape_id

    def generate_unique_creation_id(self) -> str:
        """
        Generate a unique creation ID (a16:creationId) for PowerPoint elements.

        Returns:
            A unique GUID string for creation ID
        """
        while True:
            creation_id = str(uuid.uuid4()).upper()
            if creation_id not in self.used_creation_ids:
                self.used_creation_ids.add(creation_id)
                logger.debug(f"Generated unique creation ID: {creation_id}")
                return creation_id

    def reserve_slide_id(self, slide_id: int):
        """
        Reserve a specific slide ID to prevent conflicts.

        Args:
            slide_id: The slide ID to reserve
        """
        self.used_slide_ids.add(slide_id)
        self.next_slide_id = max(self.next_slide_id, slide_id + 1)

    def reserve_shape_id(self, shape_id: int):
        """
        Reserve a specific shape ID to prevent conflicts.

        Args:
            shape_id: The shape ID to reserve
        """
        self.used_shape_ids.add(shape_id)
        self.next_shape_id = max(self.next_shape_id, shape_id + 1)

    def reserve_creation_id(self, creation_id: str):
        """
        Reserve a specific creation ID to prevent conflicts.

        Args:
            creation_id: The creation ID to reserve
        """
        self.used_creation_ids.add(creation_id)

    def get_id_mapping(self, source_slide: Slide) -> Dict[str, int]:
        """
        Generate ID mapping for all shapes in a source slide.

        This creates a mapping from old shape IDs to new unique shape IDs
        that can be used when copying the slide.

        Args:
            source_slide: The slide to generate ID mapping for

        Returns:
            Dictionary mapping old shape IDs to new unique shape IDs
        """
        id_mapping = {}

        for shape in source_slide.shapes:
            old_shape_id = getattr(shape, 'shape_id', None)
            if old_shape_id:
                new_shape_id = self.generate_unique_shape_id()
                id_mapping[str(old_shape_id)] = new_shape_id

        return id_mapping

    def get_stats(self) -> Dict[str, int]:
        """
        Get statistics about ID usage.

        Returns:
            Dictionary with ID usage statistics
        """
        return {
            'used_slide_ids': len(self.used_slide_ids),
            'used_shape_ids': len(self.used_shape_ids),
            'used_creation_ids': len(self.used_creation_ids),
            'next_slide_id': self.next_slide_id,
            'next_shape_id': self.next_shape_id,
        }