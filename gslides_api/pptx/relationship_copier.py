"""
Relationship copier for PowerPoint presentations.

Handles copying of slide relationships including images, charts, hyperlinks,
and other embedded objects to ensure no broken references.
"""

import io
import logging
import os
import tempfile
from typing import Any, Dict, List, Optional, Tuple

from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.parts.chart import ChartPart
from pptx.parts.slide import SlidePart
from pptx.slide import Slide

logger = logging.getLogger(__name__)


class RelationshipCopier:
    """
    Manages copying of relationships between slides.

    Handles the complex task of copying images, charts, hyperlinks, and other
    embedded objects while maintaining proper references and avoiding corruption.
    """

    def __init__(self):
        """Initialize the relationship copier."""
        self.copied_relationships: Dict[str, str] = {}
        self.temp_files: List[str] = []

    def copy_slide_relationships(
        self,
        source_slide: Slide,
        target_slide: Slide,
        exclude_notes: bool = True
    ) -> Dict[str, str]:
        """
        Copy all relationships from source slide to target slide.

        Args:
            source_slide: The slide to copy relationships from
            target_slide: The slide to copy relationships to
            exclude_notes: Whether to exclude notes slide relationships

        Returns:
            Dictionary mapping old relationship IDs to new ones
        """
        relationship_mapping = {}

        try:
            if not hasattr(source_slide, 'part') or not hasattr(target_slide, 'part'):
                logger.warning("Slides missing part attribute for relationship copying")
                return relationship_mapping

            source_part = source_slide.part
            target_part = target_slide.part

            # Copy each relationship
            for rel_id, relationship in source_part.rels.items():
                try:
                    if exclude_notes and "notesSlide" in relationship.reltype:
                        logger.debug(f"Skipping notes slide relationship: {rel_id}")
                        continue

                    # Skip slideLayout relationships - target slide already has its own layout
                    # Copying these causes duplicate layout files in the ZIP
                    if "slideLayout" in relationship.reltype:
                        logger.debug(f"Skipping slideLayout relationship: {rel_id}")
                        continue

                    new_rel_id = self._copy_single_relationship(
                        relationship,
                        source_part,
                        target_part,
                        rel_id
                    )

                    if new_rel_id:
                        relationship_mapping[rel_id] = new_rel_id
                        logger.debug(f"Copied relationship {rel_id} -> {new_rel_id}")

                except Exception as e:
                    logger.warning(f"Failed to copy relationship {rel_id}: {e}")
                    continue

        except Exception as e:
            logger.error(f"Error copying slide relationships: {e}")

        return relationship_mapping

    def _copy_single_relationship(
        self,
        relationship,
        source_part: SlidePart,
        target_part: SlidePart,
        original_rel_id: str
    ) -> Optional[str]:
        """
        Copy a single relationship from source to target part.

        Args:
            relationship: The relationship object to copy
            source_part: Source slide part
            target_part: Target slide part
            original_rel_id: Original relationship ID

        Returns:
            New relationship ID if successful, None otherwise
        """
        try:
            rel_type = relationship.reltype
            target_part_obj = relationship._target

            # Handle different relationship types
            if "image" in rel_type.lower():
                return self._copy_image_relationship(
                    relationship, source_part, target_part, original_rel_id
                )
            elif "chart" in rel_type.lower():
                return self._copy_chart_relationship(
                    relationship, source_part, target_part, original_rel_id
                )
            elif "hyperlink" in rel_type.lower():
                return self._copy_hyperlink_relationship(
                    relationship, source_part, target_part, original_rel_id
                )
            else:
                # Generic relationship copying
                return self._copy_generic_relationship(
                    relationship, source_part, target_part, original_rel_id
                )

        except Exception as e:
            logger.warning(f"Failed to copy relationship {original_rel_id}: {e}")
            return None

    def _copy_image_relationship(
        self,
        relationship,
        source_part: SlidePart,
        target_part: SlidePart,
        original_rel_id: str
    ) -> Optional[str]:
        """
        Copy an image relationship, including the image data.

        Args:
            relationship: The image relationship to copy
            source_part: Source slide part
            target_part: Target slide part
            original_rel_id: Original relationship ID

        Returns:
            New relationship ID if successful, None otherwise
        """
        try:
            image_part = relationship._target

            if not hasattr(image_part, 'blob'):
                logger.warning(f"Image part {original_rel_id} has no blob data")
                return None

            # Get image data and wrap in seekable BytesIO stream
            image_data = image_part.blob
            image_stream = io.BytesIO(image_data)

            # Use the correct SlidePart API - returns (ImagePart, rId_string)
            new_image_part, new_rel_id = target_part.get_or_add_image_part(image_stream)

            logger.debug(f"Copied image relationship {original_rel_id} -> {new_rel_id}")
            return new_rel_id

        except Exception as e:
            logger.error(f"Failed to copy image relationship {original_rel_id}: {e}")
            return None

    def _copy_chart_relationship(
        self,
        relationship,
        source_part: SlidePart,
        target_part: SlidePart,
        original_rel_id: str
    ) -> Optional[str]:
        """
        Copy a chart relationship with all embedded data.

        Charts consist of:
        - Main chart XML (chartN.xml)
        - Chart style XML (styleN.xml)
        - Chart color style XML (colorsN.xml)
        - Embedded Excel workbook (Microsoft_Excel_Worksheet.xlsx)

        Args:
            relationship: The chart relationship to copy
            source_part: Source slide part
            target_part: Target slide part
            original_rel_id: Original relationship ID

        Returns:
            New relationship ID if successful, None otherwise
        """
        try:
            source_chart_part = relationship._target
            package = target_part.package

            # 1. Generate new partname for the chart
            new_chart_partname = package.next_partname('/ppt/charts/chart%d.xml')

            # 2. Create new ChartPart with copied blob
            new_chart_part = ChartPart.load(
                partname=new_chart_partname,
                content_type=source_chart_part.content_type,
                package=package,
                blob=source_chart_part.blob
            )

            # 3. Copy chart's sub-relationships (style, colors, Excel)
            for sub_rel_id, sub_rel in source_chart_part.rels.items():
                sub_target = sub_rel._target
                if not hasattr(sub_target, 'blob'):
                    logger.debug(f"Skipping sub-relationship {sub_rel_id} without blob")
                    continue

                # Generate appropriate partname based on relationship type
                if 'chartStyle' in sub_rel.reltype:
                    new_sub_partname = package.next_partname('/ppt/charts/style%d.xml')
                elif 'chartColorStyle' in sub_rel.reltype:
                    new_sub_partname = package.next_partname('/ppt/charts/colors%d.xml')
                elif 'package' in sub_rel.reltype:
                    # Embedded Excel - use embeddings folder
                    new_sub_partname = package.next_partname(
                        '/ppt/embeddings/Microsoft_Excel_Worksheet%d.xlsx'
                    )
                else:
                    logger.warning(f"Unknown chart sub-relationship type: {sub_rel.reltype}")
                    continue

                # Create new sub-part
                new_sub_part = Part.load(
                    partname=new_sub_partname,
                    content_type=sub_target.content_type,
                    package=package,
                    blob=sub_target.blob
                )

                # Relate new chart to new sub-part
                new_chart_part.relate_to(new_sub_part, sub_rel.reltype)
                logger.debug(
                    f"Copied chart sub-relationship: {sub_rel.reltype} -> {new_sub_partname}"
                )

            # 4. Create relationship from target slide to new chart
            new_rel_id = target_part.relate_to(new_chart_part, relationship.reltype)

            logger.info(
                f"Successfully copied chart relationship {original_rel_id} -> {new_rel_id} "
                f"({source_chart_part.partname} -> {new_chart_partname})"
            )
            return new_rel_id

        except Exception as e:
            logger.error(f"Failed to copy chart relationship {original_rel_id}: {e}")
            return None

    def _copy_hyperlink_relationship(
        self,
        relationship,
        source_part: SlidePart,
        target_part: SlidePart,
        original_rel_id: str
    ) -> Optional[str]:
        """
        Copy a hyperlink relationship.

        Args:
            relationship: The hyperlink relationship to copy
            source_part: Source slide part
            target_part: Target slide part
            original_rel_id: Original relationship ID

        Returns:
            New relationship ID if successful, None otherwise
        """
        try:
            # Get the target URL
            if hasattr(relationship, '_target_ref'):
                target_url = relationship._target_ref
            else:
                target_url = str(relationship._target)

            # Create new hyperlink relationship
            new_rel_id = target_part.relate_to(target_url, relationship.reltype)

            logger.debug(f"Copied hyperlink relationship {original_rel_id} -> {new_rel_id}")
            return new_rel_id

        except Exception as e:
            logger.error(f"Failed to copy hyperlink relationship {original_rel_id}: {e}")
            return None

    def _copy_generic_relationship(
        self,
        relationship,
        source_part: SlidePart,
        target_part: SlidePart,
        original_rel_id: str
    ) -> Optional[str]:
        """
        Copy a generic relationship.

        Args:
            relationship: The relationship to copy
            source_part: Source slide part
            target_part: Target slide part
            original_rel_id: Original relationship ID

        Returns:
            New relationship ID if successful, None otherwise
        """
        try:
            # For generic relationships, try to create a relationship to the same target
            target_obj = relationship._target
            rel_type = relationship.reltype

            new_rel_id = target_part.relate_to(target_obj, rel_type)

            logger.debug(f"Copied generic relationship {original_rel_id} -> {new_rel_id}")
            return new_rel_id

        except Exception as e:
            logger.warning(f"Failed to copy generic relationship {original_rel_id}: {e}")
            return None

    def update_relationship_references(
        self,
        slide_element,
        relationship_mapping: Dict[str, str]
    ) -> bool:
        """
        Update relationship references in slide XML elements.

        Args:
            slide_element: The slide XML element to update
            relationship_mapping: Mapping of old rel IDs to new rel IDs

        Returns:
            True if successful, False otherwise
        """
        try:
            if not slide_element or not relationship_mapping:
                return False

            # Find all r:id attributes in the XML
            nsmap = {'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            rel_id_elements = slide_element.xpath('.//@r:id', namespaces=nsmap)

            updated_count = 0
            for element in rel_id_elements:
                if hasattr(element, 'getparent'):
                    parent = element.getparent()
                    old_rel_id = parent.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                    if old_rel_id in relationship_mapping:
                        new_rel_id = relationship_mapping[old_rel_id]
                        parent.set(
                            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',
                            new_rel_id
                        )
                        updated_count += 1

            logger.debug(f"Updated {updated_count} relationship references")
            return True

        except Exception as e:
            logger.error(f"Failed to update relationship references: {e}")
            return False

    def copy_notes_slide_relationships(
        self,
        source_slide: Slide,
        target_slide: Slide
    ) -> bool:
        """
        Copy notes slide relationships if present.

        Args:
            source_slide: Source slide with notes
            target_slide: Target slide to copy notes to

        Returns:
            True if successful, False otherwise
        """
        try:
            if not source_slide.has_notes_slide:
                return True  # No notes to copy

            source_notes = source_slide.notes_slide
            target_notes = target_slide.notes_slide

            if not source_notes.notes_text_frame or not target_notes.notes_text_frame:
                return True  # No text frames to copy

            # Copy the text content
            source_text = source_notes.notes_text_frame.text
            if source_text.strip():
                target_notes.notes_text_frame.text = source_text

            logger.debug("Copied notes slide content")
            return True

        except Exception as e:
            logger.warning(f"Failed to copy notes slide relationships: {e}")
            return False

    def cleanup(self):
        """Clean up temporary files created during copying."""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                logger.warning(f"Failed to clean up temp file {temp_file}: {e}")

        self.temp_files.clear()
        self.copied_relationships.clear()

    def get_relationship_stats(self) -> Dict[str, int]:
        """
        Get statistics about copied relationships.

        Returns:
            Dictionary with relationship copying statistics
        """
        return {
            'copied_relationships': len(self.copied_relationships),
            'temp_files_created': len(self.temp_files),
        }