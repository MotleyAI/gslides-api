"""
XML utilities for safe PowerPoint element manipulation.

Provides safe alternatives to copy.deepcopy() for XML elements to prevent
corruption and ensure proper namespace handling.
"""

import logging
from typing import Optional, Dict, Any
from lxml import etree
from pptx.shapes.base import BaseShape

logger = logging.getLogger(__name__)


class XmlUtils:
    """
    Utilities for safe XML manipulation in PowerPoint documents.

    Provides methods to safely copy XML elements without using deepcopy,
    which can cause corruption in python-pptx.
    """

    # PowerPoint XML namespaces
    NAMESPACES = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'a16': 'http://schemas.microsoft.com/office/drawing/2013/main-command',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    }

    @classmethod
    def safe_copy_element(cls, source_element, new_id: Optional[int] = None) -> etree.Element:
        """
        Safely copy an XML element without using deepcopy.

        This method creates a new element with the same tag, attributes, and children
        as the source, but generates new unique IDs to prevent conflicts.

        Args:
            source_element: The source XML element to copy
            new_id: Optional new ID to assign to the element

        Returns:
            A new XML element that is a safe copy of the source
        """
        if source_element is None:
            raise ValueError("Source element cannot be None")

        try:
            # Create new element with same tag
            new_element = etree.Element(source_element.tag, nsmap=source_element.nsmap)

            # Copy attributes, updating ID if provided
            for key, value in source_element.attrib.items():
                if key == 'id' and new_id is not None:
                    new_element.set(key, str(new_id))
                else:
                    new_element.set(key, value)

            # Copy text content
            if source_element.text:
                new_element.text = source_element.text
            if source_element.tail:
                new_element.tail = source_element.tail

            # Recursively copy children, updating IDs in cNvPr elements
            for child in source_element:
                new_child = cls.safe_copy_element(child)
                new_element.append(new_child)

            # Update ID in cNvPr element if this is a shape and new_id is provided
            if new_id is not None:
                cnv_pr_elements = new_element.xpath('.//p:cNvPr', namespaces=cls.NAMESPACES)
                for cnv_pr in cnv_pr_elements:
                    cnv_pr.set('id', str(new_id))

            return new_element

        except Exception as e:
            logger.error(f"Error copying XML element: {e}")
            raise

    @classmethod
    def update_element_id(cls, element, new_id: int) -> bool:
        """
        Update the ID attribute of an XML element.

        Args:
            element: The XML element to update
            new_id: The new ID value

        Returns:
            True if successful, False otherwise
        """
        try:
            if element is not None and hasattr(element, 'set'):
                element.set('id', str(new_id))
                return True
        except Exception as e:
            logger.warning(f"Failed to update element ID: {e}")

        return False

    @classmethod
    def update_creation_id(cls, element, new_creation_id: str) -> bool:
        """
        Update or add the a16:creationId attribute to prevent corruption.

        Args:
            element: The XML element to update
            new_creation_id: The new creation ID (GUID)

        Returns:
            True if successful, False otherwise
        """
        try:
            if element is None:
                return False

            # Find or create a16:creationId elements
            creation_id_xpath = './/a16:creationId'
            creation_id_elems = element.xpath(creation_id_xpath, namespaces=cls.NAMESPACES)

            if creation_id_elems:
                # Update existing creation ID
                for creation_elem in creation_id_elems:
                    creation_elem.set('id', new_creation_id)
            else:
                # Add new creation ID if not present
                # This is more complex as we need to find the right place to insert it
                # For now, we'll just try to set it on the main element if possible
                cnv_pr_elems = element.xpath('.//p:cNvPr', namespaces=cls.NAMESPACES)
                if cnv_pr_elems:
                    cnv_pr = cnv_pr_elems[0]
                    # Create a16:creationId subelement
                    creation_elem = etree.SubElement(
                        cnv_pr,
                        f"{{{cls.NAMESPACES['a16']}}}creationId",
                        nsmap={'a16': cls.NAMESPACES['a16']}
                    )
                    creation_elem.set('id', new_creation_id)

            return True

        except Exception as e:
            logger.warning(f"Failed to update creation ID: {e}")
            return False

    @classmethod
    def get_element_id(cls, element) -> Optional[int]:
        """
        Get the ID attribute from an XML element.

        Args:
            element: The XML element to get ID from

        Returns:
            The element ID as integer, or None if not found
        """
        try:
            if element is not None:
                id_str = element.get('id')
                if id_str:
                    return int(id_str)
        except (ValueError, AttributeError) as e:
            logger.debug(f"Could not get element ID: {e}")

        return None

    @classmethod
    def get_creation_id(cls, element) -> Optional[str]:
        """
        Get the a16:creationId from an XML element.

        Args:
            element: The XML element to get creation ID from

        Returns:
            The creation ID as string, or None if not found
        """
        try:
            if element is not None:
                creation_id_elems = element.xpath('.//a16:creationId', namespaces=cls.NAMESPACES)
                if creation_id_elems:
                    return creation_id_elems[0].get('id')
        except Exception as e:
            logger.debug(f"Could not get creation ID: {e}")

        return None

    @classmethod
    def validate_element(cls, element) -> Dict[str, Any]:
        """
        Validate an XML element for common issues.

        Args:
            element: The XML element to validate

        Returns:
            Dictionary with validation results
        """
        validation_result = {
            'valid': True,
            'issues': [],
            'warnings': [],
        }

        try:
            if element is None:
                validation_result['valid'] = False
                validation_result['issues'].append('Element is None')
                return validation_result

            # Check for proper namespace declarations
            if not element.nsmap:
                validation_result['warnings'].append('Element missing namespace declarations')

            # Check for creation ID if it's a shape element
            if 'sp' in element.tag or 'pic' in element.tag or 'graphicFrame' in element.tag:
                creation_id = cls.get_creation_id(element)
                if not creation_id:
                    validation_result['warnings'].append('Shape element missing creation ID')

        except Exception as e:
            validation_result['valid'] = False
            validation_result['issues'].append(f'Validation error: {e}')

        return validation_result

    @classmethod
    def clean_element_relationships(cls, element) -> bool:
        """
        Clean relationship references in an XML element.

        This removes or updates relationship IDs that might cause conflicts
        when copying elements between slides.

        Args:
            element: The XML element to clean

        Returns:
            True if successful, False otherwise
        """
        try:
            if element is None:
                return False

            # Find and clean relationship references
            rel_id_xpath = './/@r:id'
            rel_id_attrs = element.xpath(rel_id_xpath, namespaces=cls.NAMESPACES)

            for attr in rel_id_attrs:
                # For now, we'll clear the relationship ID
                # The RelationshipCopier will handle rebuilding them
                if hasattr(attr, 'getparent'):
                    parent = attr.getparent()
                    if parent is not None:
                        # Remove the r:id attribute temporarily
                        parent.attrib.pop(f"{{{cls.NAMESPACES['r']}}}id", None)

            return True

        except Exception as e:
            logger.warning(f"Failed to clean element relationships: {e}")
            return False

    @classmethod
    def remap_element_relationships(
        cls, element, relationship_mapping: Dict[str, str]
    ) -> int:
        """
        Remap relationship references in an XML element using a mapping.

        This updates r:id and r:embed attributes to point to the new relationship IDs
        after relationships have been copied to a new slide.

        Args:
            element: The XML element to update
            relationship_mapping: Dictionary mapping old relationship IDs to new ones

        Returns:
            Number of relationships remapped
        """
        remapped_count = 0
        try:
            if element is None or not relationship_mapping:
                return remapped_count

            # Find all elements with r:id attributes
            r_namespace = cls.NAMESPACES['r']
            r_id_attr = f"{{{r_namespace}}}id"
            r_embed_attr = f"{{{r_namespace}}}embed"
            r_link_attr = f"{{{r_namespace}}}link"

            # Find all r:id attributes
            for attr_name in [r_id_attr, r_embed_attr, r_link_attr]:
                # Search for elements with this attribute
                for elem in element.iter():
                    old_id = elem.get(attr_name)
                    if old_id and old_id in relationship_mapping:
                        new_id = relationship_mapping[old_id]
                        elem.set(attr_name, new_id)
                        remapped_count += 1
                        logger.debug(f"Remapped {attr_name}: {old_id} -> {new_id}")

            return remapped_count

        except Exception as e:
            logger.warning(f"Failed to remap element relationships: {e}")
            return remapped_count

    @classmethod
    def copy_shape_element(
        cls,
        source_shape: BaseShape,
        new_shape_id: int,
        new_creation_id: str,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[etree.Element]:
        """
        Copy a shape's XML element with new IDs.

        Args:
            source_shape: The source shape to copy
            new_shape_id: New unique shape ID
            new_creation_id: New unique creation ID
            relationship_mapping: Optional mapping of old relationship IDs to new ones.
                If provided, relationship references will be remapped instead of cleared.

        Returns:
            The copied XML element with updated IDs, or None if failed
        """
        try:
            if not hasattr(source_shape, '_element'):
                logger.error("Source shape has no _element attribute")
                return None

            source_element = source_shape._element
            if source_element is None:
                logger.error("Source shape element is None")
                return None

            # Create safe copy
            new_element = cls.safe_copy_element(source_element, new_shape_id)

            # Update creation ID
            cls.update_creation_id(new_element, new_creation_id)

            # Handle relationship references
            if relationship_mapping:
                # Remap relationship IDs to new ones
                remapped = cls.remap_element_relationships(new_element, relationship_mapping)
                logger.debug(f"Remapped {remapped} relationship references in shape element")
            else:
                # Clean relationship references (will be rebuilt later - legacy behavior)
                cls.clean_element_relationships(new_element)

            # Validate the result
            validation = cls.validate_element(new_element)
            if not validation['valid']:
                logger.error(f"Copied element validation failed: {validation['issues']}")
                return None

            if validation['warnings']:
                logger.warning(f"Copied element has warnings: {validation['warnings']}")

            return new_element

        except Exception as e:
            logger.error(f"Failed to copy shape element: {e}")
            return None