"""
Shape copier for PowerPoint presentations.

Handles copying of individual shapes including text boxes, images, tables,
and other elements while preserving formatting and relationships.
"""

import logging
from copy import deepcopy
from typing import Dict, List, Optional, Any, Tuple

from pptx.shapes.base import BaseShape
from pptx.shapes.picture import Picture
from pptx.shapes.graphfrm import GraphicFrame
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Emu

from .xml_utils import XmlUtils
from .id_manager import IdManager

logger = logging.getLogger(__name__)

# XML namespaces for shape properties
_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PRESENTATIONML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Fill element tags to copy
_FILL_TAGS = [
    f"{{{_DRAWINGML_NS}}}solidFill",
    f"{{{_DRAWINGML_NS}}}gradFill",
    f"{{{_DRAWINGML_NS}}}pattFill",
    f"{{{_DRAWINGML_NS}}}blipFill",
    f"{{{_DRAWINGML_NS}}}noFill",
]

# Border element tags to copy for table cells
_BORDER_TAGS = [
    f"{{{_DRAWINGML_NS}}}lnL",      # Left border
    f"{{{_DRAWINGML_NS}}}lnR",      # Right border
    f"{{{_DRAWINGML_NS}}}lnT",      # Top border
    f"{{{_DRAWINGML_NS}}}lnB",      # Bottom border
    f"{{{_DRAWINGML_NS}}}lnTlToBr", # Diagonal top-left to bottom-right
    f"{{{_DRAWINGML_NS}}}lnBlToTr", # Diagonal bottom-left to top-right
]

# Table cell properties attributes to copy
_TCPR_ATTRS = [
    "marL",    # Left margin
    "marR",    # Right margin
    "marT",    # Top margin
    "marB",    # Bottom margin
    "anchor",  # Vertical alignment
    "anchorCtr",  # Horizontal center anchor
    "horzOverflow",  # Horizontal overflow
    "vert",    # Text direction
]

# bodyPr attributes to copy for text positioning
_BODYPR_ATTRS = [
    "anchor",           # Vertical alignment (t, ctr, b)
    "anchorCtr",        # Horizontal center anchor
    "lIns",             # Left inset (internal margin)
    "rIns",             # Right inset
    "tIns",             # Top inset
    "bIns",             # Bottom inset
    "wrap",             # Text wrapping mode
    "rtlCol",           # Right-to-left column mode
    "rot",              # Text rotation
    "spcFirstLastPara", # Honor spcBef/spcAft on first/last paragraph
]


class ShapeCopier:
    """
    Handles copying of individual shapes between slides.

    Provides robust copying of different shape types while maintaining
    formatting and handling edge cases that could cause corruption.
    """

    def __init__(self, id_manager: IdManager):
        """
        Initialize shape copier.

        Args:
            id_manager: ID manager for generating unique IDs
        """
        self.id_manager = id_manager
        self.copied_shapes: List[Dict[str, Any]] = []

    def copy_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """
        Copy a shape from source to target slide.

        Args:
            source_shape: The shape to copy
            target_slide: The slide to copy the shape to
            position_offset: Optional (x, y) offset for positioning
            relationship_mapping: Optional mapping of old relationship IDs to new ones.
                Required for proper image handling in GROUP shapes and generic XML copying.

        Returns:
            The newly created shape, or None if copying failed
        """
        try:
            shape_type = source_shape.shape_type
            logger.debug(f"Copying shape type: {shape_type}")

            # Dispatch to appropriate copying method based on shape type
            if shape_type == MSO_SHAPE_TYPE.TEXT_BOX or (
                shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and hasattr(source_shape, 'text_frame')
            ):
                return self._copy_text_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.PICTURE:
                return self._copy_image_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.TABLE:
                return self._copy_table_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                return self._copy_auto_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
                return self._copy_placeholder_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.GROUP:
                return self._copy_group_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            elif shape_type == MSO_SHAPE_TYPE.FREEFORM:
                return self._copy_freeform_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

            else:
                logger.warning(f"Unsupported shape type for copying: {shape_type}")
                return self._copy_generic_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )

        except Exception as e:
            logger.error(f"Failed to copy shape: {e}")
            return None

    def _copy_text_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """Copy a text box or text-containing auto shape."""
        try:
            # Get position and size
            left, top, width, height = self._get_shape_geometry(source_shape, position_offset)

            # Create new text box
            text_box = target_slide.shapes.add_textbox(left, top, width, height)

            # Copy text content and formatting
            if hasattr(source_shape, 'text_frame') and source_shape.text_frame:
                self._copy_text_frame(source_shape.text_frame, text_box.text_frame)

            # Copy other shape properties
            self._copy_shape_properties(source_shape, text_box, relationship_mapping)

            logger.debug("Successfully copied text shape")
            return text_box

        except Exception as e:
            logger.error(f"Failed to copy text shape: {e}")
            return None

    def _copy_image_shape(
        self,
        source_shape: Picture,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """
        Copy an image shape preserving all blipFill properties.

        Uses XML-level copying to preserve critical properties that are lost
        when using python-pptx's add_picture():
        - srcRect (crop settings)
        - rotWithShape attribute
        - alphaModFix and other blip effects
        - stretch mode settings

        Args:
            source_shape: The PICTURE shape to copy
            target_slide: The slide to copy the shape to
            position_offset: Optional (x, y) offset for positioning
            relationship_mapping: Mapping of old relationship IDs to new ones

        Returns:
            None (PICTURE shapes are copied via XML and don't return a shape object)
        """
        try:
            logger.debug("Copying PICTURE shape via XML to preserve blipFill properties")

            # Generate new IDs
            new_shape_id = self.id_manager.generate_unique_shape_id()
            new_creation_id = self.id_manager.generate_unique_creation_id()

            # Copy XML element with relationship remapping
            new_element = XmlUtils.copy_shape_element(
                source_shape,
                new_shape_id,
                new_creation_id,
                relationship_mapping=relationship_mapping
            )

            if new_element is not None:
                # Apply position offset if specified
                if position_offset:
                    self._apply_position_offset_to_element(new_element, position_offset)

                # Insert into target slide
                target_slide.shapes._spTree.insert_element_before(new_element, 'p:extLst')
                logger.debug("Successfully copied PICTURE shape via XML")
                return None  # XML-copied shapes don't return a shape object

            return None

        except Exception as e:
            logger.error(f"Failed to copy image shape: {e}")
            return None

    def _copy_table_shape(
        self,
        source_shape: GraphicFrame,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """Copy a table shape."""
        try:
            if not hasattr(source_shape, 'table'):
                logger.warning("Source shape is not a table")
                return None

            source_table = source_shape.table
            rows = len(source_table.rows)
            cols = len(source_table.columns)

            # Get position and size
            left, top, width, height = self._get_shape_geometry(source_shape, position_offset)

            # Create new table
            table_shape = target_slide.shapes.add_table(rows, cols, left, top, width, height)
            target_table = table_shape.table

            # Copy table-level style properties (first_row, horz_banding, etc.)
            # This prevents default PowerPoint styling from overriding the source table's look
            self._copy_table_style_properties(source_table, target_table)

            # Copy table content with formatting preserved
            for row_idx in range(rows):
                for col_idx in range(cols):
                    if row_idx < len(source_table.rows) and col_idx < len(source_table.columns):
                        source_cell = source_table.cell(row_idx, col_idx)
                        target_cell = target_table.cell(row_idx, col_idx)

                        # Copy cell text WITH formatting (not just plain text)
                        self._copy_text_frame(
                            source_cell.text_frame, target_cell.text_frame
                        )

                        # Copy cell fill/background if present
                        try:
                            self._copy_cell_fill(source_cell, target_cell)
                        except Exception:
                            pass

            # Copy column widths
            for col_idx in range(min(cols, len(source_table.columns))):
                if col_idx < len(target_table.columns):
                    target_table.columns[col_idx].width = source_table.columns[col_idx].width

            # Copy row heights
            for row_idx in range(min(rows, len(source_table.rows))):
                if row_idx < len(target_table.rows):
                    target_table.rows[row_idx].height = source_table.rows[row_idx].height

            # Copy other shape properties
            self._copy_shape_properties(source_shape, table_shape, relationship_mapping)

            logger.debug("Successfully copied table shape")
            return table_shape

        except Exception as e:
            logger.error(f"Failed to copy table shape: {e}")
            return None

    def _copy_auto_shape(
        self,
        source_shape: Shape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """Copy an auto shape (geometric shape)."""
        try:
            # Get position and size
            left, top, width, height = self._get_shape_geometry(source_shape, position_offset)

            # Get auto shape type
            if hasattr(source_shape, 'auto_shape_type'):
                auto_shape_type = source_shape.auto_shape_type
            else:
                # Default to rectangle if we can't determine the type
                from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
                auto_shape_type = MSO_AUTO_SHAPE_TYPE.RECTANGLE

            # Create new auto shape
            auto_shape = target_slide.shapes.add_shape(auto_shape_type, left, top, width, height)

            # Copy text if present
            if hasattr(source_shape, 'text_frame') and source_shape.text_frame:
                self._copy_text_frame(source_shape.text_frame, auto_shape.text_frame)

            # Copy other shape properties
            self._copy_shape_properties(source_shape, auto_shape, relationship_mapping)

            logger.debug("Successfully copied auto shape")
            return auto_shape

        except Exception as e:
            logger.error(f"Failed to copy auto shape: {e}")
            return None

    def _copy_placeholder_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """Copy a placeholder shape."""
        try:
            # Placeholder copying is complex because they're tied to slide layouts
            # For now, we'll convert placeholders to regular text boxes
            if hasattr(source_shape, 'text_frame') and source_shape.text_frame:
                return self._copy_text_shape(
                    source_shape, target_slide, position_offset, relationship_mapping
                )
            else:
                logger.warning("Placeholder shape has no text frame, skipping")
                return None

        except Exception as e:
            logger.error(f"Failed to copy placeholder shape: {e}")
            return None

    def _copy_group_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """
        Copy a GROUP shape including all nested shapes.

        GROUP shapes (device mockups, etc.) contain nested shapes that may include
        images with r:embed references. We use XML-level copying with proper
        relationship remapping to preserve all nested content.

        Args:
            source_shape: The GROUP shape to copy
            target_slide: The slide to copy the shape to
            position_offset: Optional (x, y) offset for positioning
            relationship_mapping: Mapping of old relationship IDs to new ones

        Returns:
            None (GROUP shapes are copied via XML and don't return a shape object)
        """
        try:
            logger.debug(f"Copying GROUP shape with {len(list(source_shape.shapes))} nested shapes")

            # Generate new IDs for the group shape itself
            new_shape_id = self.id_manager.generate_unique_shape_id()
            new_creation_id = self.id_manager.generate_unique_creation_id()

            # Copy XML element with relationship remapping
            new_element = XmlUtils.copy_shape_element(
                source_shape,
                new_shape_id,
                new_creation_id,
                relationship_mapping=relationship_mapping
            )

            if new_element is not None:
                # Regenerate ALL nested shape IDs to avoid conflicts
                # GROUP shapes can have deeply nested structures
                cnv_pr_elements = new_element.xpath(
                    './/p:cNvPr', namespaces=XmlUtils.NAMESPACES
                )
                for cnv_pr in cnv_pr_elements:
                    nested_id = self.id_manager.generate_unique_shape_id()
                    cnv_pr.set('id', str(nested_id))

                # Insert into target slide
                target_slide.shapes._spTree.insert_element_before(new_element, 'p:extLst')
                logger.debug(f"Successfully copied GROUP shape with {len(cnv_pr_elements)} nested elements")
                return None  # We can't return the actual shape object for XML-copied shapes

            return None

        except Exception as e:
            logger.error(f"Failed to copy GROUP shape: {e}")
            return None

    def _copy_freeform_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """
        Copy a FREEFORM shape (custom geometry paths).

        FREEFORM shapes use custom geometry (a:custGeom) and may have
        image fills or other relationships that need remapping.

        Args:
            source_shape: The FREEFORM shape to copy
            target_slide: The slide to copy the shape to
            position_offset: Optional (x, y) offset for positioning
            relationship_mapping: Mapping of old relationship IDs to new ones

        Returns:
            None (FREEFORM shapes are copied via XML and don't return a shape object)
        """
        try:
            logger.debug("Copying FREEFORM shape via XML")

            # Generate new IDs
            new_shape_id = self.id_manager.generate_unique_shape_id()
            new_creation_id = self.id_manager.generate_unique_creation_id()

            # Copy XML element with relationship remapping
            new_element = XmlUtils.copy_shape_element(
                source_shape,
                new_shape_id,
                new_creation_id,
                relationship_mapping=relationship_mapping
            )

            if new_element is not None:
                # Insert into target slide
                target_slide.shapes._spTree.insert_element_before(new_element, 'p:extLst')
                logger.debug("Successfully copied FREEFORM shape via XML")
                return None  # We can't return the actual shape object for XML-copied shapes

            return None

        except Exception as e:
            logger.error(f"Failed to copy FREEFORM shape: {e}")
            return None

    def _copy_generic_shape(
        self,
        source_shape: BaseShape,
        target_slide: Slide,
        position_offset: Optional[Tuple[float, float]] = None,
        relationship_mapping: Optional[Dict[str, str]] = None
    ) -> Optional[BaseShape]:
        """Copy a generic shape using XML manipulation."""
        try:
            # This is a fallback for unsupported shape types
            # We'll use XML copying as a last resort
            logger.warning(f"Using generic XML copying for shape type: {source_shape.shape_type}")

            # Generate new IDs
            new_shape_id = self.id_manager.generate_unique_shape_id()
            new_creation_id = self.id_manager.generate_unique_creation_id()

            # Copy XML element with relationship remapping
            new_element = XmlUtils.copy_shape_element(
                source_shape,
                new_shape_id,
                new_creation_id,
                relationship_mapping=relationship_mapping
            )

            if new_element is not None:
                # Insert into target slide
                target_slide.shapes._spTree.insert_element_before(new_element, 'p:extLst')
                logger.debug("Successfully copied generic shape via XML")
                return None  # We can't return the actual shape object in this case

            return None

        except Exception as e:
            logger.error(f"Failed to copy generic shape: {e}")
            return None

    def _get_shape_geometry(
        self,
        shape: BaseShape,
        position_offset: Optional[Tuple[float, float]] = None
    ) -> Tuple[int, int, int, int]:
        """Get shape geometry (left, top, width, height) with optional offset."""
        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height

        if position_offset:
            offset_x, offset_y = position_offset
            left += Inches(offset_x)
            top += Inches(offset_y)

        return left, top, width, height

    def _apply_position_offset_to_element(
        self,
        element,
        position_offset: Tuple[float, float]
    ):
        """
        Apply position offset to an XML element's xfrm (transform).

        Args:
            element: The XML element containing an a:xfrm child
            position_offset: (x, y) offset in inches to apply
        """
        try:
            offset_x, offset_y = position_offset
            offset_x_emu = int(Inches(offset_x))
            offset_y_emu = int(Inches(offset_y))

            # Find the xfrm element (could be a:xfrm or p:xfrm)
            xfrm = element.find(f".//{{{_DRAWINGML_NS}}}xfrm")
            if xfrm is None:
                xfrm = element.find(f".//{{{_PRESENTATIONML_NS}}}xfrm")

            if xfrm is not None:
                off = xfrm.find(f"{{{_DRAWINGML_NS}}}off")
                if off is not None:
                    current_x = int(off.get('x', 0))
                    current_y = int(off.get('y', 0))
                    off.set('x', str(current_x + offset_x_emu))
                    off.set('y', str(current_y + offset_y_emu))
                    logger.debug(
                        f"Applied position offset: ({offset_x}, {offset_y}) inches"
                    )
        except Exception as e:
            logger.warning(f"Could not apply position offset: {e}")

    def _copy_text_frame(self, source_frame, target_frame):
        """Copy text frame content and formatting."""
        try:
            # Clear existing content
            target_frame.clear()

            # Copy text frame-level properties
            # These control text wrapping and auto-sizing behavior
            if source_frame.word_wrap is not None:
                target_frame.word_wrap = source_frame.word_wrap

            # Copy auto_size property
            if hasattr(source_frame, 'auto_size') and source_frame.auto_size is not None:
                target_frame.auto_size = source_frame.auto_size

            # Copy margin properties if available
            for margin in ['margin_top', 'margin_bottom', 'margin_left', 'margin_right']:
                if hasattr(source_frame, margin):
                    source_value = getattr(source_frame, margin)
                    if source_value is not None:
                        setattr(target_frame, margin, source_value)

            # Copy vertical alignment (anchor) via XML
            self._copy_text_frame_anchor(source_frame, target_frame)

            # Copy paragraphs
            for para_idx, source_para in enumerate(source_frame.paragraphs):
                if para_idx == 0:
                    # Use existing first paragraph
                    target_para = target_frame.paragraphs[0]
                else:
                    # Add new paragraph
                    target_para = target_frame.add_paragraph()

                # Copy paragraph properties
                target_para.level = source_para.level
                if hasattr(source_para, 'alignment'):
                    target_para.alignment = source_para.alignment

                # Copy line spacing properties
                try:
                    if source_para.line_spacing is not None:
                        target_para.line_spacing = source_para.line_spacing
                    if source_para.space_before is not None:
                        target_para.space_before = source_para.space_before
                    if source_para.space_after is not None:
                        target_para.space_after = source_para.space_after
                except Exception as e:
                    logger.debug(f"Could not copy line spacing: {e}")

                # Copy XML-level paragraph properties (marL, indent, spcBef, lnSpc, buChar)
                # These are not handled by python-pptx's paragraph properties
                self._copy_paragraph_xml_properties(source_para, target_para)

                # Copy runs
                for run_idx, source_run in enumerate(source_para.runs):
                    if run_idx == 0 and target_para.runs:
                        # Use existing first run
                        target_run = target_para.runs[0]
                    else:
                        # Add new run
                        target_run = target_para.add_run()

                    # Copy text and basic formatting
                    target_run.text = source_run.text

                    # Copy XML-level run properties (bold, italic, color, font family, etc.)
                    # This is needed because python-pptx doesn't expose all XML attributes reliably
                    self._copy_run_xml_properties(source_run, target_run)

                    if hasattr(source_run, 'font') and hasattr(target_run, 'font'):
                        try:
                            # Copy basic font properties
                            if source_run.font.bold is not None:
                                target_run.font.bold = source_run.font.bold
                            if source_run.font.italic is not None:
                                target_run.font.italic = source_run.font.italic
                            if source_run.font.size is not None:
                                target_run.font.size = source_run.font.size
                            if source_run.font.name is not None:
                                target_run.font.name = source_run.font.name

                            # Copy underline
                            if source_run.font.underline is not None:
                                target_run.font.underline = source_run.font.underline

                            # Copy font color (RGB or theme)
                            try:
                                if source_run.font.color.rgb is not None:
                                    target_run.font.color.rgb = source_run.font.color.rgb
                                elif source_run.font.color.theme_color is not None:
                                    target_run.font.color.theme_color = source_run.font.color.theme_color
                            except Exception:
                                pass  # Color copying can fail for various reasons

                        except Exception as e:
                            logger.debug(f"Could not copy font properties: {e}")

        except Exception as e:
            logger.warning(f"Failed to copy text frame formatting: {e}")
            # At least copy the basic text
            try:
                target_frame.text = source_frame.text
            except Exception:
                pass

    def _copy_paragraph_xml_properties(self, source_para, target_para):
        """Copy XML-level paragraph properties that python-pptx doesn't handle.

        This copies paragraph properties that are stored as XML attributes/elements
        but not exposed through python-pptx's paragraph API:
        - marL: left margin for bullet indentation
        - indent: first-line indent (typically negative for hanging indent)
        - spcBef: space before paragraph
        - lnSpc: line spacing
        - buChar: bullet character
        """
        try:
            source_pPr = source_para._element.find(f"{{{_DRAWINGML_NS}}}pPr")
            if source_pPr is None:
                return

            target_pPr = target_para._element.get_or_add_pPr()

            # Copy marL (left margin) attribute - critical for bullet indentation
            if source_pPr.get("marL"):
                target_pPr.set("marL", source_pPr.get("marL"))

            # Copy indent (first-line indent) attribute - typically negative for hanging indent
            if source_pPr.get("indent"):
                target_pPr.set("indent", source_pPr.get("indent"))

            # Copy spcBef element (space before paragraph)
            source_spcBef = source_pPr.find(f"{{{_DRAWINGML_NS}}}spcBef")
            if source_spcBef is not None:
                existing = target_pPr.find(f"{{{_DRAWINGML_NS}}}spcBef")
                if existing is not None:
                    target_pPr.remove(existing)
                target_pPr.append(deepcopy(source_spcBef))

            # Copy lnSpc element (line spacing)
            source_lnSpc = source_pPr.find(f"{{{_DRAWINGML_NS}}}lnSpc")
            if source_lnSpc is not None:
                existing = target_pPr.find(f"{{{_DRAWINGML_NS}}}lnSpc")
                if existing is not None:
                    target_pPr.remove(existing)
                target_pPr.append(deepcopy(source_lnSpc))

            # Copy buChar element (bullet character)
            source_buChar = source_pPr.find(f"{{{_DRAWINGML_NS}}}buChar")
            if source_buChar is not None:
                existing = target_pPr.find(f"{{{_DRAWINGML_NS}}}buChar")
                if existing is not None:
                    target_pPr.remove(existing)
                target_pPr.append(deepcopy(source_buChar))

        except Exception as e:
            logger.debug(f"Could not copy paragraph XML properties: {e}")

    def _copy_run_xml_properties(self, source_run, target_run):
        """Copy XML-level run properties that python-pptx doesn't handle reliably.

        This copies run properties that are stored as XML attributes but may not be
        correctly exposed through python-pptx's Font API due to inheritance rules:
        - b: bold
        - i: italic
        - lang: language
        - u: underline
        - strike: strikethrough
        - cap: capitalization
        - sz: font size
        """
        try:
            source_rPr = source_run._r.find(f"{{{_DRAWINGML_NS}}}rPr")
            if source_rPr is None:
                return

            # Get or create target rPr
            target_rPr = target_run._r.find(f"{{{_DRAWINGML_NS}}}rPr")
            if target_rPr is None:
                # Create rPr element if it doesn't exist
                from lxml import etree
                target_rPr = etree.SubElement(
                    target_run._r, f"{{{_DRAWINGML_NS}}}rPr"
                )
                # Move rPr to be first child (required position)
                target_run._r.insert(0, target_rPr)

            # Copy key attributes
            for attr in ['b', 'i', 'lang', 'u', 'strike', 'cap', 'sz']:
                value = source_rPr.get(attr)
                if value is not None:
                    target_rPr.set(attr, value)

            # Copy solidFill element (text color)
            source_fill = source_rPr.find(f"{{{_DRAWINGML_NS}}}solidFill")
            if source_fill is not None:
                existing = target_rPr.find(f"{{{_DRAWINGML_NS}}}solidFill")
                if existing is not None:
                    target_rPr.remove(existing)
                target_rPr.append(deepcopy(source_fill))

            # Copy latin font element (font family)
            source_latin = source_rPr.find(f"{{{_DRAWINGML_NS}}}latin")
            if source_latin is not None:
                existing = target_rPr.find(f"{{{_DRAWINGML_NS}}}latin")
                if existing is not None:
                    target_rPr.remove(existing)
                target_rPr.append(deepcopy(source_latin))

            # Copy ea (East Asian) font element
            source_ea = source_rPr.find(f"{{{_DRAWINGML_NS}}}ea")
            if source_ea is not None:
                existing = target_rPr.find(f"{{{_DRAWINGML_NS}}}ea")
                if existing is not None:
                    target_rPr.remove(existing)
                target_rPr.append(deepcopy(source_ea))

            # Copy cs (Complex Script) font element
            source_cs = source_rPr.find(f"{{{_DRAWINGML_NS}}}cs")
            if source_cs is not None:
                existing = target_rPr.find(f"{{{_DRAWINGML_NS}}}cs")
                if existing is not None:
                    target_rPr.remove(existing)
                target_rPr.append(deepcopy(source_cs))

            # Copy sym (Symbol) font element
            source_sym = source_rPr.find(f"{{{_DRAWINGML_NS}}}sym")
            if source_sym is not None:
                existing = target_rPr.find(f"{{{_DRAWINGML_NS}}}sym")
                if existing is not None:
                    target_rPr.remove(existing)
                target_rPr.append(deepcopy(source_sym))

        except Exception as e:
            logger.debug(f"Could not copy run XML properties: {e}")

    def _copy_shape_properties(
        self,
        source_shape: BaseShape,
        target_shape: BaseShape,
        relationship_mapping: Optional[Dict[str, str]] = None
    ):
        """Copy basic shape properties like fill, line, etc.

        Args:
            source_shape: The shape to copy properties from
            target_shape: The shape to copy properties to
            relationship_mapping: Optional mapping of old relationship IDs to new ones.
                Required for proper handling of blipFill elements that reference images.
        """
        try:
            # Copy name/title if available
            if hasattr(source_shape, 'name') and hasattr(target_shape, 'name'):
                target_shape.name = source_shape.name

            # Copy rotation
            if hasattr(source_shape, 'rotation') and hasattr(target_shape, 'rotation'):
                target_shape.rotation = source_shape.rotation

            # Copy alt text (title attribute in p:cNvPr XML element)
            # This is critical for element matching during export
            try:
                source_cnvpr = source_shape._element.xpath(".//p:cNvPr")
                target_cnvpr = target_shape._element.xpath(".//p:cNvPr")
                if source_cnvpr and target_cnvpr:
                    source_title = source_cnvpr[0].attrib.get("title")
                    if source_title is not None:  # Copy even if empty string
                        target_cnvpr[0].attrib["title"] = source_title
                        logger.debug(f"Copied alt text: '{source_title}'")
                else:
                    logger.debug(f"cnvpr elements not found: source={bool(source_cnvpr)}, target={bool(target_cnvpr)}")
            except Exception as e:
                logger.warning(f"Could not copy alt text: {e}", exc_info=True)

            # Copy fill properties (background colors, gradients, etc.)
            self._copy_shape_fill(source_shape, target_shape, relationship_mapping)

        except Exception as e:
            logger.debug(f"Could not copy all shape properties: {e}")

    def _copy_shape_fill(
        self,
        source_shape: BaseShape,
        target_shape: BaseShape,
        relationship_mapping: Optional[Dict[str, str]] = None
    ):
        """Copy fill properties (solidFill, gradFill, etc.) from source to target.

        This copies background colors, gradients, and other fill properties
        that are defined in the spPr (shape properties) element.

        Args:
            source_shape: The shape to copy fill from
            target_shape: The shape to copy fill to
            relationship_mapping: Optional mapping of old relationship IDs to new ones.
                Required for proper handling of blipFill elements that reference images.
        """
        try:
            # Get spPr elements - may be under p:spPr or a:spPr depending on shape type
            source_spPr = source_shape._element.find(f".//{{{_PRESENTATIONML_NS}}}spPr")
            if source_spPr is None:
                source_spPr = source_shape._element.find(f".//{{{_DRAWINGML_NS}}}spPr")

            target_spPr = target_shape._element.find(f".//{{{_PRESENTATIONML_NS}}}spPr")
            if target_spPr is None:
                target_spPr = target_shape._element.find(f".//{{{_DRAWINGML_NS}}}spPr")

            if source_spPr is None or target_spPr is None:
                logger.debug("Could not find spPr elements for fill copying")
                return

            # Find and copy fill element
            for fill_tag in _FILL_TAGS:
                source_fill = source_spPr.find(fill_tag)
                if source_fill is not None:
                    # Remove existing fill from target
                    for tag in _FILL_TAGS:
                        existing = target_spPr.find(tag)
                        if existing is not None:
                            target_spPr.remove(existing)

                    # Copy the fill element (deepcopy is safe for individual elements)
                    new_fill = deepcopy(source_fill)

                    # Remap relationship IDs for blipFill elements that reference images
                    if relationship_mapping:
                        remapped = XmlUtils.remap_element_relationships(
                            new_fill, relationship_mapping
                        )
                        if remapped > 0:
                            logger.debug(
                                f"Remapped {remapped} relationship(s) in fill element"
                            )

                    target_spPr.append(new_fill)
                    logger.debug(f"Copied fill element: {fill_tag}")
                    break  # Only one fill type can be active

        except Exception as e:
            logger.debug(f"Could not copy fill: {e}")

    def _copy_table_style_properties(self, source_table, target_table):
        """Copy table-level style properties from source to target table.

        This copies boolean properties that control PowerPoint's built-in table styling:
        - first_row: Header row styling (typically dark background)
        - first_col: First column styling
        - last_row: Last row styling (typically for totals)
        - last_col: Last column styling
        - horz_banding: Alternating row colors
        - vert_banding: Alternating column colors

        Also copies/removes the tableStyleId to match the source table's theme styling.
        """
        try:
            # Copy boolean style properties
            target_table.first_row = source_table.first_row
            target_table.first_col = source_table.first_col
            target_table.last_row = source_table.last_row
            target_table.last_col = source_table.last_col
            target_table.horz_banding = source_table.horz_banding
            target_table.vert_banding = source_table.vert_banding

            # Copy the tableStyleId from source to target (or remove if source doesn't have one)
            try:
                source_tbl = source_table._tbl
                target_tbl = target_table._tbl

                source_tblPr = source_tbl.tblPr
                target_tblPr = target_tbl.tblPr

                if source_tblPr is not None and target_tblPr is not None:
                    # Find tableStyleId in source
                    ns = {"a": _DRAWINGML_NS}
                    source_style_id = source_tblPr.find(f"{{{_DRAWINGML_NS}}}tableStyleId")
                    target_style_id = target_tblPr.find(f"{{{_DRAWINGML_NS}}}tableStyleId")

                    # Remove existing tableStyleId from target
                    if target_style_id is not None:
                        target_tblPr.remove(target_style_id)

                    # Copy source tableStyleId if it exists
                    if source_style_id is not None:
                        new_style_id = deepcopy(source_style_id)
                        target_tblPr.append(new_style_id)
            except Exception as e:
                logger.debug(f"Could not copy tableStyleId: {e}")

        except Exception as e:
            logger.debug(f"Could not copy table style properties: {e}")

    def _copy_cell_fill(self, source_cell, target_cell):
        """Copy cell properties from source table cell to target table cell.

        This copies:
        - Fill properties (solidFill, gradFill, etc.)
        - Border properties (lnL, lnR, lnT, lnB)
        - Cell attributes (margins, anchor, etc.)
        """
        try:
            # Table cell properties are in a:tc/a:tcPr
            source_tc = source_cell._tc
            target_tc = target_cell._tc

            source_tcPr = source_tc.find(f"{{{_DRAWINGML_NS}}}tcPr")
            target_tcPr = target_tc.find(f"{{{_DRAWINGML_NS}}}tcPr")

            if source_tcPr is None:
                return  # No properties to copy

            # Create tcPr if it doesn't exist
            if target_tcPr is None:
                from lxml import etree
                target_tcPr = etree.SubElement(target_tc, f"{{{_DRAWINGML_NS}}}tcPr")

            # Copy tcPr attributes (margins, anchor, etc.)
            for attr in _TCPR_ATTRS:
                source_val = source_tcPr.get(attr)
                if source_val is not None:
                    target_tcPr.set(attr, source_val)
                elif attr in target_tcPr.attrib:
                    # Remove attribute if source doesn't have it
                    del target_tcPr.attrib[attr]

            # Copy fill elements
            for fill_tag in _FILL_TAGS:
                source_fill = source_tcPr.find(fill_tag)
                if source_fill is not None:
                    # Remove existing fill from target
                    for tag in _FILL_TAGS:
                        existing = target_tcPr.find(tag)
                        if existing is not None:
                            target_tcPr.remove(existing)

                    # Copy the fill element
                    new_fill = deepcopy(source_fill)
                    target_tcPr.append(new_fill)
                    break  # Only one fill type can be active

            # Copy border elements
            for border_tag in _BORDER_TAGS:
                source_border = source_tcPr.find(border_tag)
                # Remove existing border from target first
                existing = target_tcPr.find(border_tag)
                if existing is not None:
                    target_tcPr.remove(existing)

                # Copy border if source has one
                if source_border is not None:
                    new_border = deepcopy(source_border)
                    target_tcPr.append(new_border)

        except Exception as e:
            logger.debug(f"Could not copy cell properties: {e}")

    def _copy_text_frame_anchor(self, source_frame, target_frame):
        """Copy all bodyPr attributes from source to target text frame.

        bodyPr attributes control text positioning within the text frame:
        - 'anchor' = vertical alignment (t=top, ctr=center, b=bottom)
        - 'anchorCtr' = horizontal center anchor
        - 'lIns', 'rIns', 'tIns', 'bIns' = internal margins (insets in EMUs)
        - 'wrap' = text wrapping mode
        - 'rot' = text rotation

        This property is stored in the bodyPr XML element and is not directly
        exposed as a python-pptx property in all cases.
        """
        try:
            source_bodyPr = source_frame._element.find(f"{{{_DRAWINGML_NS}}}bodyPr")
            target_bodyPr = target_frame._element.find(f"{{{_DRAWINGML_NS}}}bodyPr")

            if source_bodyPr is not None and target_bodyPr is not None:
                # Copy all relevant bodyPr attributes
                for attr in _BODYPR_ATTRS:
                    value = source_bodyPr.get(attr)
                    if value is not None:
                        target_bodyPr.set(attr, value)
                        logger.debug(f"Copied bodyPr.{attr}: {value}")
        except Exception as e:
            logger.debug(f"Could not copy text anchor: {e}")

    def get_copy_stats(self) -> Dict[str, int]:
        """
        Get statistics about shape copying operations.

        Returns:
            Dictionary with copying statistics
        """
        return {
            'shapes_copied': len(self.copied_shapes),
        }

    def clear_stats(self):
        """Clear copying statistics."""
        self.copied_shapes.clear()