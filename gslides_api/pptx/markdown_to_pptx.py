"""Convert markdown to PowerPoint formatted text using python-pptx.

This module uses the shared markdown parser from gslides-api to convert markdown text
to PowerPoint text frames with proper formatting.
"""

import logging
import os
import platform
from typing import Optional

from lxml import etree
from pydantic import BaseModel

from gslides_api.agnostic.ir import FormattedDocument, FormattedList, FormattedParagraph
from gslides_api.agnostic.markdown_parser import parse_markdown_to_ir
from gslides_api.agnostic.text import FullTextStyle, ParagraphStyle, SpacingValue
from pptx.enum.text import MSO_AUTO_SIZE, PP_PARAGRAPH_ALIGNMENT
from pptx.text.text import TextFrame
from pptx.util import Pt
from pptx.dml.color import RGBColor

logger = logging.getLogger(__name__)

# XML namespace for DrawingML (used for bullet manipulation)
_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# EMU constants for bullet spacing (914400 EMUs = 1 inch)
_EMU_PER_INCH = 914400
_BULLET_INDENT_EMU = 342900  # ~0.375 inches - standard bullet indent


def _find_font_file(font_family: str, bold: bool = False, italic: bool = False) -> str | None:
    """Try to find a font file on the system for the given font family.

    python-pptx's fit_text() requires a font file on Linux because it doesn't
    support auto-locating fonts on non-Windows systems.

    Args:
        font_family: The font family name (e.g., "Calibri", "Arial")
        bold: Whether to look for bold variant
        italic: Whether to look for italic variant

    Returns:
        Path to font file if found, None otherwise
    """
    # On Windows, fit_text() can auto-locate fonts, so return None to use default behavior
    if platform.system() == "Windows":
        return None

    # Common font directories on Linux/macOS
    font_dirs = [
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"),
        os.path.expanduser("~/.local/share/fonts"),
    ]
    if platform.system() == "Darwin":  # macOS
        font_dirs.extend(
            [
                "/Library/Fonts",
                "/System/Library/Fonts",
                os.path.expanduser("~/Library/Fonts"),
            ]
        )

    # Map common fonts to system equivalents
    # On macOS: use Helvetica/Arial (built-in)
    # On Linux: use Liberation fonts (metrically compatible with MS fonts)
    if platform.system() == "Darwin":
        font_substitutes = {
            "calibri": ["helvetica", "arial"],
            "arial": ["helvetica"],
            "times new roman": ["times"],
            "times": [],
            "courier new": ["courier"],
            "courier": [],
        }
    else:
        font_substitutes = {
            "calibri": ["liberation"],
            "arial": ["liberation"],
            "helvetica": ["liberation"],
            "times new roman": ["liberation"],
            "times": ["liberation"],
            "courier new": ["liberation"],
            "courier": ["liberation"],
        }

    # Determine style suffix
    if bold and italic:
        style_patterns = ["BoldItalic", "Bold-Italic", "bolditalic", "bold-italic"]
    elif bold:
        style_patterns = ["Bold", "bold"]
    elif italic:
        style_patterns = ["Italic", "italic"]
    else:
        style_patterns = ["Regular", "regular", ""]

    # Search for font file
    font_lower = font_family.lower()
    search_terms = [font_family, font_lower]

    # Add substitutes if available
    for key, substitutes in font_substitutes.items():
        if key in font_lower:
            for substitute in substitutes:
                search_terms.append(substitute)
                search_terms.append(substitute.capitalize())
            break

    for font_dir in font_dirs:
        if not os.path.isdir(font_dir):
            continue

        for root, dirs, files in os.walk(font_dir):
            for filename in files:
                if not filename.endswith((".ttf", ".TTF", ".otf", ".OTF")):
                    continue

                filename_lower = filename.lower()
                # Check if any search term matches
                for term in search_terms:
                    if term.lower() in filename_lower:
                        # Check style
                        for style in style_patterns:
                            if style.lower() in filename_lower or (
                                style == ""
                                and "regular" not in filename_lower
                                and "bold" not in filename_lower
                                and "italic" not in filename_lower
                            ):
                                font_path = os.path.join(root, filename)
                                logger.debug(f"Found font file for '{font_family}': {font_path}")
                                return font_path

    logger.debug(f"Could not find font file for '{font_family}' on this system")
    return None


def _get_max_font_size_from_textframe(text_frame: TextFrame) -> float | None:
    """Get the maximum font size from existing text frame runs.

    Used to cap autoscaling so that font size can only decrease, never increase.

    Args:
        text_frame: The PowerPoint text frame to read from

    Returns:
        Maximum font size in points if found, None otherwise
    """
    max_size = None
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if run.font.size is not None:
                size_pt = run.font.size.pt
                if max_size is None or size_pt > max_size:
                    max_size = size_pt
    return max_size


def _preserve_bodypr_insets(text_frame: TextFrame) -> dict:
    """Preserve bodyPr inset values before clearing text frame.

    bodyPr insets (lIns, rIns, tIns, bIns) control the internal margins
    of the text frame. These may be reset when text_frame.clear() is called,
    so we need to preserve and restore them.

    Args:
        text_frame: The PowerPoint text frame to read from

    Returns:
        Dictionary of inset attribute names to values
    """
    insets = {}
    try:
        bodyPr = text_frame._element.find(f"{{{_DRAWINGML_NS}}}bodyPr")
        if bodyPr is not None:
            for attr in ["lIns", "rIns", "tIns", "bIns"]:
                val = bodyPr.get(attr)
                if val is not None:
                    insets[attr] = val
                    logger.debug(f"Preserved bodyPr.{attr}: {val}")
    except Exception as e:
        logger.debug(f"Could not preserve bodyPr insets: {e}")
    return insets


def _restore_bodypr_insets(text_frame: TextFrame, insets: dict) -> None:
    """Restore bodyPr inset values after clearing text frame.

    Args:
        text_frame: The PowerPoint text frame to restore insets to
        insets: Dictionary of inset attribute names to values
    """
    if not insets:
        return
    try:
        bodyPr = text_frame._element.find(f"{{{_DRAWINGML_NS}}}bodyPr")
        if bodyPr is not None:
            for attr, val in insets.items():
                bodyPr.set(attr, val)
                logger.debug(f"Restored bodyPr.{attr}: {val}")
    except Exception as e:
        logger.debug(f"Could not restore bodyPr insets: {e}")


class PreservedParagraphStyles(BaseModel):
    """Container for preserved paragraph styles from template.

    Template analysis shows consistent pattern across text boxes:
    - Paragraph 0 (title): unique spacing (typically spcBef=0)
    - Paragraphs 1+ (bullets): identical spacing within each box (typically spcBef=9pt)

    This structure stores both styles so they can be applied correctly:
    - first_para_style → applied to output paragraph index 0
    - bullet_style → applied to output paragraphs index 1+
    """

    first_para_style: Optional[ParagraphStyle] = None  # Paragraph index 0 (title)
    bullet_style: Optional[ParagraphStyle] = None  # First bullet paragraph (applies to all 1+)


def _preserve_paragraph_properties(text_frame: TextFrame) -> PreservedParagraphStyles:
    """Preserve paragraph properties for first paragraph and bullet paragraphs.

    Template analysis shows consistent pattern:
    - Paragraph 0 (title): unique spacing (typically spcBef=0)
    - Paragraphs 1+ (bullets): identical spacing within each box (typically spcBef=9pt)

    Args:
        text_frame: The PowerPoint text frame to read from

    Returns:
        PreservedParagraphStyles with first_para_style and bullet_style
    """
    result = PreservedParagraphStyles()

    try:
        paragraphs = list(text_frame.paragraphs)
        logger.debug(f"_preserve_paragraph_properties: scanning {len(paragraphs)} paragraphs")

        for i, para in enumerate(paragraphs):
            pPr = para._element.find(f"{{{_DRAWINGML_NS}}}pPr")
            if pPr is not None:
                style = ParagraphStyle.from_pptx_pPr(pPr=pPr, ns=_DRAWINGML_NS)
                logger.debug(
                    f"Para {i}: marL={style.margin_left}, indent={style.indent}, "
                    f"spcBef={style.space_before}, lnSpc={style.line_spacing}, "
                    f"has_bullet={style.has_bullet_properties()}"
                )

                # First paragraph (index 0) = title/first para style
                if i == 0:
                    result.first_para_style = style
                    logger.debug(f"Preserved FIRST paragraph {i} style")

                # First bullet paragraph = bullet style (applies to all subsequent)
                if style.has_bullet_properties() and result.bullet_style is None:
                    result.bullet_style = style
                    logger.debug(f"Preserved BULLET paragraph {i} style")

                # Early exit if we found both
                if result.first_para_style is not None and result.bullet_style is not None:
                    break
            else:
                logger.debug(f"Para {i}: no pPr element found")

    except Exception as e:
        logger.warning(f"Could not preserve paragraph properties: {e}")

    logger.debug(
        f"Preserved styles: first_para={result.first_para_style is not None}, "
        f"bullet={result.bullet_style is not None}"
    )
    return result


def apply_markdown_to_textframe(
    markdown_text: str,
    text_frame: TextFrame,
    base_style: Optional[FullTextStyle] = None,
    autoscale: bool = False,
) -> None:
    """Apply markdown formatting to a PowerPoint text frame.

    Args:
        markdown_text: The markdown text to convert
        text_frame: The PowerPoint text frame to write to
        base_style: Optional base text style (from gslides-api TextStyle).
            NOTE: Only RichStyle properties (font_family, font_size, color, underline)
            are applied from base_style. Markdown-renderable properties (bold, italic)
            should come from the markdown content itself (e.g., **bold**, *italic*).
        autoscale: Whether to enable PowerPoint autoscaling

    Note:
        This function clears the existing content of the text frame before writing.
    """
    # Parse markdown to IR using shared parser
    ir_doc = parse_markdown_to_ir(markdown_text, base_style=base_style)

    # Capture original font size before clearing (for autoscale cap)
    # Autoscaling should only decrease font size, never increase it
    original_max_font_size = None
    if autoscale:
        original_max_font_size = _get_max_font_size_from_textframe(text_frame)
        # Fall back to base_style font size if no font found in text frame
        if original_max_font_size is None and base_style and base_style.rich.font_size_pt:
            original_max_font_size = base_style.rich.font_size_pt

    # Preserve bodyPr insets before clearing (they may be reset by clear())
    preserved_insets = _preserve_bodypr_insets(text_frame)

    # Preserve paragraph styles (first para + bullet styles) before clearing
    preserved_styles = _preserve_paragraph_properties(text_frame)

    # Clear existing content
    text_frame.clear()

    # Restore bodyPr insets after clearing
    _restore_bodypr_insets(text_frame, preserved_insets)

    # Enable word wrap to ensure text wraps to box width
    text_frame.word_wrap = True

    # Convert IR to PowerPoint operations (writes the text content)
    _apply_ir_to_textframe(ir_doc, text_frame, base_style, preserved_styles)

    # Apply autoscaling AFTER text is written - fit_text() calculates
    # the best font size based on actual text content and shape dimensions.
    # Note: MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE only sets a flag that PowerPoint's
    # rendering engine applies when you edit the text - it doesn't work on file open.
    # fit_text() directly calculates and sets the font size, working immediately.
    if autoscale:
        # Get font family from first paragraph/run
        font_family = "Calibri"  # Default
        if text_frame.paragraphs and text_frame.paragraphs[0].runs:
            first_run = text_frame.paragraphs[0].runs[0]
            if first_run.font.name:
                font_family = first_run.font.name

        try:
            # STEP 1: Save per-run bold/italic state before fit_text overwrites them
            # fit_text() calls _set_font() which sets bold/italic on ALL runs,
            # so we need to save and restore individual run styles.
            saved_styles: list[list[tuple[bool | None, bool | None]]] = []
            for para in text_frame.paragraphs:
                para_styles: list[tuple[bool | None, bool | None]] = []
                for run in para.runs:
                    para_styles.append((run.font.bold, run.font.italic))
                saved_styles.append(para_styles)

            # STEP 2: Call fit_text to calculate and apply best font size
            # Use bold=False, italic=False for font file lookup (sizing only)
            font_file = _find_font_file(font_family, bold=False, italic=False)

            # Use original font size as cap to prevent increasing font size
            # Autoscaling should only decrease, never increase
            # Fall back to 18pt if we couldn't determine original size
            max_size = original_max_font_size if original_max_font_size is not None else 18

            text_frame.fit_text(
                font_family=font_family,
                max_size=max_size,
                bold=False,  # Doesn't matter - we restore after
                italic=False,  # Doesn't matter - we restore after
                font_file=font_file,
            )

            # STEP 3: Restore per-run bold/italic state that fit_text overwrote
            for para_idx, para in enumerate(text_frame.paragraphs):
                if para_idx < len(saved_styles):
                    for run_idx, run in enumerate(para.runs):
                        if run_idx < len(saved_styles[para_idx]):
                            saved_bold, saved_italic = saved_styles[para_idx][run_idx]
                            if saved_bold is not None:
                                run.font.bold = saved_bold
                            if saved_italic is not None:
                                run.font.italic = saved_italic

        except Exception as e:
            # If fit_text fails (e.g., font not found), fall back to setting the flag
            # The flag at least preserves the intent for when file is edited in PowerPoint
            logger.warning(f"fit_text() failed, falling back to auto_size flag: {e}")
            text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _apply_ir_to_textframe(
    doc_ir: FormattedDocument,
    text_frame: TextFrame,
    base_style: Optional[FullTextStyle] = None,
    preserved_styles: Optional[PreservedParagraphStyles] = None,
) -> None:
    """Convert IR to PowerPoint text frame operations.

    Args:
        doc_ir: The intermediate representation document
        text_frame: The PowerPoint text frame to write to
        base_style: Optional base text style
        preserved_styles: Optional PreservedParagraphStyles with first_para_style
            and bullet_style from the original template
    """
    # Start with the first paragraph (text_frame always has one)
    first_para = True
    output_para_index = 0  # Track output paragraph index for style selection

    for element in doc_ir.elements:
        if isinstance(element, FormattedParagraph):
            # Add paragraph (reuse first paragraph, create new ones after)
            if first_para:
                p = text_frame.paragraphs[0]
                first_para = False
            else:
                p = text_frame.add_paragraph()

            # Add runs to the paragraph (handling soft line breaks properly)
            for run_ir in element.runs:
                _add_run_with_soft_breaks(p, run_ir.content, run_ir.style)

            # Apply preserved spacing properties based on output paragraph index
            # First paragraph uses first_para_style, rest use bullet_style
            if preserved_styles is not None:
                if output_para_index == 0:
                    style = preserved_styles.first_para_style
                else:
                    style = preserved_styles.bullet_style
                _apply_paragraph_spacing(p, style)

            output_para_index += 1

        elif isinstance(element, FormattedList):
            # Add list items as paragraphs with bullet formatting
            for item in element.items:
                for para in item.paragraphs:
                    # Create paragraph
                    if first_para:
                        p = text_frame.paragraphs[0]
                        first_para = False
                    else:
                        p = text_frame.add_paragraph()

                    # Set indentation level
                    p.level = item.nesting_level

                    # Enable bullet formatting via XML with proper spacing
                    # Note: Setting p.level alone only sets indentation, not bullet markers
                    # All list items use bullet_style
                    bullet_props = preserved_styles.bullet_style if preserved_styles else None
                    _enable_paragraph_bullets(
                        p,
                        level=item.nesting_level,
                        preserved_props=bullet_props,
                    )

                    # Add runs to the paragraph (handling soft line breaks properly)
                    for run_ir in para.runs:
                        _add_run_with_soft_breaks(p, run_ir.content, run_ir.style)

                    output_para_index += 1


def _set_space_before(pPr, val: str) -> None:
    """Set the space-before (spcBef) property on a paragraph properties element.

    Args:
        pPr: The paragraph properties XML element (<a:pPr>)
        val: The space before value in 100ths of a point (e.g., "900" = 9pt)
    """
    # Remove existing spcBef if present
    existing = pPr.find(f"{{{_DRAWINGML_NS}}}spcBef")
    if existing is not None:
        pPr.remove(existing)

    # Create spcBef with spcPts child: <a:spcBef><a:spcPts val="900"/></a:spcBef>
    spcBef = etree.SubElement(pPr, f"{{{_DRAWINGML_NS}}}spcBef")
    spcPts = etree.SubElement(spcBef, f"{{{_DRAWINGML_NS}}}spcPts")
    spcPts.set("val", val)


def _set_space_after(pPr, val: str) -> None:
    """Set the space-after (spcAft) property on a paragraph properties element.

    Args:
        pPr: The paragraph properties XML element (<a:pPr>)
        val: The space after value in 100ths of a point (e.g., "0" = 0pt)
    """
    # Remove existing spcAft if present
    existing = pPr.find(f"{{{_DRAWINGML_NS}}}spcAft")
    if existing is not None:
        pPr.remove(existing)

    # Create spcAft with spcPts child: <a:spcAft><a:spcPts val="0"/></a:spcAft>
    spcAft = etree.SubElement(pPr, f"{{{_DRAWINGML_NS}}}spcAft")
    spcPts = etree.SubElement(spcAft, f"{{{_DRAWINGML_NS}}}spcPts")
    spcPts.set("val", val)


def _set_line_spacing(pPr, spacing: SpacingValue) -> None:
    """Set the line spacing (lnSpc) property on a paragraph properties element.

    Args:
        pPr: The paragraph properties XML element (<a:pPr>)
        spacing: SpacingValue with either points or percentage
    """
    # Remove existing lnSpc if present
    existing = pPr.find(f"{{{_DRAWINGML_NS}}}lnSpc")
    if existing is not None:
        pPr.remove(existing)

    # Create lnSpc with either spcPct or spcPts child
    lnSpc = etree.SubElement(pPr, f"{{{_DRAWINGML_NS}}}lnSpc")
    if spacing.percentage is not None:
        # Use percentage: <a:lnSpc><a:spcPct val="110000"/></a:lnSpc>
        spcPct = etree.SubElement(lnSpc, f"{{{_DRAWINGML_NS}}}spcPct")
        spcPct.set("val", spacing.to_pptx_pct())
    elif spacing.points is not None:
        # Use points: <a:lnSpc><a:spcPts val="1800"/></a:lnSpc>
        spcPts = etree.SubElement(lnSpc, f"{{{_DRAWINGML_NS}}}spcPts")
        spcPts.set("val", spacing.to_pptx_pts())


def _apply_paragraph_spacing(
    paragraph,
    preserved_props: Optional[ParagraphStyle] = None,
) -> None:
    """Apply preserved spacing properties to a paragraph.

    This applies line spacing, space-before, and space-after properties from a
    ParagraphStyle to any paragraph (both regular and bullet paragraphs).

    IMPORTANT: XML element order in <a:pPr> affects PowerPoint rendering.
    The correct order is: lnSpc, spcBef, spcAft, then bullet elements (buClr, buSzPts, buFont, buChar).
    This function inserts spacing elements at the beginning of pPr to ensure correct order.

    Args:
        paragraph: python-pptx paragraph object
        preserved_props: Optional ParagraphStyle with preserved paragraph properties
            from the original template (space_before, space_after, line_spacing)
    """
    if preserved_props is None:
        return

    pPr = paragraph._element.get_or_add_pPr()

    # Remove existing spacing elements (we'll re-add in correct order)
    for tag in ["lnSpc", "spcBef", "spcAft"]:
        existing = pPr.find(f"{{{_DRAWINGML_NS}}}{tag}")
        if existing is not None:
            pPr.remove(existing)

    # Insert spacing elements in REVERSE order at position 0
    # This results in correct order: lnSpc, spcBef, spcAft, [other elements]

    # 3. Space after (insert at position 0 first, will be pushed to position 2)
    if preserved_props.space_after is not None:
        space_val = preserved_props.space_after.to_pptx_pts()
        spcAft = etree.Element(f"{{{_DRAWINGML_NS}}}spcAft")
        spcPts = etree.SubElement(spcAft, f"{{{_DRAWINGML_NS}}}spcPts")
        spcPts.set("val", space_val)
        pPr.insert(0, spcAft)

    # 2. Space before (insert at position 0, pushes spcAft to position 1)
    if preserved_props.space_before is not None:
        space_val = preserved_props.space_before.to_pptx_pts()
        if space_val != "0":
            spcBef = etree.Element(f"{{{_DRAWINGML_NS}}}spcBef")
            spcPts = etree.SubElement(spcBef, f"{{{_DRAWINGML_NS}}}spcPts")
            spcPts.set("val", space_val)
            pPr.insert(0, spcBef)

    # 1. Line spacing (insert at position 0, pushes others down)
    if preserved_props.line_spacing is not None:
        lnSpc = etree.Element(f"{{{_DRAWINGML_NS}}}lnSpc")
        spacing = preserved_props.line_spacing
        if spacing.percentage is not None:
            spcPct = etree.SubElement(lnSpc, f"{{{_DRAWINGML_NS}}}spcPct")
            spcPct.set("val", spacing.to_pptx_pct())
        elif spacing.points is not None:
            spcPts = etree.SubElement(lnSpc, f"{{{_DRAWINGML_NS}}}spcPts")
            spcPts.set("val", spacing.to_pptx_pts())
        pPr.insert(0, lnSpc)


def _enable_paragraph_bullets(
    paragraph,
    char: str = "•",
    level: int = 0,
    preserved_props: Optional[ParagraphStyle] = None,
) -> None:
    """Enable bullet formatting for a paragraph via XML manipulation.

    In python-pptx, setting paragraph.level only sets indentation, it does NOT
    enable bullet formatting. To actually show bullets, we need to add the
    buChar (bullet character) XML element, remove any buNone element, and set
    proper marL (left margin) and indent (first-line indent) for spacing.

    IMPORTANT: XML element order in <a:pPr> affects PowerPoint rendering.
    The correct order is: lnSpc, spcBef, spcAft, then bullet elements (buClr, buSzPts, buFont, buChar).
    This function applies spacing FIRST (which inserts at position 0), then appends bullet elements.

    Args:
        paragraph: python-pptx paragraph object
        char: Bullet character to use (default: •)
        level: Nesting level for indentation (0 = first level)
        preserved_props: Optional ParagraphStyle with preserved paragraph properties
            from the original template (margins, indents, spacing)
    """
    try:
        # Get or create paragraph properties element
        pPr = paragraph._element.get_or_add_pPr()

        # Remove buNone if it exists (it explicitly disables bullets)
        buNone = pPr.find(f"{{{_DRAWINGML_NS}}}buNone")
        if buNone is not None:
            pPr.remove(buNone)

        # Set margin and indent for hanging bullet effect
        # marL = total left margin (where text starts)
        # indent = first-line offset (negative = hanging indent, pulls bullet left)
        # Use preserved values from template if available, otherwise use defaults
        if preserved_props and preserved_props.margin_left is not None and level == 0:
            # Use preserved values for first-level bullets
            margin_left = str(preserved_props.margin_left)
            indent = (
                str(preserved_props.indent) if preserved_props.indent else str(-_BULLET_INDENT_EMU)
            )
        else:
            # Fall back to default values (or adjust based on level)
            margin_left = str(_BULLET_INDENT_EMU * (level + 1))
            indent = str(-_BULLET_INDENT_EMU)

        pPr.set("marL", margin_left)
        pPr.set("indent", indent)

        # Apply spacing properties FIRST (this inserts lnSpc, spcBef, spcAft at position 0)
        # This ensures spacing elements come BEFORE bullet elements
        _apply_paragraph_spacing(paragraph, preserved_props)

        # Remove existing bullet character if present (we'll re-add at correct position)
        existing_buChar = pPr.find(f"{{{_DRAWINGML_NS}}}buChar")
        if existing_buChar is not None:
            pPr.remove(existing_buChar)

        # Append buChar at the END (after spacing elements)
        # This ensures correct order: lnSpc, spcBef, spcAft, buChar
        buChar = etree.Element(f"{{{_DRAWINGML_NS}}}buChar")
        buChar.set("char", char)
        pPr.append(buChar)

        logger.debug(f"Enabled bullet formatting for paragraph with char='{char}', level={level}")
    except Exception as e:
        logger.warning(f"Could not enable bullet formatting: {e}")


# Soft line break character (vertical tab) - used by PowerPoint for in-cell line breaks
_SOFT_LINE_BREAK = "\x0b"
# Escaped form that python-pptx produces when \x0b is assigned to run.text
_SOFT_LINE_BREAK_ESCAPED = "_x000B_"


def _add_run_with_soft_breaks(paragraph, content: str, style: FullTextStyle) -> None:
    """Add run content to paragraph, handling soft line breaks properly.

    PowerPoint uses vertical tab (\\x0b) for soft line breaks within a paragraph.
    When assigned directly to run.text, python-pptx escapes it to '_x000B_'.
    This function splits on soft breaks and uses add_line_break() to create
    proper <a:br> elements in the XML.

    Args:
        paragraph: The python-pptx paragraph object
        content: The text content (may contain \\x0b or '_x000B_')
        style: The FullTextStyle to apply to the runs
    """
    # Normalize: replace escaped form with actual character for consistent handling
    normalized_content = content.replace(_SOFT_LINE_BREAK_ESCAPED, _SOFT_LINE_BREAK)

    # Split on soft line breaks
    parts = normalized_content.split(_SOFT_LINE_BREAK)

    for i, part in enumerate(parts):
        if part:  # Skip empty parts
            run = paragraph.add_run()
            run.text = part
            _apply_style_to_run(run, style)

        # Add line break after each part except the last
        if i < len(parts) - 1:
            paragraph.add_line_break()


def _apply_style_to_run(run, style: FullTextStyle) -> None:
    """Apply FullTextStyle to a python-pptx run.

    Args:
        run: The python-pptx run object
        style: The FullTextStyle to apply
    """
    if not style:
        return

    # Apply markdown-renderable properties
    md = style.markdown

    # CRITICAL: Always explicitly set bold/italic to avoid inheritance from defRPr
    # When bold=True: set font.bold = True
    # When bold=False: set font.bold = False (prevents inheritance from template)
    run.font.bold = True if md.bold else False
    run.font.italic = True if md.italic else False

    if md.strikethrough and hasattr(run.font, "strikethrough"):
        run.font.strikethrough = True

    if md.hyperlink:
        run.hyperlink.address = md.hyperlink

    # Apply rich properties
    rich = style.rich
    if rich.underline:
        run.font.underline = True

    if rich.font_family:
        run.font.name = rich.font_family

    if rich.font_size_pt:
        run.font.size = Pt(rich.font_size_pt)

    if rich.foreground_color:
        # Convert AbstractColor to RGBColor
        rgb = rich.foreground_color.to_rgb_tuple()
        run.font.color.rgb = RGBColor(*rgb)
