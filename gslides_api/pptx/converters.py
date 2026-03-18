"""Converters between PowerPoint font styles and platform-agnostic styles.

This module provides bidirectional conversion between python-pptx font objects
and the platform-agnostic MarkdownRenderableStyle/RichStyle classes from gslides-api.
"""

from typing import List, Optional

from pptx.dml.color import RGBColor
from pptx.table import Table
from pptx.text.text import TextFrame
from pptx.util import Pt

from gslides_api.agnostic.text import (
    AbstractColor,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    RichStyle,
)

# XML namespace for DrawingML (used for bullet detection)
_DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


# Monospace font families for code detection
MONOSPACE_FONTS = {
    "courier new",
    "courier",
    "monospace",
    "consolas",
    "monaco",
    "lucida console",
    "dejavu sans mono",
    "source code pro",
    "fira code",
    "jetbrains mono",
}


def _is_monospace(font_name: Optional[str]) -> bool:
    """Check if a font name is a monospace font."""
    if not font_name:
        return False
    return font_name.lower() in MONOSPACE_FONTS


def _pptx_color_to_abstract(pptx_color) -> Optional[AbstractColor]:
    """Convert python-pptx color to AbstractColor.

    Args:
        pptx_color: python-pptx font.color object

    Returns:
        AbstractColor or None if no color set
    """
    try:
        if pptx_color is None:
            return None
        rgb = pptx_color.rgb
        if rgb is None:
            return None
        # RGBColor is a tuple-like (r, g, b) with 0-255 values
        return AbstractColor(
            red=rgb[0] / 255,
            green=rgb[1] / 255,
            blue=rgb[2] / 255,
        )
    except (AttributeError, TypeError):
        return None


def _abstract_to_pptx_rgb(abstract_color: Optional[AbstractColor]) -> Optional[RGBColor]:
    """Convert AbstractColor to python-pptx RGBColor.

    Args:
        abstract_color: AbstractColor object

    Returns:
        RGBColor or None
    """
    if abstract_color is None:
        return None
    r, g, b = abstract_color.to_rgb_tuple()
    return RGBColor(r=r, g=g, b=b)


def _pptx_baseline_to_abstract(font) -> BaselineOffset:
    """Convert python-pptx font baseline to abstract BaselineOffset.

    Args:
        font: python-pptx font object

    Returns:
        BaselineOffset enum value
    """
    try:
        if getattr(font, "superscript", False):
            return BaselineOffset.SUPERSCRIPT
        if getattr(font, "subscript", False):
            return BaselineOffset.SUBSCRIPT
    except AttributeError:
        pass
    return BaselineOffset.NONE


def pptx_font_to_full(font, hyperlink_address: Optional[str] = None) -> FullTextStyle:
    """Extract full FullTextStyle from python-pptx font object.

    Args:
        font: python-pptx font object
        hyperlink_address: Optional hyperlink URL

    Returns:
        FullTextStyle with both markdown and rich parts
    """
    # Extract markdown-renderable properties
    markdown = MarkdownRenderableStyle(
        bold=getattr(font, "bold", False) or False,
        italic=getattr(font, "italic", False) or False,
        strikethrough=getattr(font, "strike", False) or False,
        is_code=_is_monospace(getattr(font, "name", None)),
        hyperlink=hyperlink_address,
    )

    # Get font size in points
    font_size_pt = None
    try:
        if font.size is not None:
            font_size_pt = font.size.pt
    except (AttributeError, TypeError):
        pass

    # Extract rich properties (non-markdown-renderable)
    rich = RichStyle(
        font_family=getattr(font, "name", None),
        font_size_pt=font_size_pt,
        font_weight=None,  # python-pptx doesn't expose weight directly
        foreground_color=_pptx_color_to_abstract(getattr(font, "color", None)),
        background_color=None,  # highlight_color handled separately in pptx
        underline=getattr(font, "underline", False) or False,
        small_caps=getattr(font, "small_caps", False) or False,
        all_caps=getattr(font, "all_caps", False) or False,
        baseline_offset=_pptx_baseline_to_abstract(font),
        character_spacing=None,  # Not commonly used
        shadow=getattr(font, "shadow", False) or False,
        emboss=getattr(font, "emboss", False) or False,
        imprint=getattr(font, "imprint", False) or False,
        double_strike=getattr(font, "double_strike", False) or False,
    )

    return FullTextStyle(markdown=markdown, rich=rich)


def pptx_font_to_rich(font) -> RichStyle:
    """Extract only RichStyle from python-pptx font (for styles() method).

    This is used by the styles() method to get only the non-markdown-renderable
    properties for uniqueness checking.

    Args:
        font: python-pptx font object

    Returns:
        RichStyle with only non-markdown-renderable properties
    """
    return pptx_font_to_full(font).rich


def apply_rich_style_to_pptx_run(rich: RichStyle, run) -> None:
    """Apply RichStyle properties to a python-pptx run.

    Args:
        rich: RichStyle with non-markdown-renderable properties
        run: python-pptx run object
    """
    if rich.font_family is not None:
        run.font.name = rich.font_family

    if rich.font_size_pt is not None:
        run.font.size = Pt(rich.font_size_pt)

    if rich.foreground_color is not None:
        run.font.color.rgb = _abstract_to_pptx_rgb(rich.foreground_color)

    # Apply boolean properties definitively (False means "not formatted")
    run.font.underline = rich.underline
    run.font.small_caps = rich.small_caps
    run.font.all_caps = rich.all_caps
    run.font.shadow = rich.shadow
    run.font.emboss = rich.emboss
    run.font.imprint = rich.imprint
    run.font.double_strike = rich.double_strike

    # Handle baseline offset - explicitly set both properties
    if rich.baseline_offset == BaselineOffset.SUPERSCRIPT:
        run.font.superscript = True
        run.font.subscript = False
    elif rich.baseline_offset == BaselineOffset.SUBSCRIPT:
        run.font.subscript = True
        run.font.superscript = False
    else:  # BaselineOffset.NONE
        run.font.superscript = False
        run.font.subscript = False


def apply_markdown_style_to_pptx_run(md: MarkdownRenderableStyle, run) -> None:
    """Apply MarkdownRenderableStyle to a python-pptx run.

    Args:
        md: MarkdownRenderableStyle with markdown-derivable properties
        run: python-pptx run object
    """
    # Apply boolean properties definitively (False means "not formatted")
    run.font.bold = md.bold
    run.font.italic = md.italic
    run.font.strike = md.strikethrough

    # Set Courier New for code if no font family already set
    if md.is_code and not run.font.name:
        run.font.name = "Courier New"

    if md.hyperlink:
        run.hyperlink.address = md.hyperlink


def apply_full_style_to_pptx_run(style: FullTextStyle, run) -> None:
    """Apply complete FullTextStyle to a python-pptx run.

    Args:
        style: FullTextStyle with both markdown and rich parts
        run: python-pptx run object
    """
    apply_rich_style_to_pptx_run(rich=style.rich, run=run)
    apply_markdown_style_to_pptx_run(md=style.markdown, run=run)


# =============================================================================
# Markdown Generation Functions (PPT -> Markdown)
# =============================================================================


def _escape_markdown_for_table(text: str) -> str:
    """Escape text for use in markdown table cells.

    Only escapes pipe characters and converts newlines to <br> tags.
    Does NOT escape curly braces to preserve template variables like {var}.

    Args:
        text: Raw text content

    Returns:
        Text safe for markdown table cells
    """
    # Escape pipe characters (table cell delimiters)
    text = text.replace("|", "\\|")
    # Convert newlines to HTML break tags for multiline cells
    text = text.replace("\n", "<br>")
    return text


def pptx_run_to_markdown(run) -> str:
    """Convert a python-pptx run to markdown string.

    Uses pptx_font_to_full() to extract styling and applies markdown formatting.

    Args:
        run: python-pptx run object

    Returns:
        Markdown-formatted string
    """
    text = run.text
    if not text:
        return ""

    # Get hyperlink if present
    hyperlink_address = None
    if run.hyperlink and run.hyperlink.address:
        hyperlink_address = run.hyperlink.address

    # Extract style using existing converter
    style = pptx_font_to_full(run.font, hyperlink_address=hyperlink_address)
    md = style.markdown

    # Apply formatting in correct order (inner to outer)
    # Code formatting (backticks)
    if md.is_code:
        text = f"`{text}`"
    # Bold + Italic combined
    elif md.bold and md.italic:
        text = f"***{text}***"
    # Bold only
    elif md.bold:
        text = f"**{text}**"
    # Italic only
    elif md.italic:
        text = f"*{text}*"

    # Strikethrough (can combine with other formatting)
    if md.strikethrough:
        text = f"~~{text}~~"

    # Hyperlink wraps everything
    if md.hyperlink:
        text = f"[{text}]({md.hyperlink})"

    return text


def _paragraph_has_bullet(paragraph) -> bool:
    """Check if a paragraph has bullet formatting in XML.

    A paragraph has bullets if:
    - buNone is NOT present (which explicitly disables bullets), AND
    - Either buChar (character bullet) or buAutoNum (numbered) is present

    Args:
        paragraph: python-pptx paragraph object

    Returns:
        True if paragraph has bullet formatting, False otherwise
    """
    try:
        # Get paragraph properties element
        pPr = paragraph._element.get_or_add_pPr()

        # If buNone exists, bullets are explicitly disabled
        if pPr.find(f"{{{_DRAWINGML_NS}}}buNone") is not None:
            return False

        # Check for actual bullet elements
        if pPr.find(f"{{{_DRAWINGML_NS}}}buChar") is not None:
            return True
        if pPr.find(f"{{{_DRAWINGML_NS}}}buAutoNum") is not None:
            return True

        return False
    except Exception:
        return False


def pptx_paragraph_to_markdown(paragraph) -> str:
    """Convert a python-pptx paragraph to markdown string.

    Handles bullet points and list indentation.

    Args:
        paragraph: python-pptx paragraph object

    Returns:
        Markdown-formatted paragraph string
    """
    # Build paragraph from runs
    parts: List[str] = []
    for run in paragraph.runs:
        parts.append(pptx_run_to_markdown(run))

    text = "".join(parts)

    # Handle bullet points / list indentation
    # Check if paragraph has actual bullet formatting in XML
    has_bullet = _paragraph_has_bullet(paragraph)
    level = getattr(paragraph, "level", 0) or 0

    if has_bullet:
        # Apply bullet formatting with appropriate indentation
        indent = "  " * level
        text = f"{indent}- {text}"

    return text


def pptx_text_frame_to_markdown(text_frame: Optional[TextFrame]) -> str:
    """Convert a python-pptx text frame to markdown string.

    Args:
        text_frame: python-pptx TextFrame object (can be None)

    Returns:
        Markdown-formatted string with paragraphs joined by newlines
    """
    if not text_frame:
        return ""

    lines: List[str] = []
    for paragraph in text_frame.paragraphs:
        para_text = pptx_paragraph_to_markdown(paragraph)
        # Only add non-empty paragraphs (but preserve intentional empty lines)
        if para_text.strip() or lines:  # Include empty lines after first content
            lines.append(para_text)

    return "\n".join(lines)


def pptx_table_to_markdown(table: Table) -> str:
    """Convert a python-pptx Table to markdown table string.

    Args:
        table: python-pptx Table object

    Returns:
        Markdown table string
    """
    if not table or not table.rows:
        return ""

    rows: List[List[str]] = []

    for row in table.rows:
        row_cells: List[str] = []
        for cell in row.cells:
            # Extract text from cell's text frame
            cell_text = pptx_text_frame_to_markdown(cell.text_frame)
            # Escape for table cell usage
            cell_text = _escape_markdown_for_table(cell_text)
            row_cells.append(cell_text)
        rows.append(row_cells)

    if not rows:
        return ""

    # Build markdown table
    lines: List[str] = []

    # Header row (first row)
    header = rows[0]
    lines.append("| " + " | ".join(header) + " |")

    # Separator row
    separator = "| " + " | ".join(["---"] * len(header)) + " |"
    lines.append(separator)

    # Data rows
    for row in rows[1:]:
        # Ensure row has same number of columns as header
        while len(row) < len(header):
            row.append("")
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)
