import logging
from typing import Any, Dict, List, Optional, Tuple

from typeguard import typechecked

from gslides_api.agnostic.converters import (
    gslides_style_to_rich,
    rich_style_to_gslides,
    text_elements_to_ir,
)
from gslides_api.agnostic.ir_to_markdown import ir_to_markdown
from gslides_api.agnostic.text import RichStyle
from gslides_api.domain.domain import Dimension, GSlidesBaseModel, Unit
from gslides_api.domain.request import Range, RangeType
from gslides_api.domain.table_cell import TableCellLocation
from gslides_api.domain.text import TextElement, TextStyle
from gslides_api.markdown.from_markdown import markdown_to_text_elements, text_elements_to_requests
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.request.request import (
    DeleteParagraphBulletsRequest,
    DeleteTextRequest,
    UpdateTextStyleRequest,
)

logger = logging.getLogger(__name__)


@typechecked
class TextContent(GSlidesBaseModel):
    """Represents text content with its elements and lists."""

    textElements: Optional[List[TextElement]] = None
    lists: Optional[Dict[str, Any]] = None

    def styles(self, skip_whitespace: bool = True) -> List[RichStyle] | None:
        """Extract all unique RichStyle objects from the text elements.

        Returns only RichStyle (non-markdown-renderable properties like colors, fonts).
        Text that differs only in bold/italic/strikethrough is considered ONE style,
        since those properties are stored in the markdown string itself.

        Args:
            skip_whitespace: If True, skip text runs that contain only whitespace.
                           If False, include styles from whitespace-only text runs.
        """
        if not self.textElements:
            return None
        styles = []
        for te in self.textElements:
            if te.textRun is None:
                continue
            if skip_whitespace and te.textRun.content.strip() == "":
                continue
            rich_style = gslides_style_to_rich(te.textRun.style)
            if rich_style not in styles:
                styles.append(rich_style)
        return styles

    def to_requests(
        self, element_id: str, location: TableCellLocation | None = None
    ) -> List[GSlidesAPIRequest]:
        """Convert the text content to a list of requests to update the text in the element."""
        requests, _ = text_elements_to_requests(self.textElements, [], element_id)
        for r in requests:
            if hasattr(r, "cellLocation"):
                r.cellLocation = location
        return requests

    @property
    def has_text(self):
        return len(self.textElements) > 0 and self.textElements[-1].endIndex > 0

    def read_text(self, as_markdown: bool = True) -> str:
        if not self.has_text:
            return ""
        if as_markdown:
            if not self.textElements:
                return ""

            # Convert to IR first, then to markdown (uses run consolidation)
            ir_doc = text_elements_to_ir(self.textElements)
            return ir_to_markdown(ir_doc)
        else:
            out = []
            for te in self.textElements:
                if te.textRun is not None:
                    out.append(te.textRun.content)
                elif te.paragraphMarker is not None:
                    if len(out) > 0:
                        out.append("\n")
            return "".join(out)

    def delete_text_request(self, object_id: str = "") -> List[GSlidesAPIRequest]:
        """Convert the text content to a list of requests to delete the text in the element.

        Args:
            object_id: The objectId to set on the requests. If empty, caller must set it later.
        """

        # If there are any bullets, need to delete them first
        out: list[GSlidesAPIRequest] = []
        if self.lists is not None and len(self.lists) > 0:
            out.append(
                DeleteParagraphBulletsRequest(
                    objectId=object_id,
                    textRange=Range(type=RangeType.ALL),
                ),
            )

        if (not self.textElements) or self.textElements[0].endIndex == 0:
            return out

        out.append(DeleteTextRequest(objectId=object_id, textRange=Range(type=RangeType.ALL)))
        return out

    def write_text_requests(
        self,
        text: str,
        as_markdown: bool = True,
        styles: List[RichStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        size_inches: Tuple[float, float] | None = None,
    ):
        """Convert the text content to a list of requests to update the text in the element.

        Args:
            text: The text content to write (can be markdown if as_markdown=True)
            as_markdown: If True, parse text as markdown and apply formatting
            styles: List of RichStyle objects to apply. If None, uses self.styles().
                   RichStyle contains non-markdown properties (colors, fonts, etc.)
                   Markdown formatting (bold, italic) is derived from parsing the text.
            overwrite: If True, delete existing text before writing
            autoscale: If True, scale font size to fit text in the element
            size_inches: Required if autoscale=True, the size of the element in inches

        IMPORTANT: This does not set the objectId on the requests as the container doesn't know it,
        so the caller must set it before sending the requests, ditto for CellLocation if needed.
        """
        styles = styles or self.styles()

        if autoscale:
            if size_inches is None:
                raise ValueError("size_inches must be provided if autoscale is True")
            styles = self.autoscale_text(text, size_inches, styles)

        if self.has_text and overwrite:
            requests = self.delete_text_request()
        else:
            requests = []

        # Convert RichStyle to TextStyle for the markdown parser
        # The markdown parser will add bold/italic/etc based on the markdown AST
        style_args = {}
        if styles is not None:
            if len(styles) == 1:
                style_args["base_style"] = rich_style_to_gslides(styles[0])
            elif len(styles) > 1:
                style_args["heading_style"] = rich_style_to_gslides(styles[0])
                style_args["base_style"] = rich_style_to_gslides(styles[1])

        requests += markdown_to_text_elements(text, **style_args)

        # TODO: this is broken, we should use different logic to just dump raw text, asterisks, hashes and all
        if not as_markdown:
            requests = [r for r in requests if not isinstance(r, UpdateTextStyleRequest)]

        return requests

    def autoscale_text(
        self,
        text: str,
        size_inches: Tuple[float, float],
        styles: List[RichStyle] | None = None,
    ) -> List[RichStyle]:
        """Scale font sizes in RichStyle objects to fit text in the given dimensions.

        Args:
            text: The text content (used to estimate how much space is needed)
            size_inches: The dimensions (width, height) of the container in inches
            styles: List of RichStyle objects to scale

        Returns:
            New list of RichStyle objects with scaled font sizes
        """
        if not styles or len(styles) == 0:
            logger.warning("No styles provided, cannot autoscale text")
            return styles or []

        first_style = styles[0]

        my_width_in, my_height_in = size_inches

        # Get current font size in points (default to 12pt if not specified)
        current_font_size_pt = first_style.font_size_pt or 12.0

        # Determine the estimated width of the text based on font size and length
        # Rough approximation: average character width is about 0.6 * font_size_pt / 72 inches
        avg_char_width_in = (current_font_size_pt * 0.6) / 72.0
        line_height_in = (current_font_size_pt * 1.2) / 72.0  # 1.2 line spacing factor

        # Account for some padding/margins
        usable_width_in = my_width_in
        usable_height_in = my_height_in

        # Determine how many characters would fit per line at current size
        chars_per_line = int(usable_width_in / avg_char_width_in)

        # Determine how many lines of text would fit in the shape at current size
        lines_that_fit = int(usable_height_in / line_height_in)

        # Calculate total characters that would fit in the box
        total_chars_that_fit = chars_per_line * lines_that_fit

        # Count actual text length (excluding markdown formatting)
        # Simple approximation: remove common markdown characters
        clean_text = text.replace("*", "").replace("_", "").replace("#", "").replace("`", "")
        actual_text_length = len(clean_text)

        # Determine the scaling factor based on the number of characters that would fit in the box overall
        if actual_text_length <= total_chars_that_fit:
            # Text fits, no scaling needed
            scaling_factor = 1.0
        else:
            # Text doesn't fit, scale down
            scaling_factor = (
                total_chars_that_fit / actual_text_length
            ) ** 0.5  # Square root because we're scaling both width and height

        # Apply minimum scaling factor to ensure text remains readable
        scaling_factor = max(scaling_factor, 0.6)  # Don't scale below 60% of original size
        scaling_factor = min(scaling_factor, 1.0)  # Don't scale above original size

        # Apply the scaling factor to the font size of ALL styles
        scaled_styles = []

        for style in styles:
            scaled_style = style.model_copy()  # Create a copy to avoid modifying the original

            # Get the current font size for this style (default to 12pt)
            style_font_size_pt = scaled_style.font_size_pt or 12.0

            # Apply scaling factor to this style's font size
            scaled_style.font_size_pt = style_font_size_pt * scaling_factor

            scaled_styles.append(scaled_style)

        return scaled_styles
