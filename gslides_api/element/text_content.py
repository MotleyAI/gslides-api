import logging
from typing import Any, Dict, List, Optional, Tuple

from typeguard import typechecked

from gslides_api.domain.domain import Dimension, GSlidesBaseModel, Unit
from gslides_api.markdown.from_markdown import markdown_to_text_elements, text_elements_to_requests
from gslides_api.domain.request import Range, RangeType
from gslides_api.domain.table_cell import TableCellLocation
from gslides_api.request.request import (
    DeleteParagraphBulletsRequest,
    DeleteTextRequest,
    UpdateTextStyleRequest,
)
from gslides_api.request.parent import GSlidesAPIRequest
from gslides_api.domain.text import TextElement, TextStyle
from gslides_api.markdown.to_markdown import text_elements_to_markdown

logger = logging.getLogger(__name__)


@typechecked
class TextContent(GSlidesBaseModel):
    """Represents text content with its elements and lists."""

    textElements: Optional[List[TextElement]] = None
    lists: Optional[Dict[str, Any]] = None

    @property
    def styles(self) -> List[TextStyle] | None:
        """Extract all unique text styles from the text elements."""
        if not self.textElements:
            return None
        styles = []
        for te in self.textElements:
            if te.textRun is None:
                continue
            if te.textRun.content.strip() == "":
                continue
            if te.textRun.style not in styles:
                styles.append(te.textRun.style)
        return styles

    def to_markdown(self) -> str | None:
        """Convert the shape's text content back to markdown format.

        This method reconstructs markdown from the Google Slides API response,
        handling formatting like bold, italic, bullet points, nested lists, and code spans.
        """
        if not self.textElements:
            return None

        return text_elements_to_markdown(self.textElements)

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
            return self.to_markdown()
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
        styles: List[TextStyle] | None = None,
        overwrite: bool = True,
        autoscale: bool = False,
        size_inches: Tuple[float, float] | None = None,
    ):
        """Convert the text content to a list of requests to update the text in the element.
        IMPORTANT: This does not set the objectId on the requests as the container doesn't know it,
        so the caller must set it before sending the requests, ditto for CellLocation if needed.
        """
        styles = styles or self.styles

        if autoscale:
            if size_inches is None:
                raise ValueError("size_inches must be provided if autoscale is True")
            styles = self.autoscale_text(text, size_inches, styles)

        if self.has_text and overwrite:
            requests = self.delete_text_request()
        else:
            requests = []

        style_args = {}
        if styles is not None:
            if len(styles) == 1:
                style_args["base_style"] = styles[0]
            elif len(styles) > 1:
                style_args["heading_style"] = styles[0]
                style_args["base_style"] = styles[1]

        requests += markdown_to_text_elements(text, **style_args)

        # TODO: this is broken, we should use different logic to just dump raw text, asterisks, hashes and all
        if not as_markdown:
            requests = [r for r in requests if not isinstance(r, UpdateTextStyleRequest)]

        return requests

    def autoscale_text(
        self,
        text: str,
        size_inches: Tuple[float, float],
        styles: List[TextStyle] | None = None,
    ) -> List[TextStyle]:
        # Unfortunately, GS

        # For now, just derive the scaling factor based on first style
        if not styles or len(styles) == 0:
            logger.warning("No styles provided, cannot autoscale text")
            return styles or []

        first_style = styles[0]

        my_width_in, my_height_in = size_inches
        # if location is not None:  # this must be a table, with overridden absolute_size
        #     my_width_in, my_height_in = self.absolute_size(OutputUnit.IN, location=location)
        # else:
        #     my_width_in, my_height_in = self.absolute_size(OutputUnit.IN)

        # Get current font size in points (default to 12pt if not specified)
        current_font_size_pt = 12.0
        if first_style.fontSize and first_style.fontSize.magnitude:
            if first_style.fontSize.unit.value == "PT":
                current_font_size_pt = first_style.fontSize.magnitude
            elif first_style.fontSize.unit.value == "EMU":
                # Convert EMU to points: 1 point = 12,700 EMUs
                current_font_size_pt = first_style.fontSize.magnitude / 12700.0

        # Determine the estimated width of the text based on font size and length
        # Rough approximation: average character width is about 0.6 * font_size_pt / 72 inches
        avg_char_width_in = (current_font_size_pt * 0.6) / 72.0
        line_height_in = (current_font_size_pt * 1.2) / 72.0  # 1.2 line spacing factor

        # Account for some padding/margins (assume 10% on each side)
        usable_width_in = my_width_in * 0.8
        usable_height_in = my_height_in * 0.8

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
            ) ** 0.5  # Square root for more gradual scaling

        # Apply minimum scaling factor to ensure text remains readable
        scaling_factor = max(scaling_factor, 0.3)  # Don't scale below 30% of original size

        # Apply the scaling factor to the font size of ALL styles

        scaled_styles = []

        for style in styles:
            scaled_style = style.model_copy()  # Create a copy to avoid modifying the original

            # Get the current font size for this style
            style_font_size_pt = 12.0  # default
            if scaled_style.fontSize and scaled_style.fontSize.magnitude:
                if scaled_style.fontSize.unit.value == "PT":
                    style_font_size_pt = scaled_style.fontSize.magnitude
                elif scaled_style.fontSize.unit.value == "EMU":
                    # Convert EMU to points: 1 point = 12,700 EMUs
                    style_font_size_pt = scaled_style.fontSize.magnitude / 12700.0

            # Apply scaling factor to this style's font size
            new_font_size_pt = style_font_size_pt * scaling_factor
            scaled_style.fontSize = Dimension(magnitude=new_font_size_pt, unit=Unit.PT)

            scaled_styles.append(scaled_style)

        return scaled_styles
