"""Platform-agnostic representations and utilities."""

from gslides_api.agnostic.converters import (
    EMU_PER_CM,
    EMU_PER_INCH,
    EMU_PER_PT,
    full_style_to_gslides,
    gslides_style_to_full,
    gslides_style_to_rich,
    markdown_style_to_gslides,
    rich_style_to_gslides,
)
from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
    IRElementType,
)
from gslides_api.agnostic.markdown_parser import parse_markdown_to_ir
from gslides_api.agnostic.text import (
    AbstractColor,
    AbstractTextRun,
    BaselineOffset,
    FullTextStyle,
    MarkdownRenderableStyle,
    ParagraphAlignment,
    ParagraphStyle,
    RichStyle,
    SpacingValue,
)

__all__ = [
    # EMU conversion constants
    "EMU_PER_CM",
    "EMU_PER_INCH",
    "EMU_PER_PT",
    # Text style classes
    "AbstractColor",
    "AbstractTextRun",
    "BaselineOffset",
    "FullTextStyle",
    "MarkdownRenderableStyle",
    "ParagraphAlignment",
    "ParagraphStyle",
    "RichStyle",
    "SpacingValue",
    # Converters
    "full_style_to_gslides",
    "gslides_style_to_full",
    "gslides_style_to_rich",
    "markdown_style_to_gslides",
    "rich_style_to_gslides",
    # IR classes
    "FormattedDocument",
    "FormattedList",
    "FormattedListItem",
    "FormattedParagraph",
    "FormattedTextRun",
    "IRElementType",
    "parse_markdown_to_ir",
]
