"""Platform-agnostic representations and utilities."""

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
    IRElementType,
)
from gslides_api.agnostic.markdown_parser import parse_markdown_to_ir

__all__ = [
    "FormattedDocument",
    "FormattedList",
    "FormattedListItem",
    "FormattedParagraph",
    "FormattedTextRun",
    "IRElementType",
    "parse_markdown_to_ir",
]
