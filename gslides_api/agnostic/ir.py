"""Platform-agnostic intermediate representation for formatted text.

This module provides data structures that represent formatted text in a way
that can be converted to either Google Slides or PowerPoint (or other formats).

The IR uses platform-agnostic style classes:
- FullTextStyle: Complete style with both markdown-renderable and rich parts
- RichStyle: Non-markdown-renderable properties (colors, fonts, etc.)
- MarkdownRenderableStyle: Properties that can be encoded in markdown (bold, etc.)
"""

from enum import Enum
from typing import List, Optional, Union

from pydantic import BaseModel, Field

from gslides_api.agnostic.text import FullTextStyle, RichStyle


class IRElementType(Enum):
    """Type of IR element."""

    PARAGRAPH = "paragraph"
    LIST = "list"
    HEADING = "heading"


class FormattedTextRun(BaseModel):
    """Platform-agnostic representation of a styled text run.

    A text run is a contiguous piece of text with consistent styling.
    The style includes both markdown-renderable properties (bold, italic, etc.)
    and rich properties (colors, fonts, etc.).
    """

    content: str = Field(description="The text content")
    style: FullTextStyle = Field(
        default_factory=FullTextStyle, description="Style applied to this text run"
    )


class FormattedParagraph(BaseModel):
    """Platform-agnostic representation of a paragraph.

    A paragraph contains one or more text runs and ends with a line break.
    """

    runs: List[FormattedTextRun] = Field(
        default_factory=list, description="Text runs in this paragraph"
    )
    is_heading: bool = Field(
        default=False, description="Whether this paragraph is a heading"
    )
    heading_level: Optional[int] = Field(
        default=None, description="Heading level (1-6) if this is a heading"
    )


class FormattedListItem(BaseModel):
    """Platform-agnostic representation of a list item.

    A list item can contain paragraphs and has a nesting level for nested lists.
    """

    paragraphs: List[FormattedParagraph] = Field(
        default_factory=list, description="Paragraphs in this list item"
    )
    nesting_level: int = Field(
        default=0, description="Nesting level (0 = top level, 1 = nested once, etc.)"
    )


class FormattedList(BaseModel):
    """Platform-agnostic representation of a list.

    A list contains list items and can be ordered (numbered) or unordered (bullets).
    The style is RichStyle (non-markdown-renderable) since markdown formatting
    is stored in the markdown string itself.
    """

    items: List[FormattedListItem] = Field(
        default_factory=list, description="Items in this list"
    )
    ordered: bool = Field(default=False, description="Whether this is an ordered list")
    style: Optional[RichStyle] = Field(
        default=None, description="Base rich style applied to all items in the list"
    )


class FormattedDocument(BaseModel):
    """Platform-agnostic representation of a formatted document.

    A document is a sequence of elements (paragraphs and lists).
    """

    elements: List[Union[FormattedParagraph, FormattedList]] = Field(
        default_factory=list, description="Elements in this document"
    )
