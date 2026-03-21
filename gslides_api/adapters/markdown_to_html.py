"""Convert markdown to HTML formatted text using BeautifulSoup.

This module uses the shared markdown parser from gslides-api to convert markdown text
to HTML elements with proper formatting tags.
"""

import logging
from typing import Optional

from bs4 import BeautifulSoup, Tag

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.markdown_parser import parse_markdown_to_ir
from gslides_api.agnostic.text import FullTextStyle

logger = logging.getLogger(__name__)


def apply_markdown_to_html_element(
    markdown_text: str,
    html_element: Tag,
    base_style: Optional[FullTextStyle] = None,
) -> None:
    """Apply markdown formatting to an HTML element.

    Args:
        markdown_text: The markdown text to convert
        html_element: The BeautifulSoup Tag to write to (e.g., <div>, <p>)
        base_style: Optional base text style (from gslides-api TextStyle).
            NOTE: Only RichStyle properties (font_family, font_size, color, underline)
            are applied from base_style. Markdown-renderable properties (bold, italic)
            should come from the markdown content itself (e.g., **bold**, *italic*).

    Note:
        This function clears the existing content of the element before writing.
    """
    # Parse markdown to IR using shared parser
    ir_doc = parse_markdown_to_ir(markdown_text, base_style=base_style)

    # Clear existing content
    html_element.clear()

    # Get soup for creating new tags
    soup = BeautifulSoup("", "lxml")

    # Convert IR to HTML
    _apply_ir_to_html_element(ir_doc, html_element, soup, base_style)


def _apply_ir_to_html_element(
    ir_doc: FormattedDocument,
    html_element: Tag,
    soup: BeautifulSoup,
    base_style: Optional[FullTextStyle] = None,
) -> None:
    """Convert IR document to HTML content."""
    # Process each paragraph/list in the document
    for item in ir_doc.elements:
        if isinstance(item, FormattedParagraph):
            _add_paragraph_to_html(item, html_element, soup, base_style)
        elif isinstance(item, FormattedList):
            _add_list_to_html(item, html_element, soup, base_style)


def _add_paragraph_to_html(
    paragraph: FormattedParagraph,
    parent: Tag,
    soup: BeautifulSoup,
    base_style: Optional[FullTextStyle] = None,
) -> None:
    """Add a paragraph to HTML element."""
    # For HTML, we'll add runs inline without creating separate <p> tags
    # unless it's the first paragraph (to avoid extra spacing)

    for run in paragraph.runs:
        _add_run_to_html(run, parent, soup, base_style)

    # Add line break after paragraph (except if it's the last one)
    # We'll just add the content inline for now and let HTML handle spacing


def _add_run_to_html(
    run: FormattedTextRun,
    parent: Tag,
    soup: BeautifulSoup,
    base_style: Optional[FullTextStyle] = None,
) -> None:
    """Add a text run with formatting to HTML element."""
    text = run.content

    # Apply formatting by wrapping in HTML tags
    # FullTextStyle has markdown (bold, italic, strikethrough) and rich (underline, color, etc.)
    bold = run.style.markdown.bold if run.style.markdown else False
    italic = run.style.markdown.italic if run.style.markdown else False
    strikethrough = run.style.markdown.strikethrough if run.style.markdown else False
    underline = run.style.rich.underline if run.style.rich else False

    if bold and italic:
        # Bold + Italic
        wrapper = soup.new_tag("strong")
        inner = soup.new_tag("em")
        inner.string = text
        wrapper.append(inner)
        parent.append(wrapper)
    elif bold:
        # Bold only
        wrapper = soup.new_tag("strong")
        wrapper.string = text
        parent.append(wrapper)
    elif italic:
        # Italic only
        wrapper = soup.new_tag("em")
        wrapper.string = text
        parent.append(wrapper)
    elif underline:
        # Underline
        wrapper = soup.new_tag("u")
        wrapper.string = text
        parent.append(wrapper)
    elif strikethrough:
        # Strikethrough
        wrapper = soup.new_tag("s")
        wrapper.string = text
        parent.append(wrapper)
    else:
        # Plain text - always append (don't use .string which replaces)
        parent.append(text)


def _add_list_to_html(
    formatted_list: FormattedList,
    parent: Tag,
    soup: BeautifulSoup,
    base_style: Optional[FullTextStyle] = None,
) -> None:
    """Add a formatted list to HTML element."""
    # Create <ul> or <ol> based on list type
    list_tag = soup.new_tag("ol" if formatted_list.ordered else "ul")

    for item in formatted_list.items:
        li_tag = soup.new_tag("li")

        # Add each paragraph's runs to the list item
        for paragraph in item.paragraphs:
            for run in paragraph.runs:
                _add_run_to_html(run, li_tag, soup, base_style)

        list_tag.append(li_tag)

    parent.append(list_tag)


def _process_li_children(li: Tag, depth: int = 0) -> str:
    """Recursively process children of a <li> to preserve inline formatting.

    Unlike get_text(), this preserves bold, italic, etc. as markdown markers.
    """
    parts = []
    for child in li.children:
        if isinstance(child, str):
            parts.append(child.strip())
        elif isinstance(child, Tag):
            tag_name = child.name.lower() if child.name else ""
            if tag_name in ("strong", "b"):
                parts.append(f"**{child.get_text().strip()}**")
            elif tag_name in ("em", "i"):
                parts.append(f"*{child.get_text().strip()}*")
            elif tag_name == "u":
                parts.append(f"__{child.get_text().strip()}__")
            elif tag_name in ("s", "strike", "del"):
                parts.append(f"~~{child.get_text().strip()}~~")
            elif tag_name == "br":
                # Preserve line breaks within list items
                parts.append("\n")
            elif tag_name == "span":
                # Recurse into spans to pick up nested formatting
                parts.append(_process_li_children(child, depth=depth))
            else:
                parts.append(child.get_text().strip())
    return " ".join(part for part in parts if part and part != "\n").strip()


def convert_html_to_markdown(html_element: Tag) -> str:
    """Convert HTML element content to markdown format.

    Args:
        html_element: BeautifulSoup Tag to read from

    Returns:
        Markdown-formatted string
    """

    def process_element(elem, depth=0):
        """Recursively process HTML elements."""
        if isinstance(elem, str):
            return elem

        if elem.name == "strong" or elem.name == "b":
            # Strip whitespace from inside the tag for proper markdown
            return f"**{elem.get_text().strip()}**"
        elif elem.name == "em" or elem.name == "i":
            return f"*{elem.get_text().strip()}*"
        elif elem.name == "u":
            # HTML underline doesn't have direct markdown equivalent
            return f"__{elem.get_text().strip()}__"
        elif elem.name == "s" or elem.name == "strike" or elem.name == "del":
            return f"~~{elem.get_text().strip()}~~"
        elif elem.name == "ul":
            # Bulleted list - recursively process children to preserve formatting
            result = []
            for li in elem.find_all("li", recursive=False):
                li_content = _process_li_children(li, depth=depth)
                result.append(f"{'  ' * depth}- {li_content}")
            return "\n".join(result)
        elif elem.name == "ol":
            # Numbered list - recursively process children to preserve formatting
            result = []
            for i, li in enumerate(elem.find_all("li", recursive=False), 1):
                li_content = _process_li_children(li, depth=depth)
                result.append(f"{'  ' * depth}{i}. {li_content}")
            return "\n".join(result)
        else:
            # For other elements, recursively process children
            result = []
            for child in elem.children:
                processed = process_element(child, depth)
                if processed:
                    result.append(processed)
            return "".join(result)

    return process_element(html_element)
