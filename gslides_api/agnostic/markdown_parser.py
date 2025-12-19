"""Platform-agnostic markdown parser that converts markdown to intermediate representation.

This module uses marko to parse markdown and converts it to a platform-agnostic IR
that can then be converted to Google Slides, PowerPoint, or other formats.
"""

import copy
import logging
from typing import Any, Optional

import marko

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedListItem,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.text import FullTextStyle

logger = logging.getLogger(__name__)


def parse_markdown_to_ir(
    markdown_text: str,
    base_style: Optional[FullTextStyle] = None,
    heading_style: Optional[FullTextStyle] = None,
) -> FormattedDocument:
    """Parse markdown string into platform-agnostic intermediate representation.

    Args:
        markdown_text: The markdown text to parse
        base_style: Optional base style to apply to all text
        heading_style: Optional style to apply to headings

    Returns:
        FormattedDocument containing the parsed and styled content
    """
    base_style = base_style or FullTextStyle()

    if heading_style is None:
        heading_style = copy.deepcopy(base_style)
        heading_style.markdown.bold = True

    # Parse markdown with marko
    doc = marko.Markdown().parse(markdown_text)

    # Convert AST to IR
    return _markdown_ast_to_ir(doc, base_style=base_style, heading_style=heading_style)


def _markdown_ast_to_ir(
    markdown_ast: Any,
    base_style: Optional[FullTextStyle] = None,
    heading_style: Optional[FullTextStyle] = None,
    list_depth: int = 0,
) -> FormattedDocument:
    """Convert marko AST to platform-agnostic IR.

    Args:
        markdown_ast: Marko AST node to convert
        base_style: Base text style
        heading_style: Heading text style
        list_depth: Current nesting level for lists

    Returns:
        FormattedDocument with parsed content
    """
    base_style = base_style or FullTextStyle()
    if heading_style is None:
        heading_style = copy.deepcopy(base_style)
        heading_style.markdown.bold = True

    document = FormattedDocument()

    if not isinstance(markdown_ast, marko.block.Document):
        # If not a document, wrap in document
        logger.warning(f"Expected Document, got {type(markdown_ast)}, wrapping")
        temp_doc = marko.block.Document()
        temp_doc.children = [markdown_ast] if not isinstance(markdown_ast, list) else markdown_ast
        markdown_ast = temp_doc

    # Process each child of the document
    for child in markdown_ast.children:
        elements = _process_ast_node(child, base_style, heading_style, list_depth)
        document.elements.extend(elements)

    return document


def _process_ast_node(
    node: Any,
    base_style: FullTextStyle,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> list[FormattedParagraph | FormattedList]:
    """Process a single AST node and return IR elements.

    Args:
        node: Marko AST node
        base_style: Base text style
        heading_style: Heading text style
        list_depth: Current list nesting depth

    Returns:
        List of IR elements (paragraphs or lists)
    """
    if isinstance(node, marko.block.Paragraph):
        return [_process_paragraph(node, base_style, heading_style, list_depth)]

    elif isinstance(node, marko.block.Heading):
        return [_process_heading(node, heading_style, list_depth)]

    elif isinstance(node, marko.block.List):
        return [_process_list(node, base_style, heading_style, list_depth)]

    elif isinstance(node, marko.block.BlankLine):
        # Blank lines create empty paragraphs
        return [FormattedParagraph(runs=[])]

    else:
        logger.warning(f"Unsupported block element: {type(node)}")
        return []


def _process_paragraph(
    para: marko.block.Paragraph,
    base_style: FullTextStyle,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> FormattedParagraph:
    """Process a paragraph node into a FormattedParagraph.

    Args:
        para: Marko paragraph node
        base_style: Base text style
        heading_style: Heading style
        list_depth: Current list depth

    Returns:
        FormattedParagraph with styled text runs
    """
    runs = []
    for child in para.children:
        runs.extend(_process_inline_node(child, base_style, heading_style, list_depth))

    return FormattedParagraph(runs=runs, is_heading=False)


def _process_heading(
    heading: marko.block.Heading,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> FormattedParagraph:
    """Process a heading node into a FormattedParagraph with heading flag.

    Args:
        heading: Marko heading node
        heading_style: Heading text style
        list_depth: Current list depth

    Returns:
        FormattedParagraph marked as heading
    """
    runs = []
    for child in heading.children:
        runs.extend(_process_inline_node(child, heading_style, heading_style, list_depth))

    return FormattedParagraph(
        runs=runs,
        is_heading=True,
        heading_level=heading.level if hasattr(heading, 'level') else 1
    )


def _process_list(
    list_node: marko.block.List,
    base_style: FullTextStyle,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> FormattedList:
    """Process a list node into a FormattedList.

    Args:
        list_node: Marko list node
        base_style: Base text style
        heading_style: Heading style
        list_depth: Current list depth

    Returns:
        FormattedList with list items
    """
    items = []
    for child in list_node.children:
        if isinstance(child, marko.block.ListItem):
            # _process_list_item returns a list (main item + nested items)
            items.extend(_process_list_item(child, base_style, heading_style, list_depth))

    return FormattedList(
        items=items,
        ordered=list_node.ordered if hasattr(list_node, 'ordered') else False,
        style=base_style.rich if base_style else None
    )


def _process_list_item(
    list_item: marko.block.ListItem,
    base_style: FullTextStyle,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> list[FormattedListItem]:
    """Process a list item node into FormattedListItems.

    Args:
        list_item: Marko list item node
        base_style: Base text style
        heading_style: Heading style
        list_depth: Current list depth

    Returns:
        List of FormattedListItem objects - the main item plus any nested items
    """
    paragraphs = []
    nested_items = []

    for child in list_item.children:
        if isinstance(child, marko.block.Paragraph):
            paragraphs.append(_process_paragraph(child, base_style, heading_style, list_depth + 1))
        elif isinstance(child, marko.block.List):
            # Nested list - process and keep items with their correct nesting levels
            nested_list = _process_list(child, base_style, heading_style, list_depth + 1)
            nested_items.extend(nested_list.items)
        else:
            logger.warning(f"Unsupported list item child: {type(child)}")

    # Return the main item followed by any nested items
    result = [FormattedListItem(
        paragraphs=paragraphs,
        nesting_level=list_depth
    )]
    result.extend(nested_items)
    return result


def _process_inline_node(
    node: Any,
    base_style: FullTextStyle,
    heading_style: FullTextStyle,
    list_depth: int = 0,
) -> list[FormattedTextRun]:
    """Process an inline node into text runs.

    Args:
        node: Marko inline node
        base_style: Base text style
        heading_style: Heading style
        list_depth: Current list depth

    Returns:
        List of FormattedTextRun objects
    """
    if isinstance(node, marko.inline.RawText):
        return [FormattedTextRun(content=node.children, style=base_style)]

    elif isinstance(node, marko.inline.LineBreak):
        return [FormattedTextRun(content="\n", style=base_style)]

    elif isinstance(node, marko.inline.CodeSpan):
        code_style = copy.deepcopy(base_style)
        code_style.markdown.is_code = True
        code_style.rich.font_family = "Courier New"
        return [FormattedTextRun(content=node.children, style=code_style)]

    elif isinstance(node, marko.inline.Emphasis):
        italic_style = copy.deepcopy(base_style)
        italic_style.markdown.italic = not italic_style.markdown.italic
        runs = []
        for child in node.children:
            runs.extend(_process_inline_node(child, italic_style, heading_style, list_depth))
        return runs

    elif isinstance(node, marko.inline.StrongEmphasis):
        bold_style = copy.deepcopy(base_style)
        bold_style.markdown.bold = True
        runs = []
        for child in node.children:
            runs.extend(_process_inline_node(child, bold_style, heading_style, list_depth))
        return runs

    elif isinstance(node, marko.inline.Link):
        link_style = copy.deepcopy(base_style)
        link_style.markdown.hyperlink = node.dest
        link_style.rich.underline = True
        runs = []
        for child in node.children:
            runs.extend(_process_inline_node(child, link_style, heading_style, list_depth))
        return runs

    else:
        logger.warning(f"Unsupported inline element: {type(node)}")
        return []
