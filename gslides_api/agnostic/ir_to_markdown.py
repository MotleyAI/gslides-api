"""Convert platform-agnostic IR to markdown.

This module provides functions to convert FormattedDocument (IR) to markdown text,
handling:
1. Consolidation of adjacent runs with identical formatting
2. Proper placement of spaces outside markdown markers
3. Support for lists, paragraphs, and various text formatting
"""

import copy
from typing import List, Optional

from gslides_api.agnostic.ir import (
    FormattedDocument,
    FormattedList,
    FormattedParagraph,
    FormattedTextRun,
)
from gslides_api.agnostic.text import FullTextStyle, MarkdownRenderableStyle


def ir_to_markdown(doc: FormattedDocument) -> str:
    """Convert IR document to markdown with proper run consolidation.

    Args:
        doc: The FormattedDocument to convert

    Returns:
        Markdown string representation
    """
    result_lines = []

    for element in doc.elements:
        if isinstance(element, FormattedParagraph):
            line = _paragraph_to_markdown(element)
            if line is not None:
                result_lines.append(line)
            else:
                # Empty paragraph represents a blank line - preserve it
                result_lines.append("")
        elif isinstance(element, FormattedList):
            list_lines = _list_to_markdown(element)
            result_lines.extend(list_lines)

    return "\n".join(result_lines).rstrip() if result_lines else ""


def _paragraph_to_markdown(para: FormattedParagraph) -> Optional[str]:
    """Convert a paragraph to markdown.

    Args:
        para: The paragraph to convert

    Returns:
        Markdown string for the paragraph, or None if empty
    """
    if not para.runs:
        return None

    # Step 1: Consolidate adjacent runs with identical formatting
    consolidated = _consolidate_runs(para.runs)

    # Step 2: Format each consolidated run to markdown
    parts = []
    for run in consolidated:
        formatted = _format_run_to_markdown(run.content, run.style)
        parts.append(formatted)

    result = "".join(parts).rstrip()
    return result if result else None


def _list_to_markdown(list_element: FormattedList) -> List[str]:
    """Convert a list to markdown lines.

    Args:
        list_element: The list to convert

    Returns:
        List of markdown lines
    """
    lines = []
    # Track item numbers per nesting level (each level has its own counter)
    item_numbers_by_level = {}

    for item in list_element.items:
        nesting_level = item.nesting_level
        indent = "    " * nesting_level  # 4 spaces per level

        # Initialize or increment counter for this nesting level
        if nesting_level not in item_numbers_by_level:
            item_numbers_by_level[nesting_level] = 1
        else:
            item_numbers_by_level[nesting_level] += 1

        # Reset counters for all deeper nesting levels when we go back to a higher level
        levels_to_reset = [lvl for lvl in item_numbers_by_level if lvl > nesting_level]
        for lvl in levels_to_reset:
            del item_numbers_by_level[lvl]

        if list_element.ordered:
            bullet = f"{item_numbers_by_level[nesting_level]}. "
        else:
            bullet = "* "

        for i, para in enumerate(item.paragraphs):
            para_md = _paragraph_to_markdown(para)
            if para_md:
                if i == 0:
                    lines.append(f"{indent}{bullet}{para_md}")
                else:
                    # Continuation lines in the same list item
                    lines.append(f"{indent}    {para_md}")

    return lines


def _consolidate_runs(runs: List[FormattedTextRun]) -> List[FormattedTextRun]:
    """Merge adjacent runs with identical markdown-renderable style.

    This prevents issues like `**text1****text2**` when two adjacent bold runs
    are formatted separately.

    Args:
        runs: List of text runs to consolidate

    Returns:
        Consolidated list of runs
    """
    if not runs:
        return []

    result = []
    current = FormattedTextRun(
        content=runs[0].content,
        style=copy.deepcopy(runs[0].style),
    )

    for run in runs[1:]:
        if _same_markdown_style(current.style.markdown, run.style.markdown):
            # Same formatting - merge content
            current.content += run.content
        else:
            # Different formatting - save current and start new
            result.append(current)
            current = FormattedTextRun(
                content=run.content,
                style=copy.deepcopy(run.style),
            )

    # Don't forget the last run
    result.append(current)
    return result


def _same_markdown_style(a: MarkdownRenderableStyle, b: MarkdownRenderableStyle) -> bool:
    """Check if two markdown styles are identical for consolidation purposes.

    Args:
        a: First style
        b: Second style

    Returns:
        True if styles are identical
    """
    return (
        a.bold == b.bold
        and a.italic == b.italic
        and a.strikethrough == b.strikethrough
        and a.hyperlink == b.hyperlink
        and a.is_code == b.is_code
    )


def _extract_whitespace_parts(content: str) -> tuple:
    """Extract leading spaces, inner text, trailing spaces, and trailing newlines.

    This helper ensures consistent whitespace handling across all formatting types
    (hyperlinks, code spans, bold, italic, etc.).

    Args:
        content: The text content to split

    Returns:
        Tuple of (leading_space, text_content, trailing_space, trailing_newlines)
    """
    leading_space = ""
    for char in content:
        if char in " \t":
            leading_space += char
        else:
            break

    temp_content = content.rstrip("\n")
    trailing_newlines = content[len(temp_content):]

    trailing_space = ""
    for char in reversed(temp_content):
        if char in " \t":
            trailing_space = char + trailing_space
        else:
            break

    text_content = content.strip(" \t").rstrip("\n")
    return leading_space, text_content, trailing_space, trailing_newlines


def _format_run_to_markdown(content: str, style: FullTextStyle) -> str:
    """Apply markdown formatting to content, moving spaces outside markers.

    This follows the GSlides pattern of keeping spaces outside markdown markers
    to produce valid markdown like `**text** ` instead of `**text **`.

    Args:
        content: The text content
        style: The full text style

    Returns:
        Markdown-formatted string
    """
    if not style:
        return content

    md = style.markdown

    # Don't apply formatting markers to whitespace-only content
    if not content.strip(" \t\n"):
        return content

    # Extract whitespace parts once for all formatting types
    leading, text, trailing, newlines = _extract_whitespace_parts(content)

    # Handle hyperlinks first (they take precedence)
    if md.hyperlink:
        if text:
            return leading + f"[{text}]({md.hyperlink})" + trailing + newlines
        return content

    # Handle code spans
    if md.is_code or (style.rich.font_family and style.rich.is_monospace()):
        if text:
            return leading + f"`{text}`" + trailing + newlines
        return content

    # Apply formatting only to the text content
    if text:
        # Handle strikethrough first (can combine with other formatting)
        if md.strikethrough:
            text = f"~~{text}~~"

        # Handle combined bold and italic (***text***)
        if md.bold and md.italic:
            text = f"***{text}***"
        # Handle bold only
        elif md.bold:
            text = f"**{text}**"
        # Handle italic only
        elif md.italic:
            text = f"*{text}*"

    # Reconstruct with preserved spacing OUTSIDE markers
    return leading + text + trailing + newlines
