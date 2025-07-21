from typing import List

from gslides_api import TextElement


def normalize_numbered_glyph(glyph: str) -> str:
    """Normalize the glyph for numbered lists."""
    if glyph.endswith("."):
        # Try to extract the number
        number_part = glyph[:-1]
    else:
        number_part = glyph
    latin = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"]
    alpha = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"]

    if number_part in latin:
        out = str(latin.index(number_part) + 1)
    elif number_part in alpha:
        out = str(alpha.index(number_part) + 1)
    elif number_part.isdigit():
        out = number_part
    else:
        raise ValueError(f"Unsupported glyph format: {glyph}")

    return f"{out}. "


def text_elements_to_markdown(elements: List[TextElement]):
    result = []
    current_paragraph = []

    # Track list state for proper nesting
    current_list_info = None  # Will store (listId, nestingLevel, glyph, is_numbered)
    pending_bullet_info = None  # Store bullet info until we get the text content

    for i, te in enumerate(elements):
        # Handle paragraph markers (for bullets and paragraph breaks)
        if te.paragraphMarker is not None:
            # Check if this is a bullet point
            if te.paragraphMarker.bullet is not None:
                bullet = te.paragraphMarker.bullet
                list_id = bullet.listId if hasattr(bullet, "listId") else None
                nesting_level = bullet.nestingLevel if hasattr(bullet, "nestingLevel") else 0
                glyph = bullet.glyph if hasattr(bullet, "glyph") and bullet.glyph else "●"

                # Determine if this is a numbered list based on the glyph
                is_numbered = _is_numbered_list_glyph(glyph)

                # Store the bullet info to be used when we encounter the text content
                pending_bullet_info = (list_id, nesting_level, glyph, is_numbered)
                continue
            else:
                # Regular paragraph marker - clear any pending bullet info
                pending_bullet_info = None
                continue

        # Handle text runs
        if te.textRun is not None:
            content = te.textRun.content
            style = te.textRun.style

            # Handle bullet points - add bullet marker at start of line
            if pending_bullet_info and content.strip() and not current_paragraph:
                list_id, nesting_level, glyph, is_numbered = pending_bullet_info

                # Generate the appropriate indentation and bullet marker
                indent = _get_list_indentation(nesting_level)
                bullet_marker = _format_bullet_marker_with_nesting(glyph)

                current_paragraph.append(indent + bullet_marker)
                current_list_info = pending_bullet_info
                pending_bullet_info = None  # Clear after use

            # Apply formatting based on style
            formatted_content = _apply_markdown_formatting(content, style)
            current_paragraph.append(formatted_content)

            # Handle line breaks
            if "\n" in content:
                # Join current paragraph and add to result
                paragraph_text = "".join(current_paragraph).rstrip()
                if True:  # paragraph_text:
                    result.append(paragraph_text)
                current_paragraph = []

    # Add any remaining paragraph content
    if current_paragraph:
        paragraph_text = "".join(current_paragraph).rstrip()
        if paragraph_text:
            result.append(paragraph_text)

    return "\n".join(result).rstrip() if result else None


def _apply_markdown_formatting(content: str, style) -> str:
    """Apply markdown formatting to content based on text style."""
    if style is None or (len(set(content)) == 1 and next(iter(set(content))) == " "):
        return content

    # Handle hyperlinks first (they take precedence)
    if hasattr(style, "link") and style.link:
        # Handle both dict and object cases
        url = None
        if isinstance(style.link, dict) and "url" in style.link:
            url = style.link["url"]
        elif hasattr(style.link, "url"):
            url = style.link.url

        if url:
            # For links, format as [text](url)
            clean_content = content.strip()
            if clean_content:
                return f"[{clean_content}]({url})"
        return content

    # Handle code spans (different font family)
    if (
        hasattr(style, "fontFamily")
        and style.fontFamily
        and style.fontFamily.lower() in ["courier new", "courier", "monospace"]
    ):
        # For code spans, only format the non-whitespace content
        if content.strip():
            return f"`{content.strip()}`"
        return content

    # For formatting, we need to preserve leading/trailing spaces
    # but only format the actual text content
    leading_space = ""
    trailing_space = ""
    text_content = content

    # Extract leading spaces
    for char in content:
        if char in " \t":
            leading_space += char
        else:
            break

    # Extract trailing spaces (but not newlines)
    temp_content = content.rstrip("\n")
    trailing_newlines = content[len(temp_content) :]

    for char in reversed(temp_content):
        if char in " \t":
            trailing_space = char + trailing_space
        else:
            break

    # Get the actual text content without leading/trailing spaces
    text_content = content.strip(" \t").rstrip("\n")

    # Apply formatting only to the text content
    if text_content:
        # Handle strikethrough first (can combine with other formatting)
        if hasattr(style, "strikethrough") and style.strikethrough:
            text_content = f"~~{text_content}~~"

        # Handle combined bold and italic (***text***)
        if hasattr(style, "bold") and style.bold and hasattr(style, "italic") and style.italic:
            text_content = f"***{text_content}***"
        # Handle bold only
        elif hasattr(style, "bold") and style.bold:
            text_content = f"**{text_content}**"
        # Handle italic only
        elif hasattr(style, "italic") and style.italic:
            text_content = f"*{text_content}*"

    # Reconstruct with preserved spacing
    return leading_space + text_content + trailing_space + trailing_newlines


def _is_numbered_list_glyph(glyph: str) -> bool:
    """Determine if a glyph represents a numbered list item."""
    if not glyph:
        return False

    # Check if the glyph contains digits or letters (indicating numbering)
    return any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph)


def _get_list_indentation(nesting_level: int | None) -> str:
    """Get the appropriate indentation for a list item based on nesting level."""
    if nesting_level is None:
        nesting_level = 0

    # Use 2 spaces per nesting level for markdown compatibility
    return "    " * nesting_level


def _format_bullet_marker_with_nesting(glyph: str) -> str:
    """Format the bullet marker based on the glyph and nesting level.

    According to the user's requirement, nested lists should be consistently
    either all ordered or all unordered throughout the nesting hierarchy.
    """
    if not glyph:
        return "* "

    if any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph):
        # For numbered lists, convert all levels to numbered format
        # Since markdown doesn't support nested numbering well, we'll use "1. " format for all levels
        # and rely on indentation to show the hierarchy
        return normalize_numbered_glyph(glyph)
    else:
        # This is a bullet list - use bullets for all levels
        return "* "


def _format_bullet_marker(glyph: str) -> str:
    """Format the bullet marker based on the glyph from the API."""
    if not glyph:
        return "* "

    # Check if this looks like a numbered list
    if any(char.isdigit() for char in glyph) or any(char.isalpha() for char in glyph):
        # This is a numbered list - use the glyph as-is if it ends with period
        if glyph.endswith("."):
            return f"{glyph} "
        else:
            return f"{glyph}. "
    else:
        # This is a bullet list
        return "* "
