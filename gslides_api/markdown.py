import copy
from typing import Optional, Any, List

import marko
from marko.ext.gfm import gfm
from marko.ext.gfm.elements import Strikethrough
from marko.inline import RawText
from pydantic import BaseModel


from gslides_api import TextElement
from gslides_api.domain import BulletGlyphPreset
from gslides_api.text import TextRun, TextStyle
from gslides_api.request.domain import Range, RangeType
from gslides_api.request.request import CreateParagraphBulletsRequest


class ItemList(BaseModel):
    children: List[TextElement]

    @property
    def start_index(self):
        return self.children[0].startIndex

    @property
    def end_index(self):
        return self.children[-1].endIndex


class BulletPointGroup(ItemList):
    pass


class NumberedListGroup(ItemList):
    pass


def markdown_to_text_elements(
    markdown_text: str,
    base_style: Optional[TextStyle] = None,
    heading_style: Optional[TextStyle] = None,
    start_index: int = 0,
    bullet_glyph_preset: Optional[BulletGlyphPreset] = BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE,
    numbered_glyph_preset: Optional[
        BulletGlyphPreset
    ] = BulletGlyphPreset.NUMBERED_DIGIT_ALPHA_ROMAN,
) -> list[TextElement | CreateParagraphBulletsRequest]:

    heading_style = heading_style or copy.deepcopy(base_style)
    heading_style = heading_style or TextStyle()
    heading_style.bold = True
    # TODO: handle heading levels properly, with font size bumps for heading levels?

    # Use GFM parser to support strikethrough and other GitHub Flavored Markdown features
    doc = gfm.parse(markdown_text)
    elements_and_bullets = markdown_ast_to_text_elements(
        doc, base_style=base_style, heading_style=heading_style
    )
    elements = [e for e in elements_and_bullets if isinstance(e, TextElement)]
    list_items = [b for b in elements_and_bullets if isinstance(b, ItemList)]

    # Assign indices to text elements
    for element in elements:
        element.startIndex = start_index
        element.endIndex = start_index + len(element.textRun.content)
        start_index = element.endIndex

    # Sort bullets by start index, in reverse order so trimming the tabs doesn't mess others' indices
    list_items.sort(key=lambda b: b.start_index, reverse=True)
    for item in list_items:
        elements.append(
            CreateParagraphBulletsRequest(
                objectId="",
                textRange=Range(
                    type=RangeType.FIXED_RANGE,
                    startIndex=item.start_index,
                    endIndex=item.end_index,
                ),
                bulletPreset=(
                    bullet_glyph_preset
                    if isinstance(item, BulletPointGroup)
                    else numbered_glyph_preset
                ),
            )
        )

    return elements


def markdown_ast_to_text_elements(
    markdown_ast: Any,
    base_style: Optional[TextStyle] = None,
    heading_style: Optional[TextStyle] = None,
    list_depth: int = 0,
) -> list[TextElement | BulletPointGroup | NumberedListGroup]:
    style = base_style or TextStyle()
    line_break = TextElement(
        endIndex=0,
        textRun=TextRun(content="\n", style=style),
    )

    if isinstance(markdown_ast, (marko.inline.RawText, marko.inline.LineBreak)):
        out = [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=style),
            )
        ]
    elif isinstance(markdown_ast, marko.block.BlankLine):
        out = [line_break]

    elif isinstance(markdown_ast, marko.inline.CodeSpan):
        style = copy.deepcopy(style)
        style.fontFamily = "Courier New"
        style.weightedFontFamily = None
        style.foregroundColor = {
            "opaqueColor": {"rgbColor": {"red": 0.8, "green": 0.2, "blue": 0.2}}
        }
        out = [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=style),
            )
        ]
    elif isinstance(markdown_ast, marko.inline.Emphasis):
        style = copy.deepcopy(style)
        style.italic = not style.italic
        out = markdown_ast_to_text_elements(markdown_ast.children[0], style, list_depth)

    elif isinstance(markdown_ast, marko.inline.StrongEmphasis):
        style = copy.deepcopy(style)
        style.bold = True
        out = markdown_ast_to_text_elements(markdown_ast.children[0], style, list_depth)

    elif isinstance(markdown_ast, marko.inline.Link):
        # Handle hyperlinks by setting the link property in the style
        style = copy.deepcopy(style)
        style.link = {"url": markdown_ast.dest}
        # Process the link text (children)
        out = sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        )

    elif isinstance(markdown_ast, Strikethrough):
        # Handle strikethrough text
        style = copy.deepcopy(style)
        style.strikethrough = True
        out = sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        )

    elif isinstance(markdown_ast, marko.block.Paragraph):
        out = sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        ) + [line_break]
    elif isinstance(markdown_ast, marko.block.Heading):

        out = sum(
            [
                markdown_ast_to_text_elements(child, heading_style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        ) + [line_break]

    elif isinstance(markdown_ast, marko.block.List):
        # Handle lists - need to pass down whether this is ordered or not
        pre_out = sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth + 1)
                for child in markdown_ast.children
            ],
            [],
        )
        # Create the appropriate group type based on whether this is an ordered list
        if list_depth == 0:
            if markdown_ast.ordered:
                out = pre_out + [NumberedListGroup(children=pre_out)]
            else:
                out = pre_out + [BulletPointGroup(children=pre_out)]
        else:
            out = pre_out
    elif isinstance(markdown_ast, marko.block.Document):
        out = sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        )
    elif isinstance(markdown_ast, marko.block.ListItem):
        # https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createparagraphbulletsrequest
        # The bullet creation API is really messed up, forcing us to insert tabs that will be
        # discarded as soon as the bullets are created. So we deal with it as best we can
        # TODO: handle nested lists
        out = [
            TextElement(endIndex=0, textRun=TextRun(content="\t", style=style))
            for _ in range(list_depth)
        ] + sum(
            [
                markdown_ast_to_text_elements(child, style, list_depth)
                for child in markdown_ast.children
            ],
            [],
        )

    else:
        raise NotImplementedError(f"Unsupported markdown element: {markdown_ast}")

    for element in out:
        assert isinstance(
            element, (TextElement, BulletPointGroup, NumberedListGroup)
        ), f"Expected TextElement, BulletPointGroup, or NumberedListGroup, got {type(element)}"
    return out


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
                if paragraph_text:
                    result.append(paragraph_text)
                current_paragraph = []

    # Add any remaining paragraph content
    if current_paragraph:
        paragraph_text = "".join(current_paragraph).rstrip()
        if paragraph_text:
            result.append(paragraph_text)

    return "\n".join(result) if result else None


def _apply_markdown_formatting(content: str, style) -> str:
    """Apply markdown formatting to content based on text style."""
    if style is None:
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
    return "  " * nesting_level


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
