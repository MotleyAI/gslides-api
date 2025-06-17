import copy
from typing import Optional, Any, List

import marko
from marko.ext.gfm import gfm
from marko.ext.gfm.elements import Strikethrough
from marko.inline import RawText
from pydantic import BaseModel


from gslides_api import TextElement
from gslides_api.domain import BulletGlyphPreset, TextRun, TextStyle
from gslides_api.requests import CreateParagraphBulletsRequest, Range, RangeType


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
    start_index: int = 0,
    bullet_glyph_preset: Optional[BulletGlyphPreset] = BulletGlyphPreset.BULLET_DISC_CIRCLE_SQUARE,
    numbered_glyph_preset: Optional[BulletGlyphPreset] = BulletGlyphPreset.NUMBERED_DIGIT_ALPHA_ROMAN,
) -> list[TextElement | CreateParagraphBulletsRequest]:
    # Use GFM parser to support strikethrough and other GitHub Flavored Markdown features
    doc = gfm.parse(markdown_text)
    elements_and_bullets = markdown_ast_to_text_elements(doc, base_style=base_style)
    elements = [e for e in elements_and_bullets if isinstance(e, TextElement)]
    bullets = [b for b in elements_and_bullets if isinstance(b, BulletPointGroup)]
    numbered_lists = [n for n in elements_and_bullets if isinstance(n, NumberedListGroup)]

    # Assign indices to text elements
    for element in elements:
        element.startIndex = start_index
        element.endIndex = start_index + len(element.textRun.content)
        start_index = element.endIndex

    # Sort bullets by start index, in reverse order so trimming the tabs doesn't mess others' indices
    bullets.sort(key=lambda b: b.start_index, reverse=True)
    for bullet in bullets:
        elements.append(
            CreateParagraphBulletsRequest(
                objectId="",
                textRange=Range(
                    type=RangeType.FIXED_RANGE,
                    startIndex=bullet.start_index,
                    endIndex=bullet.end_index,
                ),
                bulletPreset=bullet_glyph_preset,
            )
        )

    # Sort numbered lists by start index, in reverse order
    numbered_lists.sort(key=lambda n: n.start_index, reverse=True)
    for numbered_list in numbered_lists:
        elements.append(
            CreateParagraphBulletsRequest(
                objectId="",
                textRange=Range(
                    type=RangeType.FIXED_RANGE,
                    startIndex=numbered_list.start_index,
                    endIndex=numbered_list.end_index,
                ),
                bulletPreset=numbered_glyph_preset,
            )
        )

    return elements


def markdown_ast_to_text_elements(
    markdown_ast: Any, base_style: Optional[TextStyle] = None, is_ordered: Optional[bool] = None
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
        out = markdown_ast_to_text_elements(markdown_ast.children[0], style)

    elif isinstance(markdown_ast, marko.inline.StrongEmphasis):
        style = copy.deepcopy(style)
        style.bold = True
        out = markdown_ast_to_text_elements(markdown_ast.children[0], style)

    elif isinstance(markdown_ast, marko.inline.Link):
        # Handle hyperlinks by setting the link property in the style
        style = copy.deepcopy(style)
        style.link = {"url": markdown_ast.dest}
        # Process the link text (children)
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )

    elif isinstance(markdown_ast, Strikethrough):
        # Handle strikethrough text
        style = copy.deepcopy(style)
        style.strikethrough = True
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )

    elif isinstance(markdown_ast, marko.block.Paragraph):
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        ) + [line_break]
    elif isinstance(markdown_ast, marko.block.Heading):
        # TODO: handle heading levels properly, with font size bumps
        style = copy.deepcopy(style)
        style.bold = True
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        ) + [line_break]

    elif isinstance(markdown_ast, marko.block.List):
        # Handle lists - need to pass down whether this is ordered or not
        out = sum(
            [markdown_ast_to_text_elements(child, style, markdown_ast.ordered) for child in markdown_ast.children], []
        )
    elif isinstance(markdown_ast, marko.block.Document):
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )
    elif isinstance(markdown_ast, marko.block.ListItem):
        # https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations/request#createparagraphbulletsrequest
        # The bullet creation API is really messed up, forcing us to insert tabs that will be
        # discarded as soon as the bullets are created. So we deal with it as best we can
        pre_out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )

        # Create the appropriate group type based on whether this is an ordered list
        if is_ordered:
            list_group = NumberedListGroup(children=pre_out)
        else:
            list_group = BulletPointGroup(children=pre_out)

        out = (
            [TextElement(endIndex=0, textRun=TextRun(content="\t", style=style))]
            + pre_out
            + [list_group]
        )

    else:
        raise NotImplementedError(f"Unsupported markdown element: {markdown_ast}")

    for element in out:
        assert isinstance(
            element, (TextElement, BulletPointGroup, NumberedListGroup)
        ), f"Expected TextElement, BulletPointGroup, or NumberedListGroup, got {type(element)}"
    return out
