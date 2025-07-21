import copy
from typing import Optional, Any, List, Union

import marko
from marko.ext.gfm import gfm
from marko.ext.gfm.elements import Strikethrough
from marko.inline import RawText
from pydantic import BaseModel, field_validator


from gslides_api import TextElement
from gslides_api.domain import BulletGlyphPreset
from gslides_api.text import TextRun, TextStyle
from gslides_api.request.domain import Range, RangeType
from gslides_api.request.request import CreateParagraphBulletsRequest


class ItemList(BaseModel):
    children: List[TextElement]

    @field_validator("children", mode="before")
    @classmethod
    def flatten_children(cls, v: List[Union[TextElement, "ItemList"]]) -> List[TextElement]:
        """Flatten nested ItemLists by replacing them with their children."""
        flattened = []
        for item in v:
            if isinstance(item, ItemList):
                flattened.extend(item.children)
            else:
                flattened.append(item)
        return flattened

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
    base_style = base_style or TextStyle()
    line_break = TextElement(
        endIndex=0,
        textRun=TextRun(content="\n", style=base_style),
    )

    if isinstance(markdown_ast, (marko.inline.RawText, marko.inline.LineBreak)):
        out = [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=base_style),
            )
        ]
    elif isinstance(markdown_ast, marko.block.BlankLine):
        out = [line_break]

    elif isinstance(markdown_ast, marko.inline.CodeSpan):
        base_style = copy.deepcopy(base_style)
        base_style.fontFamily = "Courier New"
        base_style.weightedFontFamily = None
        base_style.foregroundColor = {
            "opaqueColor": {"rgbColor": {"red": 0.8, "green": 0.2, "blue": 0.2}}
        }
        out = [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=base_style),
            )
        ]
    elif isinstance(markdown_ast, marko.inline.Emphasis):
        base_style = copy.deepcopy(base_style)
        base_style.italic = not base_style.italic
        out = markdown_ast_to_text_elements(
            markdown_ast.children[0], base_style, heading_style, list_depth=list_depth
        )

    elif isinstance(markdown_ast, marko.inline.StrongEmphasis):
        base_style = copy.deepcopy(base_style)
        base_style.bold = True
        out = markdown_ast_to_text_elements(
            markdown_ast.children[0], base_style, heading_style, list_depth=list_depth
        )

    elif isinstance(markdown_ast, marko.inline.Link):
        # Handle hyperlinks by setting the link property in the style
        base_style = copy.deepcopy(base_style)
        base_style.link = {"url": markdown_ast.dest}
        # Process the link text (children)
        out = sum(
            [
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth
                )
                for child in markdown_ast.children
            ],
            [],
        )

    elif isinstance(markdown_ast, Strikethrough):
        # Handle strikethrough text
        base_style = copy.deepcopy(base_style)
        base_style.strikethrough = True
        out = sum(
            [
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth
                )
                for child in markdown_ast.children
            ],
            [],
        )

    elif isinstance(markdown_ast, marko.block.Paragraph):
        out = sum(
            [
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth
                )
                for child in markdown_ast.children
            ],
            [],
        ) + [line_break]
    elif isinstance(markdown_ast, marko.block.Heading):
        # Only pass heading style to children
        out = sum(
            [
                markdown_ast_to_text_elements(
                    child, heading_style, heading_style, list_depth=list_depth
                )
                for child in markdown_ast.children
            ],
            [],
        ) + [line_break]

    elif isinstance(markdown_ast, marko.block.List):
        # Handle lists - need to pass down whether this is ordered or not
        pre_out = sum(
            [
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth + 1
                )
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
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth
                )
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
            TextElement(endIndex=0, textRun=TextRun(content="\t" * list_depth, style=base_style))
        ] + sum(
            [
                markdown_ast_to_text_elements(
                    child, base_style, heading_style, list_depth=list_depth
                )
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
