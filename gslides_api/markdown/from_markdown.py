import copy
from typing import Optional, Any, List, Tuple, Union

import marko
from marko.ext.gfm import gfm
from marko.ext.gfm.elements import Strikethrough
from marko.inline import RawText
from pydantic import BaseModel, field_validator


from gslides_api.request.request import InsertTextRequest, UpdateTextStyleRequest
from gslides_api.text import TextElement
from gslides_api.domain import BulletGlyphPreset
from gslides_api.text import TextRun, TextStyle
from gslides_api.request.domain import Range, RangeType
from gslides_api.request.request import CreateParagraphBulletsRequest, GSlidesAPIRequest


class LineBreakAfterParagraph(TextElement):
    pass


class ListItemTab(TextElement):
    pass


class LineBreakInsideList(TextElement):
    previous_element: Optional[TextElement | UpdateTextStyleRequest] = None


class ItemList(BaseModel):
    children: List[TextElement]
    style: Optional[TextStyle] = None

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

    @property
    def total_items(self):
        return sum([isinstance(c, ListItemTab) for c in self.children])


class BulletPointGroup(ItemList):
    pass


class NumberedListGroup(ItemList):
    pass


class UpdateWholeListStyleRequest(UpdateTextStyleRequest):
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
) -> list[GSlidesAPIRequest]:

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

    # Newlines inside lists will have to be inserted later, we store in them a reference to the previous element
    prev_elem = None
    for e in elements:
        if isinstance(e, LineBreakInsideList):
            e.previous_element = prev_elem
        prev_elem = e

    # Put the newlines inside lists aside, we'll insert them after creating the bullets
    newlines_inside_lists = [e for e in elements if isinstance(e, LineBreakInsideList)]
    elements = [e for e in elements if not isinstance(e, LineBreakInsideList)]

    # Assign indices to remaining text elements
    for element in elements:
        element.startIndex = start_index
        element.endIndex = start_index + len(element.textRun.content)
        start_index = element.endIndex

    # If the final element is a line break after paragraph/heading, remove it
    if elements and isinstance(elements[-1], LineBreakAfterParagraph):
        last_break = elements.pop()
        if list_items and list_items[-1].children[-1] == last_break:
            list_items[-1].children.pop()

    # Now convert the elements to requests, and store the reference to the style request inside the relevant newline
    requests, re_newlines_inside_lists = text_elements_to_requests(
        elements, newlines_inside_lists, objectId=""
    )

    # Sort bullets by start index, in reverse order so trimming the tabs doesn't mess others' indices
    list_items.sort(key=lambda b: b.start_index, reverse=True)
    for item in list_items:
        bullet_request = CreateParagraphBulletsRequest(
            objectId="",
            textRange=Range(
                type=RangeType.FIXED_RANGE,
                startIndex=item.start_index,
                endIndex=item.end_index,
            ),
            bulletPreset=(
                bullet_glyph_preset if isinstance(item, BulletPointGroup) else numbered_glyph_preset
            ),
        )
        requests.append(bullet_request)
        if item.style is not None:
            # This is needed to fix the color of bullet points that the previous request creates,
            # Which otherwise will be a random mixture of black and the color of the text
            requests.append(
                UpdateWholeListStyleRequest(
                    objectId="", style=item.style, textRange=bullet_request.textRange, fields="*"
                )
            )

    tab_end_indices = [e.endIndex for e in elements if isinstance(e, ListItemTab)]
    requests = adjust_text_style_indices_for_tab_removal(requests, tab_end_indices)

    # now that we have created the correct bullet points, and adjusted the text indices for that,
    # we can put the newlines inside lists back

    return requests


def adjust_text_style_indices_for_tab_removal(
    requests: List[GSlidesAPIRequest], tab_end_indices: list[int]
) -> List[GSlidesAPIRequest]:
    other_requests = [r for r in requests if not isinstance(r, UpdateTextStyleRequest)]
    style_requests = [
        r
        for r in requests
        if isinstance(r, UpdateTextStyleRequest) and not isinstance(r, UpdateWholeListStyleRequest)
    ]
    list_style_requests = [
        UpdateTextStyleRequest.model_validate(r.model_dump())
        for r in requests
        if isinstance(r, UpdateWholeListStyleRequest)
    ]
    # First apply blanket style for each list to cover bullet points, then custom styles for contents
    all_style_requests = list_style_requests + style_requests

    for s in all_style_requests:
        s.textRange.startIndex -= sum([t <= s.textRange.startIndex for t in tab_end_indices])
        s.textRange.endIndex -= sum([t <= s.textRange.endIndex for t in tab_end_indices])
        if s.textRange.startIndex < s.textRange.endIndex:
            other_requests.append(s)

    return other_requests


def markdown_ast_to_text_elements(
    markdown_ast: Any,
    base_style: Optional[TextStyle] = None,
    heading_style: Optional[TextStyle] = None,
    list_depth: int = 0,
) -> list[TextElement | BulletPointGroup | NumberedListGroup]:
    base_style = base_style or TextStyle()
    if heading_style is None:
        heading_style = copy.deepcopy(base_style)
        heading_style.bold = True

    line_break = TextElement(
        endIndex=0,
        textRun=TextRun(content="\n", style=base_style),
    )

    line_break_inside_list = LineBreakInsideList(
        endIndex=0, textRun=TextRun(content="\n", style=base_style)
    )

    line_break_after_paragraph = LineBreakAfterParagraph(
        endIndex=0,
        textRun=TextRun(content="\n", style=base_style),
    )

    if isinstance(markdown_ast, marko.inline.RawText):
        out = [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=base_style),
            )
        ]
    elif isinstance(markdown_ast, (marko.block.BlankLine, marko.inline.LineBreak)):
        if list_depth == 0:
            out = [line_break]
        else:
            raise ValueError("Google Slides API doesn't support newlines inside list items")
            # out = [line_break_inside_list]

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
        ) + [line_break_after_paragraph]
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
        ) + [line_break_after_paragraph]

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
            children_no_line_breaks = [c for c in pre_out if not isinstance(c, LineBreakInsideList)]
            if markdown_ast.ordered:
                out = pre_out + [
                    NumberedListGroup(children=children_no_line_breaks, style=base_style)
                ]
            else:
                out = pre_out + [
                    BulletPointGroup(children=children_no_line_breaks, style=base_style)
                ]
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
            ListItemTab(endIndex=0, textRun=TextRun(content="\t", style=base_style))
            for _ in range(list_depth)
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


def matching_newline(
    newlines_inside_lists: List[LineBreakInsideList], startIndex: int
) -> LineBreakInsideList | None:
    for n in newlines_inside_lists:
        if (
            isinstance(n.previous_element, TextElement)
            and n.previous_element.startIndex == startIndex
        ):
            return n
    return None


def text_elements_to_requests(
    text_elements: List[TextElement | GSlidesAPIRequest],
    newlines_inside_lists: List[LineBreakInsideList],
    objectId: str,
) -> Tuple[List[GSlidesAPIRequest], List[LineBreakInsideList]]:
    requests = []
    newlines = []
    for te in text_elements:
        if isinstance(te, GSlidesAPIRequest):
            te.objectId = objectId
            requests.append(te)
            continue
        else:
            assert isinstance(te, TextElement), f"Expected TextElement, got {te}"
        if te.textRun is None:
            # An empty text run will have a non-None ParagraphMarker
            # Apparently there's no direct way to insert ParagraphMarkers, instead they have to be created
            # as a side effect of inserting text or by specialized calls like createParagraphBullets
            # So we just ignore them when inserting text
            continue

        # Create InsertTextRequest
        insert_request = InsertTextRequest(
            objectId=objectId,
            text=te.textRun.content,
            insertionIndex=te.startIndex,
        )

        # Create UpdateTextStyleRequest
        text_range = Range(
            type=RangeType.FIXED_RANGE,
            startIndex=te.startIndex or 0,
            endIndex=te.endIndex,
        )
        update_style_request = UpdateTextStyleRequest(
            objectId=objectId,
            textRange=text_range,
            style=te.textRun.style,
            fields="*",
        )
        newline = matching_newline(newlines_inside_lists, te.startIndex)

        # We save the reference to the previous UpdateTextStyleRequest so we can reuse the
        # index adjustment for tab removal from it
        if newline is not None:
            newline.previous_element = update_style_request
        newlines.append(newline)

        requests.extend([insert_request, update_style_request])
    return requests, newlines
