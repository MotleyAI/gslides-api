import copy
from typing import Optional, Any

import marko
from marko.inline import RawText

from gslides_api import TextElement
from gslides_api.domain import TextRun, TextStyle


def markdown_to_text_elements(
    markdown_text: str, base_style: Optional[TextStyle] = None
) -> list[TextElement]:
    doc = marko.parse(markdown_text)
    elements = markdown_ast_to_text_elements(doc, base_style=base_style)
    start_index = 0
    for element in elements:
        element.startIndex = start_index
        element.endIndex = start_index + len(element.textRun.content)
        start_index = element.endIndex
    return elements


def markdown_ast_to_text_elements(
    markdown_ast: Any, base_style: Optional[TextStyle] = None
) -> list[TextElement]:
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
        # TODO: handle code spans properly
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
        style.bold = not style.bold
        out = markdown_ast_to_text_elements(markdown_ast.children[0], style)

    elif isinstance(markdown_ast, marko.block.Paragraph):
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        ) + [line_break]
    elif isinstance(markdown_ast, marko.block.Heading):
        # TODO: handle heading levels properly, with font size bumps?
        style = copy.deepcopy(style)
        style.bold = True
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )

    elif isinstance(markdown_ast, (marko.block.List, marko.block.Document)):
        # TODO: handle list items properly, with bullets/numbers
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )
    elif isinstance(markdown_ast, marko.block.ListItem):
        out = sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )
    else:
        raise NotImplementedError(f"Unsupported markdown element: {markdown_ast}")

    for element in out:
        assert isinstance(element, TextElement), f"Expected TextElement, got {type(element)}"
    return out
