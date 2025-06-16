import copy
from typing import Optional, Any

import marko
from marko.inline import RawText

from gslides_api import TextElement
from gslides_api.domain import TextRun, TextStyle


def markdown_to_text_elements(markdown_text: str) -> list[TextElement]:
    doc = marko.parse(markdown_text)
    elements = markdown_ast_to_text_elements(doc)
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

    if isinstance(markdown_ast, (marko.inline.RawText, marko.inline.LineBreak)):
        return [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=style),
            )
        ]
    elif isinstance(markdown_ast, marko.block.BlankLine):
        return [
            TextElement(
                endIndex=0,
                textRun=TextRun(content="\n", style=style),
            )
        ]
    elif isinstance(markdown_ast, marko.inline.CodeSpan):
        # TODO: handle code spans properly
        return [
            TextElement(
                endIndex=0,
                textRun=TextRun(content=markdown_ast.children, style=style),
            )
        ]
    elif isinstance(markdown_ast, marko.inline.Emphasis):
        style = copy.deepcopy(style)
        style.italic = not style.italic
        return markdown_ast_to_text_elements(markdown_ast.children[0], style)

    elif isinstance(markdown_ast, marko.inline.StrongEmphasis):
        style = copy.deepcopy(style)
        style.bold = not style.bold
        return markdown_ast_to_text_elements(markdown_ast.children[0], style)

    elif isinstance(
        markdown_ast, (marko.block.Paragraph, marko.block.Heading, marko.block.Document)
    ):
        return sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )
    elif isinstance(markdown_ast, (marko.block.List, marko.block.ListItem)):
        # TODO: handle list items properly, with bullets/numbers
        return sum(
            [markdown_ast_to_text_elements(child, style) for child in markdown_ast.children], []
        )
    else:
        raise NotImplementedError(f"Unsupported markdown element: {markdown_ast}")
