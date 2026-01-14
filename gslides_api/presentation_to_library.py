"""Convert a Google Slides Presentation to a SlideLayoutLibrary."""

from gslides_api.presentation import Presentation
from gslides_api.element.base import ElementKind
from gslides_api.agnostic.library import SlideLayoutLibrary
from gslides_api.agnostic.presentation import MarkdownSlide
from gslides_api.agnostic.element import (
    MarkdownTextElement,
    MarkdownChartElement,
    MarkdownTableElement,
    MarkdownImageElement,
    MarkdownSlideElement,
    TableData,
)


def presentation_to_library(presentation: Presentation) -> SlideLayoutLibrary:
    """
    Convert a Presentation to a SlideLayoutLibrary.

    Only includes slides with non-empty speaker notes and elements with non-empty alt-titles.

    Element type mapping:
    - Images with alt-title starting with "chart" (case-insensitive) → MarkdownChartElement
    - Other images → MarkdownImageElement
    - Tables → MarkdownTableElement
    - Shapes/text → MarkdownTextElement

    Args:
        presentation: The Google Slides Presentation to convert.

    Returns:
        SlideLayoutLibrary containing the converted slides.
    """
    markdown_slides = []

    for slide in presentation.slides or []:
        # Get speaker notes text
        notes_text = slide.speaker_notes.read_text().strip()
        if not notes_text:
            continue  # Skip slides without speaker notes

        elements = []
        for element in slide.page_elements_flat:
            # Skip elements without alt-title
            if not element.title or element.title.endswith(
                ".png"
            ):  # TODO: Hacky fix for PPTX export, remove when fixed
                continue

            md_element = _convert_element(element)
            if md_element:
                elements.append(md_element)

        if elements:  # Only add slide if it has elements
            markdown_slides.append(MarkdownSlide(name=notes_text, elements=elements))

    return SlideLayoutLibrary(slides=markdown_slides)


def _convert_element(element) -> MarkdownSlideElement | None:
    """
    Convert a page element to a MarkdownSlideElement.

    Args:
        element: A PageElement from Google Slides (shape, image, table, etc.)

    Returns:
        The appropriate MarkdownSlideElement subclass, or None for unsupported types.
    """
    name = element.title

    if element.type == ElementKind.IMAGE:
        if name.lower().startswith("chart"):
            # Chart: use placeholder content (images don't have semantic data)
            return MarkdownChartElement(name=name, content="Chart placeholder")
        else:
            # Image: use contentUrl
            url = element.image.contentUrl or "https://via.placeholder.com/400x300"
            return MarkdownImageElement(name=name, content=url, metadata={"alt_text": name})

    elif element.type == ElementKind.TABLE:
        try:
            table_data = element.extract_table_data()
        except ValueError:
            table_data = TableData(headers=["Column"], rows=[["Data"]])
        return MarkdownTableElement(name=name, content=table_data)

    elif element.type == ElementKind.SHAPE:
        text_content = element.read_text(as_markdown=True)
        return MarkdownTextElement(name=name, content=text_content)

    return None  # Skip unsupported element types
