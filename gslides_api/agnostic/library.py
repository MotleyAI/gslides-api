from pydantic import BaseModel
import json

from gslides_api.agnostic.presentation import MarkdownDeck, MarkdownSlide
from gslides_api.agnostic.element import (
    ContentType,
    MarkdownSlideElement,
    MarkdownTextElement,
    MarkdownChartElement,
    MarkdownTableElement,
    MarkdownImageElement,
    MarkdownContentElement,
)

example_slides = [
    MarkdownSlide(
        name="Title",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownTextElement.placeholder(name="Subtitle"),
        ],
    ),
    MarkdownSlide(
        name="Section header",
        elements=[
            MarkdownTextElement.placeholder(name="Section header"),
        ],
    ),
    MarkdownSlide(
        name="Header and table",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownTableElement.placeholder(name="Table"),
        ],
    ),
    MarkdownSlide(
        name="Header and single content",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownContentElement.placeholder(name="Content"),
        ],
    ),
    MarkdownSlide(
        name="2 content slide",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownContentElement.placeholder(name="Content 1"),
            MarkdownContentElement.placeholder(name="Content 2"),
        ],
    ),
    MarkdownSlide(
        name="Chart and text slide",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownChartElement.placeholder(name="Chart"),
            MarkdownTextElement.placeholder(name="Text"),
        ],
    ),
    MarkdownSlide(
        name="Comparison",
        elements=[
            MarkdownTextElement.placeholder(name="Title"),
            MarkdownTextElement.placeholder(name="Subtitle 1"),
            MarkdownContentElement.placeholder(name="Content 1"),
            MarkdownTextElement.placeholder(name="Subtitle 2"),
            MarkdownContentElement.placeholder(name="Content 2"),
        ],
    ),
]


def _element_names_match(library_slide: MarkdownSlide, parsed_slide: MarkdownSlide) -> bool:
    """Check if parsed element names are a subset of library template names."""
    library_names = {el.name for el in library_slide.elements}
    parsed_names = {el.name for el in parsed_slide.elements}
    return parsed_names.issubset(library_names)


def _element_types_match(library_slide: MarkdownSlide, parsed_slide: MarkdownSlide) -> bool:
    """Check if parsed element types match library types for elements that exist.

    ContentType.ANY in library matches any specific content type.
    """
    library_by_name = {el.name: el for el in library_slide.elements}
    for parsed_el in parsed_slide.elements:
        if parsed_el.name not in library_by_name:
            return False  # Name not in library
        lib_el = library_by_name[parsed_el.name]
        # ContentType.ANY in library matches any content type
        if lib_el.content_type == ContentType.ANY:
            continue
        # Otherwise types must match exactly
        if lib_el.content_type != parsed_el.content_type:
            return False
    return True


def _create_empty_element(template_element: MarkdownSlideElement) -> MarkdownSlideElement:
    """Create an empty element matching the template's type.

    Empty elements have content=None.
    """
    if template_element.content_type == ContentType.TEXT:
        return MarkdownTextElement(name=template_element.name, content=None)
    elif template_element.content_type == ContentType.CHART:
        return MarkdownChartElement(name=template_element.name, content=None)
    elif template_element.content_type == ContentType.TABLE:
        return MarkdownTableElement(name=template_element.name, content=None)
    elif template_element.content_type == ContentType.IMAGE:
        return MarkdownImageElement(name=template_element.name, content=None)
    elif template_element.content_type == ContentType.ANY:
        return MarkdownContentElement(name=template_element.name, content=None)
    else:
        raise ValueError(f"Unknown content type: {template_element.content_type}")


def _merge_elements_with_template(
    library_slide: MarkdownSlide, parsed_slide: MarkdownSlide
) -> list[MarkdownSlideElement]:
    """Merge parsed elements into library template, filling missing with empty elements.

    Elements are returned in the order defined by the library template.
    Missing elements are filled with empty content of the appropriate type.
    """
    parsed_by_name = {el.name: el for el in parsed_slide.elements}
    result = []
    for lib_el in library_slide.elements:
        if lib_el.name in parsed_by_name:
            result.append(parsed_by_name[lib_el.name])
        else:
            result.append(_create_empty_element(lib_el))
    return result


class SlideLayoutLibrary(BaseModel):
    slides: list[MarkdownSlide]

    def __getitem__(self, key: str) -> MarkdownSlide:
        for slide in self.slides:
            if slide.name == key:
                return slide
        raise KeyError(f"Key {key} not found in slide library")

    def __setitem__(self, key: str, value: MarkdownSlide) -> None:
        for i, slide in enumerate(self.slides):
            if slide.name == key:
                self.slides[i] = value
                return
        self.slides.append(value)

    def values(self) -> list[MarkdownSlide]:
        return self.slides

    def keys(self) -> list[str]:
        return [slide.name for slide in self.slides]

    def items(self) -> list[tuple[str, MarkdownSlide]]:
        return [(slide.name, slide) for slide in self.slides]

    def instructions(self) -> str:
        desc = json.dumps([slide.to_markdown() for slide in self.slides])

        instructions = """Here is the list of available slides, separated by *****.
        A valid deck is a list of strings, EACH STRING DESCRIBING THE CONTENT OF ONE WHOLE SLIDE, 
        formatted as in the examples below.
        IMPORTANT: if the above examples contain "any" types, this means that element can
        be any of the other valid types: text, chart, table, image.
        Make sure that the content of the text elements does NOT contain any images or tables.
        If you don't want to populate an element in the slide you're using, pass an empty string,
        but you MUST match one of the layouts provided.
        YOU ARE NOT ALLOWED TO RETURN ELEMENTS WITH A LITERAL "any" TYPE.
        Here are the valid slide layouts:
        """

        return instructions + desc

    def slide_from_markdown(self, markdown: str, name: str | None = None) -> MarkdownSlide:
        """
        Parses a slide using MarkdownSlide.from_markdown() and tries to match it to a slide in the library.
        Matching is done by name first, then by element types.

        When matching by name:
        - Parsed element names must be a subset of library element names
        - Parsed element types must match library types for provided elements
        - Missing elements are filled with empty content of the appropriate type

        If no name match, falls back to matching by element types only.
        If no match is found, an exception is raised.

        :param markdown: Markdown string representing a slide
        :return: MarkdownSlide with elements merged from library template
        """
        # 1. Parse the markdown
        parsed_slide = MarkdownSlide.from_markdown(markdown)

        # 2. Try to match by slide name first
        if parsed_slide.name:
            for library_slide in self.slides:
                if library_slide.name == parsed_slide.name:
                    # Verify element names are a subset of library names
                    if not _element_names_match(library_slide, parsed_slide):
                        parsed_names = [el.name for el in parsed_slide.elements]
                        library_names = [el.name for el in library_slide.elements]
                        raise ValueError(
                            f"Element names don't match for slide '{parsed_slide.name}'. "
                            f"Expected subset of {library_names}, got {parsed_names}"
                        )
                    # Verify element types match for provided elements
                    if not _element_types_match(library_slide, parsed_slide):
                        parsed_types = [
                            (el.name, el.content_type.value) for el in parsed_slide.elements
                        ]
                        library_types = [
                            (el.name, el.content_type.value) for el in library_slide.elements
                        ]
                        raise ValueError(
                            f"Element types don't match for slide '{parsed_slide.name}'. "
                            f"Library types: {library_types}, got {parsed_types}"
                        )
                    # Merge elements: fill missing with empty content
                    parsed_slide.elements = _merge_elements_with_template(
                        library_slide, parsed_slide
                    )
                    if name:
                        parsed_slide.name = name

                    return parsed_slide

        # 3. Try to match by element types (fallback)
        for library_slide in self.slides:
            if _element_types_match(library_slide, parsed_slide):
                # Update parsed slide name if matched
                parsed_slide.name = library_slide.name
                # Merge elements: fill missing with empty content
                parsed_slide.elements = _merge_elements_with_template(library_slide, parsed_slide)
                if name:
                    parsed_slide.name = name

                return parsed_slide

        # 4. No match found
        parsed_names = [el.name for el in parsed_slide.elements]
        raise ValueError(f"No matching slide layout found. Parsed element names: {parsed_names}")


if __name__ == "__main__":
    library = SlideLayoutLibrary(slides=example_slides)

    print(example_slides[-2].to_markdown())
    print("Yay!")
