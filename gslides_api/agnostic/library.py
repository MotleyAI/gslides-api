from pydantic import BaseModel

from gslides_api.agnostic.presentation import MarkdownDeck, MarkdownSlide
from gslides_api.agnostic.element import (
    ContentType,
    MarkdownTextElement,
    MarkdownChartElement,
    MarkdownTableElement,
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
    """Check if element names match in order between library template and parsed slide."""
    if len(library_slide.elements) != len(parsed_slide.elements):
        return False
    return all(
        lib_el.name == parsed_el.name
        for lib_el, parsed_el in zip(library_slide.elements, parsed_slide.elements)
    )


def _element_types_match(library_slide: MarkdownSlide, parsed_slide: MarkdownSlide) -> bool:
    """Check if element types match, with ContentType.ANY matching any specific type."""
    if len(library_slide.elements) != len(parsed_slide.elements):
        return False
    for lib_el, parsed_el in zip(library_slide.elements, parsed_slide.elements):
        # ContentType.ANY in library matches any content type
        if lib_el.content_type == ContentType.ANY:
            continue
        # Otherwise types must match exactly
        if lib_el.content_type != parsed_el.content_type:
            return False
    return True


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
        desc = "\n*****\n".join([slide.to_markdown() for slide in self.slides])

        instructions = """Here is the list of available slides, separated by *****.
        A valid deck is a list of strings, one per slide, formatted as in the examples below.
        IMPORTANT: if the above examples contain "any" types, this means that element can
        be any of the other valid types: text, chart, table, image.
        Make sure that the content of the text elements does NOT contain any images or tables.
        If you don't want to populate an element in the slide you're using, pass an empty string,
        but you MUST match one of the layouts provided.
        YOU ARE NOT ALLOWED TO RETURN ELEMENTS WITH A LITERAL "any" TYPE.
        Here are the valid slide layouts:
        """

        return instructions + desc

    def slide_from_markdown(self, markdown: str) -> MarkdownSlide:
        """
        Parses a slide using MarkdownSlide.from_markdown() and tries to match it to a slide in the library.
        Matching is done by name first, then by element names, then by element types.
        If slide names match, verification is done by element names and element types.
        If element names match, verification is done by element types.
        If no match is found, an exception is raised.
        :param markdown:
        :return:
        """
        # 1. Parse the markdown
        parsed_slide = MarkdownSlide.from_markdown(markdown)

        # 2. Try to match by slide name first
        if parsed_slide.name:
            for library_slide in self.slides:
                if library_slide.name == parsed_slide.name:
                    # Verify element names match first
                    if not _element_names_match(library_slide, parsed_slide):
                        parsed_names = [el.name for el in parsed_slide.elements]
                        library_names = [el.name for el in library_slide.elements]
                        raise ValueError(
                            f"Element names don't match for slide '{parsed_slide.name}'. "
                            f"Expected {library_names}, got {parsed_names}"
                        )
                    # Verify element types match
                    if not _element_types_match(library_slide, parsed_slide):
                        parsed_types = [el.content_type.value for el in parsed_slide.elements]
                        library_types = [el.content_type.value for el in library_slide.elements]
                        raise ValueError(
                            f"Element types don't match for slide '{parsed_slide.name}'. "
                            f"Expected {library_types}, got {parsed_types}"
                        )
                    return parsed_slide

        # 3. Try to match by element names
        for library_slide in self.slides:
            if _element_types_match(library_slide, parsed_slide):
                # Update parsed slide name if matched
                parsed_slide.name = library_slide.name
                return parsed_slide

        # 4. No match found
        parsed_names = [el.name for el in parsed_slide.elements]
        raise ValueError(f"No matching slide layout found. Parsed element names: {parsed_names}")


if __name__ == "__main__":
    library = SlideLayoutLibrary(slides=example_slides)

    print(example_slides[-2].to_markdown())
    print("Yay!")
