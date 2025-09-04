import logging
import re
from typing import Literal, Union

from pydantic import BaseModel, Field

from gslides_api.markdown.element import (
    MarkdownChartElement,
    ContentType,
    MarkdownImageElement,
    MarkdownTableElement,
    MarkdownTextElement,
)

logger = logging.getLogger(__name__)

# Union type for all element types
MarkdownSlideElementUnion = Union[
    MarkdownTextElement, MarkdownImageElement, MarkdownTableElement, MarkdownChartElement
]


class MarkdownSlide(BaseModel):
    elements: list[MarkdownSlideElementUnion] = Field(default_factory=list)
    name: str | None = None

    def to_markdown(self) -> str:
        """Convert slide back to markdown format."""
        lines = []

        # Add slide name comment if present
        if self.name:
            lines.append(f"<!-- slide: {self.name} -->")

        # Add element content
        element_content = "\n\n".join(element.to_markdown() for element in self.elements)
        if element_content:
            lines.append(element_content)

        return "\n".join(lines)

    @classmethod
    def _create_element(
        cls, name: str, content: str, content_type: ContentType
    ) -> MarkdownSlideElementUnion:
        """Create the appropriate element type based on content_type."""
        if content_type == ContentType.TEXT:
            return MarkdownTextElement(name=name, content=content)
        elif content_type == ContentType.IMAGE:
            # For images, use from_markdown to properly parse URL and metadata
            return MarkdownImageElement.from_markdown(name=name, markdown_content=content)
        elif content_type == ContentType.TABLE:
            # TableElement will validate and parse the content in its validator
            return MarkdownTableElement(name=name, content=content)
        elif content_type == ContentType.CHART:
            return MarkdownChartElement(name=name, content=content)
        else:
            # Fallback to TextElement for unknown types
            return MarkdownTextElement(name=name, content=content)

    @classmethod
    def from_markdown(
        cls, slide_content: str, on_invalid_element: Literal["warn", "raise"] = "warn"
    ) -> "MarkdownSlide":
        """Parse a single slide's markdown content into elements."""
        elements = []
        slide_name = None

        # Check for slide name comment at the beginning
        slide_name_match = re.match(r"^\s*<!--\s*slide:\s*([^>]+)\s*-->\s*", slide_content)
        if slide_name_match:
            slide_name = slide_name_match.group(1).strip()
            # Remove the slide name comment from content
            slide_content = slide_content[slide_name_match.end() :]

        # Split content by HTML comments
        parts = re.split(r"(<!-- *(\w+): *([^>]+) *-->)", slide_content)

        current_content = parts[0].strip() if parts else ""

        # Handle initial content before first HTML comment (default text element)
        if current_content:
            elements.append(
                cls._create_element(
                    name="Default",
                    content=current_content,
                    content_type=ContentType.TEXT,
                )
            )

        # Process parts with HTML comments
        i = 1
        while i < len(parts):
            if i + 2 < len(parts):
                element_type = parts[i + 1].strip()  # Element type
                element_name = parts[i + 2].strip()  # Element name
                content = parts[i + 3].strip() if i + 3 < len(parts) else ""

                # Validate element type
                try:
                    content_type = ContentType(element_type)
                except ValueError:
                    if on_invalid_element == "raise":
                        raise ValueError(f"Invalid element type: {element_type}")
                    else:
                        logger.warning(f"Invalid element type '{element_type}', treating as text")
                        content_type = ContentType.TEXT

                if content:
                    try:
                        element = cls._create_element(
                            name=element_name,
                            content=content,
                            content_type=content_type,
                        )
                        elements.append(element)
                    except ValueError as e:
                        if on_invalid_element == "raise":
                            raise ValueError(
                                f"Invalid content for {content_type.value} element '{element_name}': {e}"
                            )
                        else:
                            logger.warning(
                                f"Invalid content for {content_type.value} element '{element_name}': {e}. Converting to text element."
                            )
                            # Create as text element if validation fails
                            elements.append(MarkdownTextElement(name=element_name, content=content))

                i += 4
            else:
                i += 1

        return cls(elements=elements, name=slide_name)


class MarkdownDeck(BaseModel):
    slides: list[MarkdownSlide] = Field(default_factory=list)

    def dumps(self) -> str:
        """Convert deck back to markdown format."""
        slide_contents = []

        for slide in self.slides:
            slide_md = slide.to_markdown()
            if slide_md.strip():
                slide_contents.append(slide_md)

        return "---\n" + "\n\n---\n".join(slide_contents) + "\n"

    @classmethod
    def loads(
        cls,
        markdown_content: str,
        on_invalid_element: Literal["warn", "raise"] = "warn",
    ) -> "MarkdownDeck":
        """Parse markdown content into a MarkdownDeck."""
        # Remove optional leading --- if present
        content = markdown_content.strip()
        if content.startswith("---"):
            content = content[3:].lstrip()

        # Split by slide separators
        slide_parts = content.split("\n---\n")

        slides = []
        for slide_content in slide_parts:
            slide_content = slide_content.strip()
            if slide_content:
                slide = MarkdownSlide.from_markdown(slide_content, on_invalid_element)
                if slide.elements or slide.name:  # Add slides with content or name
                    slides.append(slide)

        return cls(slides=slides)


if __name__ == "__main__":
    example_md = """
---
# Slide Title

<!-- text: Text_1 -->
## Introduction

Content here...

<!-- text: Details -->
## Details

More content...

<!-- image: Image_1 -->
![Image](https://example.com/image.jpg)

<!-- chart: Chart_1 -->
```json
{
    "data": [1, 2, 3]
}
```

<!-- table: Table_1 -->
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |

---
<!-- slide: Next Slide -->
# Next Slide

<!-- text: Summary -->
## Summary

Final thoughts
"""

    deck = MarkdownDeck.loads(example_md)
    print("=== Loaded and dumped example ===")
    print(deck.dumps())

    # Demonstrate from_df functionality
    try:
        import pandas as pd

        print("\n=== DataFrame to TableElement example ===")
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Age": [25, 30], "City": ["NYC", "SF"]})

        table_element = MarkdownTableElement.from_df(df, name="People")
        print("Generated markdown:")
        print(table_element.to_markdown())

    except ImportError:
        print("\n=== pandas not available, skipping DataFrame example ===")
