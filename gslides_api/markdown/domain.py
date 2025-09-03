import logging
import re
from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, Field


logger = logging.getLogger(__name__)


class ContentType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    CHART = "chart"
    TABLE = "table"


class MarkdownSlideElement(BaseModel):
    name: str
    content: str
    content_type: ContentType
    metadata: dict[str, Any] = Field(default_factory=dict)

    def to_markdown(self) -> str:
        """Convert element back to markdown format."""
        lines = []
        
        # Add HTML comment for element type and name (except for default text)
        if not (self.content_type == ContentType.TEXT and self.name == "Default"):
            lines.append(f"<!-- {self.content_type.value}: {self.name} -->")
        
        # Add content
        lines.append(self.content.rstrip())

        
        return "\n".join(lines)


class MarkdownSlide(BaseModel):
    elements: list[MarkdownSlideElement] = Field(default_factory=list)

    def to_markdown(self) -> str:
        """Convert slide back to markdown format."""
        return "\n\n".join(element.to_markdown() for element in self.elements)

    @classmethod
    def from_markdown(
        cls, 
        slide_content: str, 
        on_invalid_element: Literal["warn", "raise"] = "warn"
    ) -> "MarkdownSlide":
        """Parse a single slide's markdown content into elements."""
        elements = []
        
        # Split content by HTML comments
        parts = re.split(r'(<!-- *(\w+): *([^>]+) *-->)', slide_content)
        
        current_content = parts[0].strip() if parts else ""
        
        # Handle initial content before first HTML comment (default text element)
        if current_content:
            elements.append(MarkdownSlideElement(
                name="Default",
                content=current_content,
                content_type=ContentType.TEXT
            ))
        
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
                    elements.append(MarkdownSlideElement(
                        name=element_name,
                        content=content,
                        content_type=content_type
                    ))
                
                i += 4
            else:
                i += 1
        
        return cls(elements=elements)


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
        on_invalid_element: Literal["warn", "raise"] = "warn"
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
                if slide.elements:  # Only add slides with content
                    slides.append(slide)
        
        return cls(slides=slides)


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
# Next Slide

<!-- text: Summary -->
## Summary

Final thoughts
"""
