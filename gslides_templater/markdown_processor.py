"""
Debug version of markdown processor to see what's happening
"""

import html
import re
import logging
from typing import List, Dict, Any, Optional, Tuple, Set
from pydantic import BaseModel, Field

from gslides_api.domain import TextStyle

logger = logging.getLogger(__name__)

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

try:
    import markdown

    MARKDOWN_AVAILABLE = True
except ImportError:
    MARKDOWN_AVAILABLE = False


class MarkdownConfig(BaseModel):
    """Configuration for Markdown processing"""

    header_sizes: Dict[int, int] = Field(
        default_factory=lambda: {1: 36, 2: 28, 3: 24, 4: 20, 5: 18, 6: 16}
    )
    max_text_length: int = Field(default=1000000, ge=1, le=10000000)
    allowed_html_tags: Set[str] = Field(default_factory=lambda: {"size", "font", "color"})
    regex_timeout: float = Field(default=1.0, ge=0.1, le=10.0)


class MarkdownProcessingError(Exception):
    """Base exception for Markdown processing"""

    pass


class UnsafeContentError(MarkdownProcessingError):
    """Content contains unsafe elements"""

    pass


class MarkdownProcessor:
    """
    Debug version of Markdown processor
    """

    def __init__(self, config: MarkdownConfig = None):
        """Initialize processor with configuration."""
        self.config = config or MarkdownConfig()
        self.header_sizes = self.config.header_sizes
        logger.info("MarkdownProcessor initialized")

    def create_slides_requests(
        self, element_id: str, markdown_text: str, base_style: Optional[TextStyle] = None
    ) -> List[Dict[str, Any]]:
        """
        Create requests for Google Slides API from Markdown text.
        DEBUG VERSION with extensive logging
        """
        logger.info(f"=== Processing markdown for element {element_id} ===")
        logger.info(f"Input text: {repr(markdown_text)}")

        requests = []

        if not markdown_text.strip():
            logger.info("Empty text, returning no requests")
            return requests

        try:
            # Step 1: Parse markdown manually
            segments = self._parse_markdown_simple(markdown_text)
            logger.info(f"Parsed {len(segments)} segments:")
            for i, seg in enumerate(segments):
                logger.info(f"  Segment {i}: content={repr(seg['content'])}, style={seg['style']}")

            if not segments:
                logger.warning("No segments found, using plain text")
                return [
                    {
                        "insertText": {
                            "objectId": element_id,
                            "insertionIndex": 0,
                            "text": markdown_text,
                        }
                    }
                ]

            # Step 2: Build complete text content
            full_text = "".join(segment["content"] for segment in segments)
            logger.info(f"Full text to insert: {repr(full_text)}")

            # Step 3: Insert all text at once
            requests.append(
                {"insertText": {"objectId": element_id, "insertionIndex": 0, "text": full_text}}
            )
            logger.info("Added text insertion request")

            # Step 4: Apply formatting to specific ranges
            current_index = 0

            for i, segment in enumerate(segments):
                content = segment["content"]
                style = segment["style"]

                logger.info(
                    f"Processing segment {i}: pos={current_index}, len={len(content)}, style={bool(style)}"
                )

                if style and content.strip():  # Only apply formatting to non-empty, styled content
                    # Calculate the field names for the style
                    style_fields = []
                    for key in style.keys():
                        style_fields.append(key)

                    if style_fields:
                        format_request = {
                            "updateTextStyle": {
                                "objectId": element_id,
                                "style": style,
                                "fields": ",".join(style_fields),
                                "textRange": {
                                    "type": "FIXED_RANGE",
                                    "startIndex": current_index,
                                    "endIndex": current_index + len(content),
                                },
                            }
                        }
                        requests.append(format_request)
                        logger.info(
                            f"Added formatting request: fields={style_fields}, range={current_index}-{current_index + len(content)}"
                        )

                current_index += len(content)

            logger.info(f"Total requests generated: {len(requests)}")
            for i, req in enumerate(requests):
                logger.info(f"Request {i}: {list(req.keys())}")

        except Exception as e:
            logger.error(f"Error processing markdown: {e}", exc_info=True)
            # Fallback to plain text insertion
            requests = [
                {"insertText": {"objectId": element_id, "insertionIndex": 0, "text": markdown_text}}
            ]
            logger.info("Using fallback plain text insertion")

        return requests

    def _parse_markdown_simple(self, text: str) -> List[Dict[str, Any]]:
        """
        Simple markdown parser that actually removes markdown symbols
        """
        segments = []
        lines = text.split("\n")

        for line_idx, line in enumerate(lines):
            logger.info(f"Processing line {line_idx}: {repr(line)}")

            if not line.strip():
                # Add newline for empty lines (except the last one)
                if line_idx < len(lines) - 1:
                    segments.append({"content": "\n", "style": {}})
                continue

            # Check if line is a header
            header_match = re.match(r"^(#{1,6})\s+(.+)", line.strip())
            if header_match:
                level = len(header_match.group(1))
                # Remove markdown symbols from header content
                content = self._clean_inline_formatting(header_match.group(2))
                logger.info(f"Found header level {level}: {repr(content)}")

                segments.append(
                    {
                        "content": content,
                        "style": {
                            "fontSize": {
                                "magnitude": self.header_sizes.get(level, 16),
                                "unit": "PT",
                            },
                            "bold": True,
                        },
                    }
                )
                # Add newline after header
                if line_idx < len(lines) - 1:
                    segments.append({"content": "\n", "style": {}})
                continue

            # Handle quote lines
            if line.strip().startswith(">"):
                content = re.sub(r"^\s*>\s*", "", line)
                content = self._clean_inline_formatting(content)
                logger.info(f"Found quote: {repr(content)}")
                segments.append(
                    {
                        "content": content,
                        "style": {
                            "italic": True,
                            "foregroundColor": {
                                "opaqueColor": {"rgbColor": {"red": 0.5, "green": 0.5, "blue": 0.5}}
                            },
                        },
                    }
                )
                if line_idx < len(lines) - 1:
                    segments.append({"content": "\n", "style": {}})
                continue

            # Parse inline formatting for regular lines
            line_segments = self._parse_line_formatting_clean(line)
            segments.extend(line_segments)

            # Add newline after line (except the last one)
            if line_idx < len(lines) - 1:
                segments.append({"content": "\n", "style": {}})

        return segments

    def _clean_inline_formatting(self, text: str) -> str:
        """
        Remove markdown formatting symbols and return clean text
        """
        # Remove bold and italic
        text = re.sub(r"\*\*\*([^*]+)\*\*\*", r"\1", text)  # Bold + Italic
        text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)  # Bold
        text = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"\1", text)  # Italic
        text = re.sub(r"~~([^~]+)~~", r"\1", text)  # Strikethrough
        text = re.sub(r"`([^`]+)`", r"\1", text)  # Code

        return text

    def _parse_line_formatting_clean(self, line: str) -> List[Dict[str, Any]]:
        """
        Parse line and return segments with formatting, removing markdown symbols
        """
        if not line.strip():
            return [{"content": line, "style": {}}]

        segments = []
        current_pos = 0

        # Patterns for inline formatting - ordered by precedence
        patterns = [
            (r"\*\*\*([^*]+)\*\*\*", {"bold": True, "italic": True}),  # Bold + Italic
            (r"\*\*([^*]+)\*\*", {"bold": True}),  # Bold
            (r"(?<!\*)\*([^*]+)\*(?!\*)", {"italic": True}),  # Italic
            (r"~~([^~]+)~~", {"strikethrough": True}),  # Strikethrough
            (
                r"`([^`]+)`",
                {  # Code
                    "fontFamily": "Courier New",
                    "foregroundColor": {
                        "opaqueColor": {"rgbColor": {"red": 0.8, "green": 0.2, "blue": 0.2}}
                    },
                },
            ),
        ]

        while current_pos < len(line):
            # Find the next formatting match
            next_match = None
            next_pos = len(line)
            next_style = {}

            for pattern, style in patterns:
                match = re.search(pattern, line[current_pos:])
                if match:
                    match_start = current_pos + match.start()
                    if match_start < next_pos:
                        next_match = match
                        next_pos = match_start
                        next_style = style
                        break  # Take the first match found

            if next_match and next_pos < len(line):
                # Add plain text before the match
                if next_pos > current_pos:
                    plain_text = line[current_pos:next_pos]
                    segments.append({"content": plain_text, "style": {}})
                    logger.info(f"Added plain text: {repr(plain_text)}")

                # Add formatted text (WITHOUT markdown symbols)
                formatted_content = next_match.group(1)  # Content inside markdown symbols
                segments.append({"content": formatted_content, "style": next_style})
                logger.info(
                    f"Added formatted text: {repr(formatted_content)} with style {next_style}"
                )

                current_pos = next_pos + len(next_match.group(0))
            else:
                # Add remaining text
                remaining = line[current_pos:]
                if remaining:
                    segments.append({"content": remaining, "style": {}})
                    logger.info(f"Added remaining text: {repr(remaining)}")
                break

        return segments

    # Stub methods for compatibility
    def parse_markdown_to_segments(self, text: str) -> List[Dict[str, Any]]:
        return self._parse_markdown_simple(text)

    def is_markdown(self, text: str) -> bool:
        return any(char in text for char in ["*", "#", "`", "~", ">"])

    def markdown_to_slides_elements(self, markdown_text: str) -> List[Dict[str, Any]]:
        segments = self._parse_markdown_simple(markdown_text)
        return [{"textRun": {"content": seg["content"], "style": seg["style"]}} for seg in segments]

    def slides_elements_to_markdown(self, elements: List[Dict[str, Any]]) -> str:
        return "".join(elem.get("textRun", {}).get("content", "") for elem in elements)

    def parse_to_components(self, markdown_content: str) -> List[Dict[str, Any]]:
        return [{"type": "text", "content": markdown_content}]

    def clean_text_for_slides(self, text: str) -> str:
        return self._clean_inline_formatting(text)

    def validate_markdown(self, markdown_text: str) -> List[str]:
        return []

    def extract_images(self, markdown_text: str) -> List[Dict[str, str]]:
        return []

    def extract_links(self, markdown_text: str) -> List[Dict[str, str]]:
        return []


# Convenience functions
def markdown_to_slides_elements(markdown_text: str) -> List[Dict[str, Any]]:
    processor = MarkdownProcessor()
    return processor.markdown_to_slides_elements(markdown_text)


def slides_elements_to_markdown(elements: List[Dict[str, Any]]) -> str:
    processor = MarkdownProcessor()
    return processor.slides_elements_to_markdown(elements)


def clean_markdown_for_slides(text: str) -> str:
    processor = MarkdownProcessor()
    return processor.clean_text_for_slides(text)
