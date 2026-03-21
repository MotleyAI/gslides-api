"""
Concrete implementation of abstract slides using BeautifulSoup for HTML.
This module provides the actual implementation that maps abstract slide operations to HTML/BeautifulSoup calls.
"""

import copy
import io
import logging
import os
import shutil
import tempfile
import uuid
from typing import Annotated, Any, List, Optional, Union

from bs4 import BeautifulSoup
from bs4.element import Tag
from pydantic import BaseModel, ConfigDict, Discriminator, Field
from pydantic import Tag as PydanticTag
from pydantic import TypeAdapter, model_validator

from gslides_api.agnostic.domain import ImageData
from gslides_api.agnostic.element import MarkdownTableElement, TableData
from gslides_api.agnostic.units import OutputUnit, from_emu, to_emu

from gslides_api.common.download import download_binary_file

from gslides_api.adapters.abstract_slides import (
    AbstractAltText,
    AbstractCredentials,
    AbstractElement,
    AbstractElementKind,
    AbstractImageElement,
    AbstractPresentation,
    AbstractShapeElement,
    AbstractSlide,
    AbstractSlideProperties,
    AbstractSlidesAPIClient,
    AbstractSpeakerNotes,
    AbstractTableElement,
    AbstractThumbnail,
)

logger = logging.getLogger(__name__)

# Type alias for HTML element inputs
HTMLElementInput = Union[Tag, dict, "HTMLElementParent"]

# Inline formatting tags to EXCLUDE from element parsing
# These are text styling elements that should not be separate blocks
INLINE_FORMATTING_TAGS = {
    "strong",
    "b",
    "em",
    "i",
    "u",
    "s",
    "strike",
    "del",
    "span",
    "br",
    "a",
    "sub",
    "sup",
    "code",
    "small",
    "mark",
}

# Tags that are internal to other structures (skip these)
INTERNAL_STRUCTURE_TAGS = {
    "li",
    "thead",
    "tbody",
    "tr",
    "td",
    "th",
    "figcaption",
}

# Selectors for common consent/overlay widgets that should not appear in thumbnails
THUMBNAIL_OVERLAY_BLOCKLIST_CSS = """
[id*="cookie" i],
[class*="cookie" i],
[id*="consent" i],
[class*="consent" i],
[id*="gdpr" i],
[class*="gdpr" i],
[id*="onetrust" i],
[class*="onetrust" i],
[id*="usercentrics" i],
[class*="usercentrics" i],
[id*="didomi" i],
[class*="didomi" i],
[id*="qc-cmp2" i],
[class*="qc-cmp2" i] {
    display: none !important;
    visibility: hidden !important;
    opacity: 0 !important;
    pointer-events: none !important;
}
"""

THUMBNAIL_OVERLAY_SUPPRESSION_SCRIPT = """
(selector) => {
    const slide = document.querySelector(selector);
    if (!slide) {
        return { hidden: 0, reason: "slide_not_found" };
    }

    const keywords = [
        "cookie",
        "consent",
        "gdpr",
        "onetrust",
        "usercentrics",
        "didomi",
        "privacy",
        "trustarc",
        "qc-cmp2",
    ];
    let hiddenCount = 0;

    for (const element of document.querySelectorAll("body *")) {
        if (!(element instanceof HTMLElement)) {
            continue;
        }
        if (slide.contains(element)) {
            continue;
        }

        const className = typeof element.className === "string"
            ? element.className
            : (element.getAttribute("class") || "");
        const signature = `${element.id || ""} ${className}`.toLowerCase();
        const hasKeyword = keywords.some((word) => signature.includes(word));

        const computed = window.getComputedStyle(element);
        const position = computed.position;
        const zIndex = Number.parseInt(computed.zIndex || "0", 10);
        const rect = element.getBoundingClientRect();
        const role = (element.getAttribute("role") || "").toLowerCase();
        const isDialogLike = element.getAttribute("aria-modal") === "true" || role === "dialog";
        const isFixedLike = position === "fixed" || position === "sticky";
        const isLargeBar =
            rect.width >= window.innerWidth * 0.6 &&
            rect.height <= window.innerHeight * 0.35 &&
            (rect.top <= 120 || rect.bottom >= window.innerHeight - 120);
        const isModalSized =
            rect.width >= window.innerWidth * 0.3 &&
            rect.height >= window.innerHeight * 0.2;
        const isOverlayCandidate =
            hasKeyword ||
            (isFixedLike && (zIndex >= 10 || isLargeBar || isDialogLike || isModalSized)) ||
            (isDialogLike && zIndex >= 10);

        if (!isOverlayCandidate) {
            continue;
        }

        element.setAttribute("data-storyline-thumbnail-hidden", "true");
        element.style.setProperty("display", "none", "important");
        element.style.setProperty("visibility", "hidden", "important");
        element.style.setProperty("pointer-events", "none", "important");
        hiddenCount += 1;
    }

    return { hidden: hiddenCount };
}
"""


def _has_includable_child(element: Tag) -> bool:
    """Check if element has child elements that would be included as blocks.

    Used to determine if this element is the "innermost" text-containing element.
    If it has children that would be included, skip this element (not innermost).
    """
    for child in element.children:
        if not isinstance(child, Tag):
            continue
        child_name = child.name.lower() if child.name else ""
        # Skip inline formatting - these don't count
        if child_name in INLINE_FORMATTING_TAGS:
            continue
        # Skip internal structures - these don't count
        if child_name in INTERNAL_STRUCTURE_TAGS:
            continue
        # If child is an image, parent has includable child
        if child_name == "img":
            return True
        # If child has text content, parent has includable child
        if child.get_text(strip=True):
            return True
    return False


def _hide_overlay_elements_for_thumbnail_sync(page: Any, slide_selector: str) -> None:
    """Hide fixed overlays (cookie bars, consent dialogs, sticky headers) before screenshot."""
    try:
        page.add_style_tag(content=THUMBNAIL_OVERLAY_BLOCKLIST_CSS)
        result = page.evaluate(THUMBNAIL_OVERLAY_SUPPRESSION_SCRIPT, slide_selector)
        hidden_count = result.get("hidden", 0) if isinstance(result, dict) else 0
        logger.debug("Suppressed %d overlay element(s) before thumbnail capture", hidden_count)
    except Exception as e:
        logger.debug("Overlay suppression failed during thumbnail capture: %s", e)


async def _hide_overlay_elements_for_thumbnail_async(page: Any, slide_selector: str) -> None:
    """Async variant of overlay suppression for batch thumbnail capture."""
    try:
        await page.add_style_tag(content=THUMBNAIL_OVERLAY_BLOCKLIST_CSS)
        result = await page.evaluate(THUMBNAIL_OVERLAY_SUPPRESSION_SCRIPT, slide_selector)
        hidden_count = result.get("hidden", 0) if isinstance(result, dict) else 0
        logger.debug("Suppressed %d overlay element(s) before thumbnail capture", hidden_count)
    except Exception as e:
        logger.debug("Overlay suppression failed during thumbnail capture: %s", e)


def html_element_discriminator(v: HTMLElementInput) -> str:
    """Discriminator to determine which HTMLElement subclass based on tag name or type field."""
    # First check if it's a direct BeautifulSoup Tag with a tag name
    if isinstance(v, Tag):
        tag_name = v.name.lower() if v.name else ""
        if tag_name == "img":
            return "image"
        elif tag_name == "table":
            return "table"
        elif tag_name in [
            "div",
            "p",
            "span",
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
            "section",
            "article",
            "ul",
            "ol",
        ]:
            return "shape"
        else:
            return "generic"

    # Then check if it's already wrapped with html_element
    elif hasattr(v, "html_element"):
        html_elem = v.html_element
        if isinstance(html_elem, Tag):
            tag_name = html_elem.name.lower() if html_elem.name else ""
            if tag_name == "img":
                return "image"
            elif tag_name == "table":
                return "table"
            elif tag_name in [
                "div",
                "p",
                "span",
                "h1",
                "h2",
                "h3",
                "h4",
                "h5",
                "h6",
                "section",
                "article",
                "ul",
                "ol",
            ]:
                return "shape"
        return "generic"

    # Finally check for type field
    else:
        element_type = getattr(v, "type", None)
        if element_type in [AbstractElementKind.SHAPE, "SHAPE", "shape"]:
            return "shape"
        elif element_type in [AbstractElementKind.IMAGE, "IMAGE", "image"]:
            return "image"
        elif element_type in [AbstractElementKind.TABLE, "TABLE", "table"]:
            return "table"
        else:
            return "generic"


class HTMLAPIClient(AbstractSlidesAPIClient):
    """HTML API client implementation using filesystem operations."""

    def __init__(self):
        # No initialization needed for filesystem-based operations
        self._auto_flush = True

    @property
    def auto_flush(self):
        return self._auto_flush

    @auto_flush.setter
    def auto_flush(self, value: bool):
        # Just store this for consistency with abstract interface
        self._auto_flush = value

    def flush_batch_update(self):
        # No batching needed for filesystem operations
        pass

    def copy_presentation(
        self, presentation_id: str, copy_title: str, folder_id: Optional[str] = None
    ) -> dict:
        """Copy a presentation directory to another location."""
        if not os.path.exists(presentation_id) or not os.path.isdir(presentation_id):
            raise FileNotFoundError(f"Presentation directory not found: {presentation_id}")

        # Determine destination folder
        if folder_id is None:
            # Copy to same folder as source
            dest_folder = os.path.dirname(presentation_id)
        else:
            # Validate folder exists
            if not os.path.exists(folder_id) or not os.path.isdir(folder_id):
                raise FileNotFoundError(f"Destination folder not found: {folder_id}")
            dest_folder = folder_id

        # Create destination path
        dest_path = os.path.join(dest_folder, copy_title)

        # Copy directory (overwrite if exists)
        shutil.copytree(presentation_id, dest_path, dirs_exist_ok=True)

        return {
            "id": dest_path,
            "name": copy_title,
            "parents": [dest_folder] if folder_id else [os.path.dirname(presentation_id)],
        }

    def create_folder(
        self, name: str, ignore_existing: bool = True, parent_folder_id: Optional[str] = None
    ) -> dict:
        """Create a folder in the filesystem."""
        if parent_folder_id is None:
            parent_folder_id = os.getcwd()

        if not os.path.exists(parent_folder_id):
            raise FileNotFoundError(f"Parent folder not found: {parent_folder_id}")

        folder_path = os.path.join(parent_folder_id, name)

        try:
            os.makedirs(folder_path, exist_ok=ignore_existing)
        except FileExistsError:
            if not ignore_existing:
                raise

        return {"id": folder_path, "name": name, "parents": [parent_folder_id]}

    def delete_file(self, file_id: str):
        """Delete a file or directory from the filesystem."""
        if os.path.exists(file_id):
            if os.path.isdir(file_id):
                shutil.rmtree(file_id)
            else:
                os.remove(file_id)

    def set_credentials(self, credentials: AbstractCredentials):
        # Do nothing as this is filesystem-based
        pass

    def get_credentials(self) -> Optional[AbstractCredentials]:
        # Return None as no credentials needed for filesystem operations
        return None

    def replace_text(
        self, slide_ids: list[str], match_text: str, replace_text: str, presentation_id: str
    ):
        """Replace text across specified slides in a presentation."""
        # This would require loading the presentation, finding slides by ID, and replacing text
        # For now, raise NotImplementedError
        raise NotImplementedError("replace_text not yet implemented for HTML adapter")

    @classmethod
    def get_default_api_client(cls) -> "HTMLAPIClient":
        """Get the default API client instance."""
        return cls()

    async def get_presentation_as_pdf(self, presentation_id: str) -> bytes:
        """Get PDF from presentation."""
        raise NotImplementedError("PDF export not implemented for HTML adapter")


class HTMLSpeakerNotes(AbstractSpeakerNotes):
    """HTML speaker notes implementation using data-notes attribute."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_section: Any = Field(exclude=True, default=None)

    def __init__(self, html_section: Tag, **kwargs):
        super().__init__(**kwargs)
        self.html_section = html_section

    def read_text(self, as_markdown: bool = True) -> str:
        """Read text from speaker notes (data-notes attribute)."""
        if not self.html_section:
            return ""

        notes = self.html_section.get("data-notes", "")
        return notes if notes else ""

    def write_text(self, api_client: "HTMLAPIClient", content: str):
        """Write text to speaker notes (data-notes attribute)."""
        if not self.html_section:
            return

        self.html_section["data-notes"] = content


class HTMLElementParent(AbstractElement):
    """Generic concrete element for HTML elements."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_element: Any = Field(exclude=True, default=None)
    html_section: Any = Field(exclude=True, default=None)
    directory_path: Optional[str] = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_html_element(cls, data: HTMLElementInput) -> dict:
        """Convert from BeautifulSoup element to our abstract representation."""
        if isinstance(data, dict):
            # Already converted
            return data
        elif isinstance(data, Tag):
            html_element = data
        elif hasattr(data, "html_element"):
            # Already wrapped element
            return data.__dict__
        else:
            raise ValueError(f"Expected BeautifulSoup Tag, got {type(data)}")

        # Extract basic properties
        object_id = html_element.get("id", "") or html_element.get("data-element-id", "")

        # Get alt text from data attributes
        alt_text_title = html_element.get("data-alt-title", None)
        alt_text_descr = html_element.get("data-alt-description", None)

        return {
            "objectId": object_id,
            "presentation_id": "",  # Will be set by parent
            "slide_id": "",
            "alt_text": AbstractAltText(title=alt_text_title, description=alt_text_descr),
            "type": "generic",
            "html_element": html_element,
        }

    def absolute_size(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        """Get the absolute size of the element by parsing CSS width/height from style attribute."""
        if not self.html_element:
            return (0.0, 0.0)

        # Parse CSS style attribute
        style = self.html_element.get("style", "")
        width_px = self._parse_css_dimension(style, "width")
        height_px = self._parse_css_dimension(style, "height")

        # Convert pixels to requested units (assuming 96 DPI for px to inches)
        if units == OutputUnit.IN or units == "in":
            return (width_px / 96.0, height_px / 96.0)
        elif units == OutputUnit.PX or units == "px":
            return (width_px, height_px)
        else:
            return (width_px / 96.0, height_px / 96.0)

    def _parse_css_dimension(self, style: str, property_name: str) -> float:
        """Parse a CSS dimension value from a style string, returning value in pixels."""
        import re

        # Look for the property in the style string
        pattern = rf"{property_name}\s*:\s*([^;]+)"
        match = re.search(pattern, style, re.IGNORECASE)
        if not match:
            return 0.0

        value_str = match.group(1).strip()

        # Parse numeric value and unit
        num_match = re.match(r"([\d.]+)\s*(px|in|pt|em|rem|%|)", value_str, re.IGNORECASE)
        if not num_match:
            return 0.0

        value = float(num_match.group(1))
        unit = num_match.group(2).lower() if num_match.group(2) else "px"

        # Convert to pixels
        if unit == "px" or unit == "":
            return value
        elif unit == "in":
            return value * 96  # 96 DPI
        elif unit == "pt":
            return value * 96 / 72  # 72 points per inch
        elif unit in ("em", "rem"):
            return value * 16  # Assume 16px base font
        elif unit == "%":
            # Percentage values need parent context, return 0 for now
            return 0.0
        else:
            return value

    def absolute_position(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        """Get the absolute position of the element (simplified - returns 0,0 for HTML)."""
        if not self.html_element:
            return (0.0, 0.0)

        # HTML uses CSS layout, not absolute positioning
        # For now, return simplified values
        # Future: parse CSS left/top from style attribute
        return (0.0, 0.0)

    def create_image_element_like(self, api_client: HTMLAPIClient) -> "HTMLImageElement":
        """Create an image element with the same properties as this element."""
        if not self.html_element:
            logger.warning("Cannot create image element: missing html_element reference")
            raise ValueError("Cannot create image element without html_element reference")

        # Create a new <img> tag - need to find the document root for new_tag
        # Navigate up to find the root BeautifulSoup object which has new_tag method
        parent = self.html_element
        while parent is not None:
            if hasattr(parent, "new_tag") and callable(getattr(parent, "new_tag", None)):
                new_img = parent.new_tag("img")
                break
            parent = parent.parent
        else:
            # Fallback: create a minimal soup with the tag
            soup = BeautifulSoup("<img/>", "lxml")
            new_img = soup.find("img")
        new_img["src"] = "placeholder.png"
        new_img["data-element-name"] = self.html_element.get("data-element-name", "")
        new_img["data-alt-title"] = self.alt_text.title or ""
        new_img["data-alt-description"] = self.alt_text.description or ""

        # Copy style attribute to preserve dimensions
        original_style = self.html_element.get("style", "")
        if original_style:
            new_img["style"] = original_style

        # Replace current element with new image
        self.html_element.replace_with(new_img)

        # Create and return HTMLImageElement wrapper
        image_element = HTMLImageElement(
            objectId=new_img.get("id", ""),
            alt_text=self.alt_text,
            html_element=new_img,
            directory_path=self.directory_path,
        )
        return image_element

    def set_alt_text(
        self,
        api_client: HTMLAPIClient,
        title: str | None = None,
        description: str | None = None,
    ):
        """Set alt text for the element using data attributes."""
        if self.html_element:
            if title is not None:
                self.html_element["data-alt-title"] = title
                self.alt_text.title = title
            if description is not None:
                self.html_element["data-alt-description"] = description
                self.alt_text.description = description


class HTMLShapeElement(AbstractShapeElement, HTMLElementParent):
    """HTML shape element implementation for text-containing elements."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_element: Any = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_html_element(cls, data: HTMLElementInput) -> dict:
        """Convert from BeautifulSoup element."""
        base_data = HTMLElementParent.convert_from_html_element(data)
        base_data["type"] = AbstractElementKind.SHAPE
        return base_data

    @property
    def has_text(self) -> bool:
        """Check if the element has text content."""
        if not self.html_element:
            return False
        text = self.html_element.get_text(strip=True)
        return bool(text)

    def write_text(
        self,
        api_client: HTMLAPIClient,
        content: str,
        autoscale: bool = False,
    ):
        """Write text to the element (supports markdown formatting)."""
        if not self.html_element:
            return

        from gslides_api.adapters.markdown_to_html import apply_markdown_to_html_element

        # Check if content has markdown formatting indicators
        has_markdown_formatting = any(
            marker in content for marker in ["**", "*", "__", "~~", "- ", "1. ", "2. "]
        )

        if has_markdown_formatting:
            # Use markdown parser to convert markdown to HTML with formatting
            apply_markdown_to_html_element(
                markdown_text=content,
                html_element=self.html_element,
                base_style=None,
            )
        else:
            # Simple text replacement (preserves template variables like {account_name})
            self.html_element.clear()
            self.html_element.string = content

    def read_text(self, as_markdown: bool = True) -> str:
        """Read text from the element."""
        if not self.html_element:
            return ""

        if as_markdown:
            from gslides_api.adapters.markdown_to_html import convert_html_to_markdown

            # Convert HTML formatting to markdown
            return convert_html_to_markdown(self.html_element)
        else:
            # Extract plain text content
            text = self.html_element.get_text(separator="\n", strip=False)
            return text

    def styles(self, skip_whitespace: bool = True) -> Optional[List[dict]]:
        """Extract style information from the element (simplified)."""
        if not self.html_element:
            return None

        # Simplified: return basic style info
        # Future: parse inline CSS styles from style attribute
        text = self.html_element.get_text()
        if skip_whitespace and not text.strip():
            return None

        return [{"text": text}]


class HTMLImageElement(AbstractImageElement, HTMLElementParent):
    """HTML image element implementation for <img> tags."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_element: Any = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_html_element(cls, data: HTMLElementInput) -> dict:
        """Convert from BeautifulSoup <img> element."""
        base_data = HTMLElementParent.convert_from_html_element(data)
        base_data["type"] = AbstractElementKind.IMAGE
        return base_data

    def replace_image(
        self,
        api_client: HTMLAPIClient,
        file: str | None = None,
        url: str | None = None,
    ):
        """Replace the image in this element."""
        if (
            not self.html_element
            or not isinstance(self.html_element.name, str)
            or self.html_element.name.lower() != "img"
        ):
            logger.warning("Cannot replace image: element is not an <img> tag")
            return

        if file and os.path.exists(file):
            # Copy file to images/ subdirectory
            if not self.directory_path:
                logger.warning("Cannot replace image: no directory_path set")
                return

            images_dir = os.path.join(self.directory_path, "images")
            os.makedirs(images_dir, exist_ok=True)

            # Copy file
            filename = os.path.basename(file)
            dest_path = os.path.join(images_dir, filename)
            shutil.copy2(file, dest_path)

            # Update src attribute (relative path)
            self.html_element["src"] = f"images/{filename}"
            logger.info(f"Replaced image with {file}, src set to images/{filename}")

        elif url:
            # For URL images, download first using utility with retries
            content, _ = download_binary_file(url)
            # Save to temp file and use file path method
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_file:
                temp_file.write(content)
                temp_file_path = temp_file.name

            try:
                self.replace_image(api_client, file=temp_file_path)
            finally:
                os.unlink(temp_file_path)

    def get_image_data(self):
        """Get the image data from the HTML element."""
        if (
            not self.html_element
            or not isinstance(self.html_element.name, str)
            or self.html_element.name.lower() != "img"
        ):
            return None

        try:
            # Get src attribute
            src = self.html_element.get("src", "")
            if not src:
                return None

            # Determine if it's a URL or local file
            if src.startswith("http://") or src.startswith("https://"):
                # Download from URL
                content, _ = download_binary_file(src)
                mime_type = "image/png"  # Default, could be detected from content
                return ImageData(content=content, mime_type=mime_type)
            else:
                # Local file (relative to directory)
                if not self.directory_path:
                    return None

                file_path = os.path.join(self.directory_path, src)
                if not os.path.exists(file_path):
                    return None

                with open(file_path, "rb") as f:
                    content = f.read()

                # Detect mime type from extension
                ext = os.path.splitext(file_path)[1].lower()
                mime_type_map = {
                    ".png": "image/png",
                    ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg",
                    ".gif": "image/gif",
                    ".svg": "image/svg+xml",
                }
                mime_type = mime_type_map.get(ext, "image/png")

                return ImageData(content=content, mime_type=mime_type)

        except Exception as e:
            logger.error(f"Error getting image data: {e}")
            return None


class HTMLTableElement(AbstractTableElement, HTMLElementParent):
    """HTML table element implementation for <table> tags."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_element: Any = Field(exclude=True, default=None)

    @model_validator(mode="before")
    @classmethod
    def convert_from_html_element(cls, data: HTMLElementInput) -> dict:
        """Convert from BeautifulSoup <table> element."""
        base_data = HTMLElementParent.convert_from_html_element(data)
        base_data["type"] = AbstractElementKind.TABLE
        return base_data

    def _get_soup(self) -> BeautifulSoup:
        """Get a BeautifulSoup object for creating new tags."""
        # Create a temporary soup for tag creation
        return BeautifulSoup("", "lxml")

    def resize(
        self,
        api_client: HTMLAPIClient,
        rows: int,
        cols: int,
        fix_width: bool = True,
        fix_height: bool = True,
        target_height_in: float | None = None,
    ) -> float:
        """Resize the table to the specified dimensions.

        Args:
            target_height_in: Ignored for HTML tables (they auto-size).

        Returns:
            Font scale factor (1.0 since HTML doesn't support font scaling during resize).
        """
        if (
            not self.html_element
            or not isinstance(self.html_element.name, str)
            or self.html_element.name.lower() != "table"
        ):
            return 1.0

        try:
            soup = self._get_soup()

            # Get or create tbody
            tbody = self.html_element.find("tbody")
            if not tbody:
                tbody = soup.new_tag("tbody")
                self.html_element.append(tbody)

            current_rows = tbody.find_all("tr", recursive=False)
            current_row_count = len(current_rows)

            # Adjust rows
            if rows > current_row_count:
                # Add rows
                for _ in range(rows - current_row_count):
                    new_row = soup.new_tag("tr")
                    for _ in range(cols):
                        new_cell = soup.new_tag("td")
                        new_row.append(new_cell)
                    tbody.append(new_row)
            elif rows < current_row_count:
                # Remove rows
                for row in current_rows[rows:]:
                    row.decompose()

            # Adjust columns in all rows
            for row in tbody.find_all("tr", recursive=False):
                cells = row.find_all(["td", "th"], recursive=False)
                current_col_count = len(cells)

                if cols > current_col_count:
                    # Add cells
                    for _ in range(cols - current_col_count):
                        new_cell = soup.new_tag("td")
                        row.append(new_cell)
                elif cols < current_col_count:
                    # Remove cells
                    for cell in cells[cols:]:
                        cell.decompose()

        except Exception as e:
            logger.error(f"Error resizing table: {e}")

        return 1.0

    def update_content(
        self,
        api_client: HTMLAPIClient,
        markdown_content: MarkdownTableElement,
        check_shape: bool = True,
        font_scale_factor: float = 1.0,
    ):
        """Update the table content with markdown data.

        Args:
            font_scale_factor: Font scale factor (currently unused for HTML, but kept for interface conformance).
        """
        if (
            not self.html_element
            or not isinstance(self.html_element.name, str)
            or self.html_element.name.lower() != "table"
        ):
            return

        try:
            soup = self._get_soup()

            # Get table data from markdown content
            if hasattr(markdown_content, "content") and hasattr(
                markdown_content.content, "headers"
            ):
                headers = markdown_content.content.headers
                data_rows = markdown_content.content.rows
            else:
                # Fallback for old interface
                headers = (
                    markdown_content.rows[0]
                    if hasattr(markdown_content, "rows") and markdown_content.rows
                    else []
                )
                data_rows = (
                    markdown_content.rows[1:]
                    if hasattr(markdown_content, "rows") and len(markdown_content.rows) > 1
                    else []
                )

            if not headers:
                return

            # Ensure table has thead and tbody
            thead = self.html_element.find("thead")
            if not thead:
                thead = soup.new_tag("thead")
                self.html_element.insert(0, thead)

            tbody = self.html_element.find("tbody")
            if not tbody:
                tbody = soup.new_tag("tbody")
                self.html_element.append(tbody)

            # Clear existing content
            thead.clear()
            tbody.clear()

            # Add header row
            header_row = soup.new_tag("tr")
            for header in headers:
                th = soup.new_tag("th")
                th.string = str(header) if header is not None else ""
                header_row.append(th)
            thead.append(header_row)

            # Add data rows
            for row_data in data_rows:
                tr = soup.new_tag("tr")
                for cell_data in row_data:
                    td = soup.new_tag("td")
                    td.string = str(cell_data) if cell_data is not None else ""
                    tr.append(td)
                tbody.append(tr)

        except Exception as e:
            logger.error(f"Error updating table content: {e}")

    def get_horizontal_border_weight(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Get weight of horizontal borders in specified units.

        For HTML tables, borders are handled via CSS and don't have a fixed weight
        that contributes to layout height, so we return 0.
        """
        return 0.0

    def get_row_count(self) -> int:
        """Get current number of rows."""
        if not self.html_element:
            return 0
        tbody = self.html_element.find("tbody")
        if tbody:
            return len(tbody.find_all("tr", recursive=False))
        return 0

    def get_column_count(self) -> int:
        """Get current number of columns."""
        if not self.html_element:
            return 0
        # Check thead first, then tbody
        thead = self.html_element.find("thead")
        if thead:
            header_row = thead.find("tr")
            if header_row:
                return len(header_row.find_all(["th", "td"], recursive=False))
        tbody = self.html_element.find("tbody")
        if tbody:
            first_row = tbody.find("tr")
            if first_row:
                return len(first_row.find_all(["td", "th"], recursive=False))
        return 0

    def to_markdown_element(self, name: str | None = None) -> MarkdownTableElement:
        """Convert HTML table to markdown table element."""
        if (
            not self.html_element
            or not isinstance(self.html_element.name, str)
            or self.html_element.name.lower() != "table"
        ):
            raise ValueError("HTMLTableElement has no valid <table> element")

        # Extract headers from thead
        thead = self.html_element.find("thead")
        headers = []
        if thead:
            header_row = thead.find("tr")
            if header_row:
                headers = [th.get_text(strip=True) for th in header_row.find_all(["th", "td"])]

        # Extract rows from tbody
        tbody = self.html_element.find("tbody")
        rows = []
        if tbody:
            for tr in tbody.find_all("tr"):
                row = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
                rows.append(row)

        # Create TableData from extracted headers and rows
        if not headers and not rows:
            table_data = None
        else:
            table_data = TableData(headers=headers, rows=rows)

        # Create MarkdownTableElement with TableData
        markdown_elem = MarkdownTableElement(
            name=name or self.alt_text.title or "Table",
            content=table_data,
        )

        return markdown_elem


# Discriminated union type for concrete elements
HTMLElement = Annotated[
    Union[
        Annotated[HTMLShapeElement, PydanticTag("shape")],
        Annotated[HTMLImageElement, PydanticTag("image")],
        Annotated[HTMLTableElement, PydanticTag("table")],
        Annotated[HTMLElementParent, PydanticTag("generic")],
    ],
    Discriminator(html_element_discriminator),
]

# TypeAdapter for validating the discriminated union
_html_element_adapter = TypeAdapter(HTMLElement)


def validate_html_element(html_element: Tag) -> HTMLElement:
    """Create the appropriate concrete element from a BeautifulSoup Tag."""
    element_type = html_element_discriminator(html_element)

    if element_type == "shape":
        data = HTMLShapeElement.convert_from_html_element(html_element)
        return HTMLShapeElement(**data)
    elif element_type == "image":
        data = HTMLImageElement.convert_from_html_element(html_element)
        return HTMLImageElement(**data)
    elif element_type == "table":
        data = HTMLTableElement.convert_from_html_element(html_element)
        return HTMLTableElement(**data)
    else:
        data = HTMLElementParent.convert_from_html_element(html_element)
        return HTMLElementParent(**data)


class HTMLSlide(AbstractSlide):
    """HTML slide implementation representing a <section> element."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_section: Any = Field(exclude=True, default=None)
    html_soup: Any = Field(exclude=True, default=None)
    directory_path: Optional[str] = Field(exclude=True, default=None)

    def __init__(self, html_section: Tag, html_soup: BeautifulSoup, directory_path: str, **kwargs):
        # Extract all content elements (innermost text-containing elements and images)
        elements = []

        for child in html_section.descendants:
            if isinstance(child, Tag):
                # Skip nested sections
                if child.name == "section":
                    continue

                tag_name = child.name.lower() if child.name else ""

                # Skip inline formatting elements
                if tag_name in INLINE_FORMATTING_TAGS:
                    continue

                # Skip internal structure elements
                if tag_name in INTERNAL_STRUCTURE_TAGS:
                    continue

                # For images, always include
                if tag_name == "img":
                    try:
                        html_elem = validate_html_element(child)
                        html_elem.directory_path = directory_path
                        html_elem.html_section = html_section
                        elements.append(html_elem)
                    except Exception as e:
                        logger.warning(f"Could not convert image element: {e}")
                    continue

                # For other elements, only include if:
                # 1. They have text content
                # 2. They are the innermost (no includable children)
                text_content = child.get_text(strip=True)
                if not text_content:
                    continue

                # Skip if this element has includable children (not innermost)
                if _has_includable_child(child):
                    continue

                try:
                    html_elem = validate_html_element(child)
                    html_elem.directory_path = directory_path
                    html_elem.html_section = html_section
                    elements.append(html_elem)
                except Exception as e:
                    logger.warning(f"Could not convert element {tag_name}: {e}")

        # Get speaker notes
        speaker_notes = HTMLSpeakerNotes(html_section)

        # Get slide properties
        is_skipped = html_section.get("data-skip", "").lower() == "true"
        slide_properties = AbstractSlideProperties(isSkipped=is_skipped)

        # Get object ID
        object_id = html_section.get("id", "") or html_section.get("data-slide-id", "")

        super().__init__(
            elements=elements,
            objectId=object_id,
            slideProperties=slide_properties,
            speaker_notes=speaker_notes,
        )

        self.html_section = html_section
        self.html_soup = html_soup
        self.directory_path = directory_path

    @property
    def page_elements_flat(self) -> list[HTMLElementParent]:
        """Flatten the elements tree into a list."""
        return self.elements

    def thumbnail(
        self, api_client: HTMLAPIClient, size: str, include_data: bool = False
    ) -> AbstractThumbnail:
        """Generate a thumbnail of the slide using Playwright."""
        if not self.directory_path:
            logger.warning("Cannot generate thumbnail: no directory_path set")
            return AbstractThumbnail(
                contentUrl="placeholder_thumbnail.png",
                width=320,
                height=240,
                mime_type="image/png",
                content=None,
            )

        html_file = os.path.join(self.directory_path, "index.html")
        if not os.path.exists(html_file):
            logger.warning(f"Cannot generate thumbnail: index.html not found at {html_file}")
            return AbstractThumbnail(
                contentUrl="placeholder_thumbnail.png",
                width=320,
                height=240,
                mime_type="image/png",
                content=None,
            )

        file_url = f"file://{os.path.abspath(html_file)}"

        # Use Playwright to capture the slide
        from playwright.sync_api import sync_playwright

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                try:
                    context = browser.new_context(
                        device_scale_factor=2,
                        viewport={"width": 1280, "height": 720},
                    )
                    page = context.new_page()
                    page.goto(file_url, wait_until="networkidle")
                    page.wait_for_timeout(1000)

                    # Find the slide section by ID or position
                    slide_id = self.objectId
                    if slide_id:
                        safe_id = slide_id.replace(":", "\\:")
                        selector = f"section#{safe_id}"
                    else:
                        # Fallback: try to find the section element directly
                        selector = "section"

                    _hide_overlay_elements_for_thumbnail_sync(page, selector)
                    slide_element = page.query_selector(selector)
                    if slide_element:
                        png_bytes = slide_element.screenshot(type="png")

                        # Get dimensions from the image
                        from PIL import Image
                        import io as _io

                        img = Image.open(_io.BytesIO(png_bytes))
                        img_width, img_height = img.size
                        img.close()

                        return AbstractThumbnail(
                            contentUrl=file_url,
                            width=img_width,
                            height=img_height,
                            mime_type="image/png",
                            content=png_bytes if include_data else None,
                        )
                    else:
                        logger.warning(f"Could not find slide element with selector: {selector}")
                finally:
                    browser.close()
        except Exception as e:
            logger.error(f"Error generating HTML thumbnail: {e}")

        # Return placeholder on failure
        return AbstractThumbnail(
            contentUrl="placeholder_thumbnail.png",
            width=320,
            height=240,
            mime_type="image/png",
            content=None,
        )


class HTMLPresentation(AbstractPresentation):
    """HTML presentation implementation representing a directory with index.html."""

    model_config = ConfigDict(arbitrary_types_allowed=True)

    html_soup: Any = Field(exclude=True, default=None)
    directory_path: Optional[str] = None
    # Storage URLs set after upload (for API adapter layer)
    uploaded_html_url: Optional[str] = None
    uploaded_zip_url: Optional[str] = None

    def __init__(
        self,
        html_soup: BeautifulSoup,
        directory_path: str,
    ):
        # Extract all top-level <section> elements as slides
        slides = []
        for section in html_soup.find_all("section", recursive=True):
            # Only process top-level sections (not nested)
            if section.find_parent("section") is None:
                try:
                    slide = HTMLSlide(section, html_soup, directory_path)
                    slides.append(slide)
                except Exception as e:
                    logger.warning(f"Could not convert section to slide: {e}")
                    continue

        # Extract presentation metadata
        presentation_id = directory_path
        title_tag = html_soup.find("title")
        title = title_tag.get_text(strip=True) if title_tag else os.path.basename(directory_path)

        super().__init__(
            slides=slides,
            presentationId=presentation_id,
            revisionId=None,  # HTML doesn't have revision IDs
            title=title,
        )

        self.html_soup = html_soup
        self.directory_path = directory_path

    @property
    def url(self) -> str:
        """Return the file path as URL (file system based)."""
        if self.directory_path:
            html_file = os.path.join(self.directory_path, "index.html")
            return f"file://{os.path.abspath(html_file)}"
        else:
            raise ValueError("No directory path specified for presentation")

    def slide_height(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Return slide height in specified units.

        HTML slides don't have a fixed height - returns default presentation height.
        """
        # Default to standard presentation height (7.5 inches for 4:3 aspect ratio)
        default_height_in = 7.5
        default_height_emu = to_emu(default_height_in, OutputUnit.IN)
        return from_emu(default_height_emu, units)

    @classmethod
    def from_id(
        cls,
        api_client: HTMLAPIClient,
        presentation_id: str,
    ) -> "HTMLPresentation":
        """Load presentation from directory path."""
        # presentation_id is the directory path
        if not os.path.exists(presentation_id) or not os.path.isdir(presentation_id):
            raise FileNotFoundError(f"Presentation directory not found: {presentation_id}")

        # Load index.html
        html_file = os.path.join(presentation_id, "index.html")
        if not os.path.exists(html_file):
            raise FileNotFoundError(f"index.html not found in {presentation_id}")

        try:
            with open(html_file, "r", encoding="utf-8") as f:
                html_content = f.read()

            html_soup = BeautifulSoup(html_content, "lxml")
            return cls(html_soup, presentation_id)
        except Exception as e:
            raise ValueError(f"Could not load presentation from {presentation_id}: {e}")

    def copy_via_drive(
        self,
        api_client: HTMLAPIClient,
        copy_title: str,
        folder_id: Optional[str] = None,
    ) -> "HTMLPresentation":
        """Copy presentation to another location."""
        if not self.directory_path:
            raise ValueError("Cannot copy presentation without a directory path")

        # Use the API client to copy the directory
        copy_result = api_client.copy_presentation(self.directory_path, copy_title, folder_id)

        # Load the copied presentation
        copied_presentation = HTMLPresentation.from_id(api_client, copy_result["id"])

        return copied_presentation

    def sync_from_cloud(self, api_client: HTMLAPIClient):
        """Re-read presentation from filesystem."""
        if not self.directory_path or not os.path.exists(self.directory_path):
            return

        # Reload from file
        html_file = os.path.join(self.directory_path, "index.html")
        with open(html_file, "r", encoding="utf-8") as f:
            html_content = f.read()

        # Update our internal representation
        html_soup = BeautifulSoup(html_content, "lxml")
        self.html_soup = html_soup

        # Rebuild slides
        slides = []
        for section in html_soup.find_all("section", recursive=True):
            if section.find_parent("section") is None:
                try:
                    slide = HTMLSlide(section, html_soup, self.directory_path)
                    slides.append(slide)
                except Exception as e:
                    logger.warning(f"Could not convert section during sync: {e}")
                    continue

        self.slides = slides

        # Update metadata
        title_tag = html_soup.find("title")
        self.title = (
            title_tag.get_text(strip=True) if title_tag else os.path.basename(self.directory_path)
        )

    def save(self, api_client: HTMLAPIClient) -> None:
        """Save/persist all changes made to this presentation."""
        if not self.directory_path:
            raise ValueError("No directory path specified for saving")

        html_file = os.path.join(self.directory_path, "index.html")

        # Ensure directory exists
        os.makedirs(self.directory_path, exist_ok=True)

        # Save the HTML
        with open(html_file, "w", encoding="utf-8") as f:
            f.write(self.html_soup.prettify())

    def insert_copy(
        self,
        source_slide: AbstractSlide,
        api_client: HTMLAPIClient,
        insertion_index: int | None = None,
    ) -> AbstractSlide:
        """Insert a copy of a slide into this presentation."""
        if not isinstance(source_slide, HTMLSlide):
            raise ValueError("Can only copy HTMLSlide instances")

        # Deep copy the section
        new_section = copy.deepcopy(source_slide.html_section)

        # Generate a unique ID for the copied section to avoid ID collisions
        unique_id = f"slide-{uuid.uuid4().hex[:8]}"
        new_section["id"] = unique_id

        # Insert into soup at the specified index
        if insertion_index is None:
            self.html_soup.body.append(new_section)
        else:
            sections = [
                s
                for s in self.html_soup.find_all("section", recursive=True)
                if s.find_parent("section") is None
            ]
            if insertion_index < len(sections):
                sections[insertion_index].insert_before(new_section)
            else:
                self.html_soup.body.append(new_section)

        # Create new slide wrapper
        new_slide = HTMLSlide(new_section, self.html_soup, self.directory_path)

        # Update our slides list
        if insertion_index is None:
            self.slides.append(new_slide)
        else:
            self.slides.insert(insertion_index, new_slide)

        return new_slide

    def delete_slide(self, slide: Union[HTMLSlide, int], api_client: HTMLAPIClient):
        """Delete a slide from the presentation."""
        if isinstance(slide, int):
            slide = self.slides[slide]

        if isinstance(slide, HTMLSlide):
            # Remove from DOM
            slide.html_section.decompose()
            # Remove from our slides list
            self.slides.remove(slide)

    def delete_slides(self, slides: List[Union[HTMLSlide, int]], api_client: HTMLAPIClient):
        """Delete multiple slides from the presentation."""
        # Convert all indices to slide objects first to avoid index shifting issues
        slides_to_delete = []
        for slide in slides:
            if isinstance(slide, int):
                slides_to_delete.append(self.slides[slide])
            else:
                slides_to_delete.append(slide)

        # Now delete all slides
        for slide in slides_to_delete:
            self.delete_slide(slide, api_client)

    def move_slide(
        self,
        slide: Union[HTMLSlide, int],
        insertion_index: int,
        api_client: HTMLAPIClient,
    ):
        """Move a slide to a new position within the presentation."""
        if isinstance(slide, int):
            slide = self.slides[slide]

        if isinstance(slide, HTMLSlide):
            # Extract the section from the DOM
            section = slide.html_section.extract()

            # Insert at new position
            sections = [
                s
                for s in self.html_soup.find_all("section", recursive=True)
                if s.find_parent("section") is None
            ]
            if insertion_index < len(sections):
                sections[insertion_index].insert_before(section)
            else:
                self.html_soup.body.append(section)

            # Update local slides list order
            self.slides.remove(slide)
            self.slides.insert(insertion_index, slide)

    def duplicate_slide(self, slide: Union[HTMLSlide, int], api_client: HTMLAPIClient) -> HTMLSlide:
        """Duplicate a slide within the presentation."""
        if isinstance(slide, int):
            slide = self.slides[slide]

        if isinstance(slide, HTMLSlide):
            # Deep copy the section
            new_section = copy.deepcopy(slide.html_section)

            # Append to DOM
            self.html_soup.body.append(new_section)

            # Create new slide wrapper
            new_slide = HTMLSlide(new_section, self.html_soup, self.directory_path)
            self.slides.append(new_slide)

            return new_slide
        else:
            raise ValueError("slide must be an HTMLSlide or int")

    async def get_slide_thumbnails(
        self,
        api_client: "HTMLAPIClient",
        slides: Optional[List["AbstractSlide"]] = None,
    ) -> List[AbstractThumbnail]:
        """Get thumbnails for slides using a single Playwright browser session.

        This is more efficient than calling thumbnail() for each slide individually
        because it opens the browser once and captures all slides in one session.

        Args:
            api_client: The HTML API client
            slides: Optional list of slides to get thumbnails for. If None, uses all slides.

        Returns:
            List of AbstractThumbnail objects with image data
        """
        import io as _io

        from PIL import Image
        from playwright.async_api import async_playwright

        target_slides = slides if slides is not None else self.slides
        thumbnails = []

        if not target_slides:
            return thumbnails

        if not self.directory_path:
            logger.warning("Cannot generate thumbnails: no directory_path set")
            return [self._create_placeholder_thumbnail() for _ in target_slides]

        html_file = os.path.join(self.directory_path, "index.html")
        if not os.path.exists(html_file):
            logger.warning(f"Cannot generate thumbnails: index.html not found at {html_file}")
            return [self._create_placeholder_thumbnail() for _ in target_slides]

        html_file_url = f"file://{os.path.abspath(html_file)}"

        logger.info(
            "Generating HTML thumbnails for %d slides from %s", len(target_slides), html_file_url
        )

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            try:
                context = await browser.new_context(
                    device_scale_factor=2,  # 2x for good quality thumbnails
                    viewport={"width": 1280, "height": 720},
                )
                page = await context.new_page()

                # Navigate to the HTML file
                await page.goto(html_file_url, wait_until="networkidle")
                await page.wait_for_timeout(1000)  # Wait for any JS rendering

                # Generate thumbnail for each slide
                for i, slide in enumerate(target_slides):
                    slide_id = slide.objectId

                    # Build selector - prefer ID-based, fallback to index
                    if slide_id:
                        # CSS ID selectors need escaping for special characters
                        safe_id = slide_id.replace(":", "\\:")
                        selector = f"section#{safe_id}"
                    else:
                        # Fallback: use nth-of-type selector
                        selector = f"section:nth-of-type({i + 1})"

                    try:
                        await _hide_overlay_elements_for_thumbnail_async(page, selector)
                        slide_element = await page.query_selector(selector)
                        if slide_element:
                            # Take screenshot of the slide section
                            png_bytes = await slide_element.screenshot(type="png")

                            # Get dimensions from image
                            img = Image.open(_io.BytesIO(png_bytes))
                            img_width, img_height = img.size
                            img.close()

                            thumbnail = AbstractThumbnail(
                                contentUrl=html_file_url,
                                width=img_width,
                                height=img_height,
                                mime_type="image/png",
                                content=png_bytes,
                                file_size=len(png_bytes),
                            )
                            thumbnails.append(thumbnail)
                            logger.debug(
                                "Generated thumbnail for slide %s (%dx%d)",
                                slide_id,
                                img_width,
                                img_height,
                            )
                        else:
                            logger.warning(
                                "Could not find slide element with selector: %s", selector
                            )
                            thumbnails.append(self._create_placeholder_thumbnail())
                    except Exception as e:
                        logger.error("Failed to capture thumbnail for slide %s: %s", slide_id, e)
                        thumbnails.append(self._create_placeholder_thumbnail())
            finally:
                await browser.close()

        logger.info("Generated %d HTML thumbnails", len(thumbnails))
        return thumbnails

    def _create_placeholder_thumbnail(self) -> AbstractThumbnail:
        """Create a placeholder thumbnail for failed captures."""
        import io as _io

        from PIL import Image

        # Create a simple gray placeholder image
        img = Image.new("RGB", (320, 240), color=(200, 200, 200))
        buffer = _io.BytesIO()
        img.save(buffer, format="PNG")
        png_bytes = buffer.getvalue()
        img.close()

        return AbstractThumbnail(
            contentUrl="",
            width=320,
            height=240,
            mime_type="image/png",
            content=png_bytes,
            file_size=len(png_bytes),
        )
