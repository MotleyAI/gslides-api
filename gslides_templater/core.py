"""
Core functionality for Google Slides Templater.
Main module providing Google Slides API integration and template system.
"""

import json
import logging
import random
import re
import threading
import time
import uuid
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Any, Optional

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from pydantic import BaseModel, Field

from .auth import authenticate, AuthConfig
from .markdown_processor import MarkdownProcessor, MarkdownConfig

logger = logging.getLogger(__name__)


class SlidesConfig(BaseModel):
    """Configuration for Slides operations"""
    max_retries: int = Field(default=3, ge=1, le=10)
    rate_limit_delay: float = Field(default=0.1, ge=0.01, le=1.0)
    batch_size: int = Field(default=50, ge=1, le=100)
    max_total_requests: int = Field(default=1000, ge=1, le=10000)
    request_timeout: float = Field(default=30.0, ge=5.0, le=300.0)


class LayoutConfig(BaseModel):
    """Configuration for slide layout (in Points, converted to EMU internally)"""
    slide_width: int = Field(default=720, ge=100, le=1920)  # Points
    slide_height: int = Field(default=540, ge=100, le=1440)  # Points
    margin_x: int = Field(default=50, ge=0, le=100)  # Points
    margin_y: int = Field(default=50, ge=0, le=100)  # Points
    default_width: int = Field(default=620, ge=100, le=1000)  # Points
    default_height: int = Field(default=200, ge=50, le=500)  # Points
    header_height: int = Field(default=80, ge=30, le=150)  # Points
    text_line_height: int = Field(default=25, ge=15, le=50)  # Points


class ElementPosition(BaseModel):
    """Position and size of slide element"""
    x: float = Field(ge=0)
    y: float = Field(ge=0)
    width: float = Field(gt=0)
    height: float = Field(gt=0)
    layer: int = Field(default=0, ge=0)


class SlideElement(BaseModel):
    """Slide element with position and content"""
    element_id: str
    element_type: str
    position: ElementPosition
    content: str = ""
    style: Dict[str, Any] = Field(default_factory=dict)


class SlidesAPIError(Exception):
    """Base exception for Slides API errors"""
    pass


class RateLimitExceededError(SlidesAPIError):
    """Rate limit exceeded"""
    pass


class MaxRetriesExceededError(SlidesAPIError):
    """Maximum retries exceeded"""
    pass


class TemplateValidationError(SlidesAPIError):
    """Template validation failed"""
    pass


def _is_safe_path(filepath: str, base_dir: str = ".") -> bool:
    """Check if filepath is safe (no path traversal)"""
    try:
        base_path = Path(base_dir).resolve()
        target_path = Path(filepath).resolve()
        return target_path.is_relative_to(base_path)
    except (ValueError, OSError):
        return False


def _emu_to_points(emu_value: float) -> float:
    """Convert EMU to Points (1 Point = 12700 EMU)"""
    return emu_value / 12700


def _points_to_emu(points_value: float) -> float:
    """Convert Points to EMU (1 Point = 12700 EMU)"""
    return points_value * 12700


class SlidesTemplater:
    """
    Main class for working with Google Slides API and template system.

    Supports:
    - Creating and editing presentations
    - Full Markdown ↔ Google Slides conversion
    - Smart template system with automatic placeholder detection
    - Preserving all formatting styles
    """

    def __init__(self, credentials,
                 slides_config: Optional[SlidesConfig] = None,
                 layout_config: Optional[LayoutConfig] = None,
                 markdown_config: Optional[MarkdownConfig] = None):
        """
        Initialize SlidesTemplater.

        Args:
            credentials: Google API credentials
            slides_config: Configuration for Slides operations
            layout_config: Configuration for slide layout
            markdown_config: Configuration for Markdown processing
        """
        if hasattr(credentials, 'credentials'):
            self.credentials = credentials.credentials
        else:
            self.credentials = credentials

        self.slides_config = slides_config or SlidesConfig()
        self.layout_config = layout_config or LayoutConfig()

        self.slides_service = build('slides', 'v1', credentials=self.credentials)
        self.drive_service = build('drive', 'v3', credentials=self.credentials)

        self.markdown_processor = MarkdownProcessor(markdown_config)

        self._last_request_time = 0
        self._request_lock = threading.Lock()
        self._cache_lock = threading.Lock()

        # For backward compatibility
        self.header_sizes = self.markdown_processor.header_sizes

    @classmethod
    def from_credentials(cls, auth_config: Optional[AuthConfig] = None, **kwargs) -> 'SlidesTemplater':
        """
        Create SlidesTemplater from credentials configuration.

        Args:
            auth_config: Authentication configuration
            **kwargs: Legacy parameters or additional configs

        Returns:
            Configured SlidesTemplater instance
        """
        # Extract auth parameters
        auth_params = {}
        slides_params = {}
        layout_params = {}
        markdown_params = {}

        for key, value in kwargs.items():
            if key in ['service_account_path', 'oauth_credentials_path', 'token_path']:
                auth_params[key] = value
            elif key.startswith('slides_'):
                slides_params[key[7:]] = value
            elif key.startswith('layout_'):
                layout_params[key[7:]] = value
            elif key.startswith('markdown_'):
                markdown_params[key[9:]] = value

        if auth_config is None:
            # Handle legacy parameter names
            legacy_mapping = {
                'service_account_path': 'service_account_file',
                'oauth_credentials_path': 'credentials_path'
            }

            for old_name, new_name in legacy_mapping.items():
                if old_name in auth_params:
                    auth_params[new_name] = auth_params.pop(old_name)

            auth_config = AuthConfig(**auth_params)

        credentials = authenticate(auth_config)

        # Create config objects
        slides_config = SlidesConfig(**slides_params) if slides_params else None
        layout_config = LayoutConfig(**layout_params) if layout_params else None
        markdown_config = MarkdownConfig(**markdown_params) if markdown_params else None

        return cls(credentials, slides_config, layout_config, markdown_config)

    def _rate_limit(self):
        """Apply rate limiting to API requests (thread-safe)."""
        with self._request_lock:
            elapsed = time.time() - self._last_request_time
            if elapsed < self.slides_config.rate_limit_delay:
                time.sleep(self.slides_config.rate_limit_delay - elapsed)
            self._last_request_time = time.time()

    def _make_request(self, func, *args, **kwargs):
        """Execute request with retry on errors and proper logging."""
        self._rate_limit()

        for attempt in range(self.slides_config.max_retries):
            try:
                result = func(*args, **kwargs).execute()
                if attempt > 0:
                    logger.info(f"Request succeeded on attempt {attempt + 1}")
                return result

            except HttpError as e:
                if e.resp.status == 429:
                    if attempt == self.slides_config.max_retries - 1:
                        raise RateLimitExceededError("Rate limit exceeded after maximum retries")

                    # Exponential backoff with jitter
                    wait_time = (2 ** attempt) + random.uniform(0, 1)
                    logger.warning(
                        f"Rate limited, retry {attempt + 1}/{self.slides_config.max_retries} in {wait_time:.1f}s")
                    time.sleep(wait_time)
                    continue

                elif e.resp.status in [500, 502, 503, 504]:
                    if attempt == self.slides_config.max_retries - 1:
                        raise MaxRetriesExceededError(
                            f"Server error after {self.slides_config.max_retries} retries: {e}")

                    wait_time = (2 ** attempt) + random.uniform(0, 0.5)
                    logger.warning(
                        f"Server error, retry {attempt + 1}/{self.slides_config.max_retries} in {wait_time:.1f}s: {e}")
                    time.sleep(wait_time)
                    continue
                else:
                    # Client error - don't retry
                    logger.error(f"Client error (status {e.resp.status}): {e}")
                    raise

            except Exception as e:
                logger.error(f"Unexpected error on attempt {attempt + 1}: {e}")
                if attempt == self.slides_config.max_retries - 1:
                    raise
                time.sleep(2 ** attempt)

        raise MaxRetriesExceededError(f"Maximum retry attempts exceeded ({self.slides_config.max_retries})")

    def create_presentation(self, title: str = "New Presentation") -> str:
        """
        Create new presentation.

        Args:
            title: Presentation title

        Returns:
            ID of created presentation
        """
        if not title or len(title.strip()) == 0:
            title = "New Presentation"

        # Validate title length - увеличено ограничение
        if len(title) > 1000:
            title = title[:1000]

        body = {'title': title.strip()}
        result = self._make_request(
            self.slides_service.presentations().create,
            body=body
        )

        presentation_id = result['presentationId']
        logger.info(f"Created presentation '{title}' with ID: {presentation_id}")
        return presentation_id

    @lru_cache(maxsize=128)
    def get_presentation(self, presentation_id: str) -> Dict[str, Any]:
        """
        Get presentation data (cached).

        Args:
            presentation_id: Presentation ID

        Returns:
            Presentation data from Google Slides API
        """
        return self._make_request(
            self.slides_service.presentations().get,
            presentationId=presentation_id
        )

    def copy_presentation(self, presentation_id: str, title: str) -> str:
        """
        Copy presentation.

        Args:
            presentation_id: Source presentation ID
            title: New presentation title

        Returns:
            ID of copied presentation
        """
        if not title or len(title.strip()) == 0:
            title = f"Copy of presentation"

        if len(title) > 1000:
            title = title[:1000]

        body = {'name': title.strip()}
        result = self._make_request(
            self.drive_service.files().copy,
            fileId=presentation_id,
            body=body
        )

        new_id = result['id']
        logger.info(f"Copied presentation {presentation_id} to {new_id}")
        return new_id

    def batch_update(self, presentation_id: str, requests: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Execute batch update with automatic chunking and size limits.

        Args:
            presentation_id: Presentation ID
            requests: List of requests to execute

        Returns:
            Results of all request executions
        """
        if not requests:
            return {'replies': []}

        # Validate total request count
        if len(requests) > self.slides_config.max_total_requests:
            raise SlidesAPIError(f"Too many requests: {len(requests)} > {self.slides_config.max_total_requests}")

        batch_size = self.slides_config.batch_size
        all_replies = []

        for i in range(0, len(requests), batch_size):
            batch = requests[i:i + batch_size]
            body = {'requests': batch}

            try:
                result = self._make_request(
                    self.slides_service.presentations().batchUpdate,
                    presentationId=presentation_id,
                    body=body
                )
                all_replies.extend(result.get('replies', []))
                logger.debug(f"Processed batch {i // batch_size + 1} ({len(batch)} requests)")

            except Exception as e:
                logger.error(f"Failed to process batch {i // batch_size + 1}: {e}")
                raise

        return {'replies': all_replies}

    def get_presentation_url(self, presentation_id: str) -> str:
        """
        Get presentation view URL.

        Args:
            presentation_id: Presentation ID

        Returns:
            Presentation URL
        """
        return f"https://docs.google.com/presentation/d/{presentation_id}/edit"

    def calculate_layout(self, components: List[Dict[str, Any]]) -> List[ElementPosition]:
        """
        Calculate layout positions for slide components.

        Args:
            components: List of slide components

        Returns:
            List of ElementPosition objects
        """
        positions = []
        current_y = self.layout_config.margin_y

        for component in components:
            component_type = component.get('type', 'text')
            content = component.get('content', '')

            if component_type == 'header':
                height = self.layout_config.header_height
                width = self.layout_config.default_width
            elif component_type == 'text':
                # Estimate height based on content length
                lines = max(1, len(content) // 80)  # Approx 80 chars per line
                height = max(50, lines * self.layout_config.text_line_height + 20)  # Убрано ограничение max
                width = self.layout_config.default_width
            elif component_type == 'image':
                height = 200  # Default image height
                width = self.layout_config.default_width
            else:
                height = self.layout_config.default_height
                width = self.layout_config.default_width

            # Check if we exceed slide height
            if current_y + height > self.layout_config.slide_height - self.layout_config.margin_y:
                # Could implement new slide logic here
                current_y = self.layout_config.margin_y

            position = ElementPosition(
                x=self.layout_config.margin_x,
                y=current_y,
                width=width,
                height=height,
                layer=0
            )
            positions.append(position)
            current_y += height + 20  # 20px spacing between elements

        return positions

    def add_markdown_slide(self, presentation_id: str, markdown_content: str,
                           slide_index: Optional[int] = None) -> str:
        """
        Add new slide with Markdown content using calculated layout.

        Args:
            presentation_id: Presentation ID
            markdown_content: Markdown text to add
            slide_index: Slide position (None = at end)

        Returns:
            Created slide ID
        """
        slide_id = f"slide_{uuid.uuid4().hex[:8]}"

        create_slide_request = {
            "createSlide": {
                "objectId": slide_id,
                "slideLayoutReference": {
                    "predefinedLayout": "BLANK"
                }
            }
        }

        if slide_index is not None:
            create_slide_request["createSlide"]["insertionIndex"] = slide_index

        self.batch_update(presentation_id, [create_slide_request])

        # Parse components and calculate layout
        components = self.markdown_processor.parse_to_components(markdown_content)
        positions = self.calculate_layout(components)

        # Create requests for all elements
        requests = []
        for component, position in zip(components, positions):
            element_id = f"element_{uuid.uuid4().hex[:8]}"

            # Create text box
            create_request = {
                "createShape": {
                    "objectId": element_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": {"magnitude": position.width, "unit": "PT"},
                            "height": {"magnitude": position.height, "unit": "PT"}
                        },
                        "transform": {
                            "translateX": position.x,
                            "translateY": position.y,
                            "scaleX": 1.0,
                            "scaleY": 1.0,
                            "unit": "PT"
                        }
                    }
                }
            }
            requests.append(create_request)

            # Add text content if present
            if component.get('content', '').strip():
                text_requests = self.markdown_processor.create_slides_requests(
                    element_id,
                    component['content']
                )
                requests.extend(text_requests)

        if requests:
            self.batch_update(presentation_id, requests)

        logger.info(f"Added slide {slide_id} with {len(components)} components")
        return slide_id

    def create_presentation_from_markdown(self, markdown_content: str,
                                          title: str = "Presentation from Markdown") -> str:
        """
        Create entire presentation from Markdown content.
        Automatically splits into slides by level 1 headers.

        Args:
            markdown_content: Full Markdown content
            title: Presentation title

        Returns:
            Created presentation ID
        """
        presentation_id = self.create_presentation(title)

        slides_content = self._split_markdown_to_slides(markdown_content)

        # Remove default empty slide
        try:
            presentation = self.get_presentation(presentation_id)
            if presentation.get('slides'):
                first_slide_id = presentation['slides'][0]['objectId']
                delete_request = {"deleteObject": {"objectId": first_slide_id}}
                self.batch_update(presentation_id, [delete_request])
        except Exception as e:
            logger.warning(f"Could not remove default slide: {e}")

        # Add slides with content
        for slide_content in slides_content:
            if slide_content.strip():
                try:
                    self.add_markdown_slide(presentation_id, slide_content)
                except Exception as e:
                    logger.error(f"Failed to add slide: {e}")
                    # Continue with other slides

        logger.info(f"Created presentation {presentation_id} with {len(slides_content)} slides")
        return presentation_id

    def _split_markdown_to_slides(self, markdown_content: str) -> List[str]:
        """
        Split Markdown content into slides by level 1 headers.

        Args:
            markdown_content: Source Markdown

        Returns:
            List of content for each slide
        """
        lines = markdown_content.split('\n')
        slides = []
        current_slide = []

        for line in lines:
            # Check for level 1 header
            if line.strip().startswith('# ') and current_slide:
                slides.append('\n'.join(current_slide))
                current_slide = [line]
            else:
                current_slide.append(line)

        if current_slide:
            slides.append('\n'.join(current_slide))

        return slides

    def add_text_box(self, presentation_id: str, slide_id: str,
                     text: str = "", position: Optional[ElementPosition] = None) -> str:
        """
        Add text box with Markdown support and configurable position.

        Args:
            presentation_id: Presentation ID
            slide_id: Slide ID
            text: Text (supports Markdown)
            position: Element position and size

        Returns:
            Created element ID
        """
        element_id = f"textbox_{uuid.uuid4().hex[:8]}"

        # Use provided position or defaults
        if position is None:
            position = ElementPosition(
                x=self.layout_config.margin_x,
                y=self.layout_config.margin_y,
                width=self.layout_config.default_width,
                height=self.layout_config.default_height
            )

        create_request = {
            "createShape": {
                "objectId": element_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {
                        "width": {"magnitude": position.width, "unit": "PT"},
                        "height": {"magnitude": position.height, "unit": "PT"}
                    },
                    "transform": {
                        "translateX": position.x,
                        "translateY": position.y,
                        "scaleX": 1.0,
                        "scaleY": 1.0,
                        "unit": "PT"
                    }
                }
            }
        }

        requests = [create_request]

        if text.strip():
            try:
                text_requests = self.markdown_processor.create_slides_requests(element_id, text)
                requests.extend(text_requests)
            except Exception as e:
                logger.error(f"Failed to process markdown text: {e}")
                # Fallback to plain text
                requests.append({
                    "insertText": {
                        "objectId": element_id,
                        "insertionIndex": 0,
                        "text": text
                    }
                })

        self.batch_update(presentation_id, requests)
        return element_id

    def set_text(self, presentation_id: str, element_id: str, text: str):
        """
        Set element text with Markdown support.
        Uses cached presentation data to avoid redundant API calls.

        Args:
            presentation_id: Presentation ID
            element_id: Element ID
            text: New text (supports Markdown)
        """
        if not text.strip():
            return

        requests = []

        try:
            # Try to delete existing text first (may fail if no text exists)
            requests.append({
                "deleteText": {
                    "objectId": element_id,
                    "textRange": {"type": "ALL"}
                }
            })

            # Add new text
            text_requests = self.markdown_processor.create_slides_requests(element_id, text)
            requests.extend(text_requests)

            if requests:
                self.batch_update(presentation_id, requests)

        except Exception as e:
            logger.warning(f"Failed to set text for {element_id}: {e}")
            # Fallback: just try to insert text
            try:
                fallback_requests = [{
                    "insertText": {
                        "objectId": element_id,
                        "insertionIndex": 0,
                        "text": text
                    }
                }]
                self.batch_update(presentation_id, fallback_requests)
            except Exception as fallback_error:
                logger.error(f"Fallback text insertion failed for {element_id}: {fallback_error}")

    def replace_image(self, presentation_id: str, image_id: str, new_url: str):
        """
        Replace image by URL.

        Args:
            presentation_id: Presentation ID
            image_id: Image ID
            new_url: New image URL
        """
        if not new_url.startswith(('http://', 'https://')):
            raise ValueError("Image URL must start with http:// or https://")

        request = {
            "replaceImage": {
                "imageObjectId": image_id,
                "url": new_url
            }
        }
        self.batch_update(presentation_id, [request])

    def _sanitize_html_comments(self, text: str) -> str:
        """Sanitize HTML comments to prevent injection"""
        if not text:
            return text

        # Only allow specific safe tags
        allowed_tags = self.markdown_processor.config.allowed_html_tags

        def replace_unsafe_tag(match):
            tag_type = match.group(1)
            if tag_type in allowed_tags:
                return match.group(0)  # Keep safe tags
            else:
                return match.group(3)  # Remove unsafe tags, keep content

        # Replace unsafe HTML comments
        safe_text = re.sub(
            r'<!-- (\w+):([^>]+) -->([^<]*?)<!-- /\1 -->',
            replace_unsafe_tag,
            text
        )

        return safe_text

    def create_template(self, presentation_id: str, template_name: str, debug: bool = False) -> Dict[str, Any]:
        """
        Create template configuration from presentation with enhanced metadata.

        Args:
            presentation_id: Source presentation ID
            template_name: Template name
            debug: Enable debug output for position extraction

        Returns:
            Template configuration with placeholders and layout info
        """
        if not template_name or len(template_name.strip()) == 0:
            raise TemplateValidationError("Template name cannot be empty")

        logger.info(f"Creating template '{template_name}' from presentation {presentation_id}")
        presentation = self.get_presentation(presentation_id)

        template = {
            'name': template_name.strip(),
            'source_presentation_id': presentation_id,
            'title': presentation.get('title', 'Untitled'),
            'created_at': self._get_timestamp(),
            'coordinate_system': {
                'units': 'EMU',
                'description': 'English Metric Units (1 Point = 12700 EMU)',
                'origin': 'Top-left corner of slide'
            },
            'slide_size': {
                'width': self.layout_config.slide_width,
                'height': self.layout_config.slide_height,
                'units': 'Points'
            },
            'layout_config': {
                'slide_width': self.layout_config.slide_width,
                'slide_height': self.layout_config.slide_height,
                'margin_x': self.layout_config.margin_x,
                'margin_y': self.layout_config.margin_y,
                'default_width': self.layout_config.default_width,
                'default_height': self.layout_config.default_height,
                'units': 'Points'
            },
            'slides': [],
            'placeholders': {}
        }

        placeholder_counter = 1
        total_elements_processed = 0

        for slide_idx, slide in enumerate(presentation.get('slides', [])):
            slide_info = {
                'slide_id': slide['objectId'],
                'slide_index': slide_idx,
                'replaceable_elements': []
            }

            # Get all page elements to determine layer order
            page_elements = slide.get('pageElements', [])
            logger.debug(f"Processing slide {slide_idx + 1} with {len(page_elements)} elements")

            for element_idx, element in enumerate(page_elements):
                total_elements_processed += 1
                element_id = element.get('objectId', f'unknown_{element_idx}')

                logger.debug(f"Checking element {element_id} (type: {self._get_element_type(element)})")

                if self._is_replaceable_element(element):
                    logger.debug(f"Element {element_id} is replaceable, processing...")

                    placeholder_name = self._generate_placeholder_name(
                        element, slide_idx, placeholder_counter
                    )

                    original_content = self._extract_content(element)
                    element_position = self._extract_element_position(element, debug)

                    # Set layer based on element order (later elements have higher layer)
                    if element_position:
                        element_position.layer = element_idx
                        logger.debug(f"Element {element_id} position: x={element_position.x}, y={element_position.y}, "
                                     f"w={element_position.width}, h={element_position.height}, layer={element_position.layer}")
                    else:
                        logger.warning(f"Could not extract position for element {element_id}")

                    # Convert to markdown if possible
                    markdown_content = self._convert_element_to_markdown(element, original_content)

                    element_info = {
                        'element_id': element_id,
                        'placeholder_name': placeholder_name,
                        'element_type': self._get_element_type(element),
                        'position': element_position.dict() if element_position else None,
                        'original_content': original_content,
                        'markdown_content': markdown_content
                    }

                    slide_info['replaceable_elements'].append(element_info)

                    template['placeholders'][placeholder_name] = {
                        'type': element_info['element_type'],
                        'slide_index': slide_idx,
                        'position': element_position.dict() if element_position else None,
                        'layer': element_position.layer if element_position else element_idx,
                        'position_units': 'EMU',  # Google Slides uses EMU (English Metric Units)
                        'description': self._get_placeholder_description(element_info['element_type']),
                        'example': markdown_content,
                        'original_example': original_content
                    }

                    placeholder_counter += 1
                    logger.info(f"Added placeholder '{placeholder_name}' for element {element_id}")
                else:
                    logger.debug(f"Element {element_id} is not replaceable, skipping")

            if slide_info['replaceable_elements']:
                template['slides'].append(slide_info)
                logger.info(
                    f"Slide {slide_idx + 1}: found {len(slide_info['replaceable_elements'])} replaceable elements")

        logger.info(
            f"Template creation completed: {len(template['placeholders'])} placeholders from {total_elements_processed} total elements")
        return template

    def _debug_element_structure(self, element: Dict[str, Any]) -> str:
        """Debug helper to understand element structure"""
        element_id = element.get('objectId', 'unknown')
        element_keys = list(element.keys())

        debug_info = f"Element {element_id} keys: {element_keys}"

        if 'elementProperties' in element:
            props = element['elementProperties']
            debug_info += f", elementProperties keys: {list(props.keys())}"

            if 'size' in props:
                size = props['size']
                debug_info += f", size keys: {list(size.keys()) if isinstance(size, dict) else type(size)}"

            if 'transform' in props:
                transform = props['transform']
                debug_info += f", transform keys: {list(transform.keys()) if isinstance(transform, dict) else type(transform)}"

        return debug_info

    def _extract_element_position(self, element: Dict[str, Any], debug: bool = False) -> Optional[ElementPosition]:
        """Extract position and size information from element"""
        element_id = element.get('objectId', 'unknown')

        try:
            # Get transform and size data directly from element
            transform_data = element.get('transform')
            size_data = element.get('size')

            if not transform_data or not size_data:
                return None

            # Extract raw values (in EMU units)
            x = transform_data.get('translateX', 0)
            y = transform_data.get('translateY', 0)

            width_data = size_data.get('width', {})
            height_data = size_data.get('height', {})

            if isinstance(width_data, dict):
                width_magnitude = width_data.get('magnitude', 0)
            else:
                width_magnitude = width_data

            if isinstance(height_data, dict):
                height_magnitude = height_data.get('magnitude', 0)
            else:
                height_magnitude = height_data

            # Validate that we have valid numbers
            if not all(isinstance(val, (int, float)) for val in [x, y, width_magnitude, height_magnitude]):
                return None

            # Store values in EMU (raw from API)
            position = ElementPosition(
                x=float(x),
                y=float(y),
                width=float(width_magnitude),
                height=float(height_magnitude),
                layer=0
            )

            return position

        except Exception as e:
            return None

    def apply_template(self, template: Dict[str, Any], data: Dict[str, Any],
                       title: str = None) -> str:
        """
        Apply template with data substitution using batch processing.

        Args:
            template: Template configuration
            data: Data for substitution (supports Markdown)
            title: New presentation title

        Returns:
            Created presentation ID
        """
        # Validate template
        if not template.get('source_presentation_id'):
            raise TemplateValidationError("Template missing source_presentation_id")

        validation_result = self.validate_template_data(template, data)
        if not validation_result['valid']:
            missing = validation_result.get('missing_placeholders', [])
            invalid = validation_result.get('invalid_types', [])
            error_msg = f"Template validation failed. Missing: {missing}, Invalid: {invalid}"
            raise TemplateValidationError(error_msg)

        new_title = title or f"Copy of {template['title']} - {self._get_timestamp()}"
        new_presentation_id = self.copy_presentation(
            template['source_presentation_id'],
            new_title
        )

        # Collect all update requests for batch processing
        all_requests = []

        for slide_info in template['slides']:
            for element_info in slide_info['replaceable_elements']:
                placeholder_name = element_info['placeholder_name']

                if placeholder_name in data:
                    value = data[placeholder_name]
                    requests = self._prepare_element_update_requests(element_info, value)
                    all_requests.extend(requests)

        # Process all updates in batches
        if all_requests:
            try:
                self.batch_update(new_presentation_id, all_requests)
                logger.info(f"Applied template with {len(all_requests)} updates")
            except Exception as e:
                logger.error(f"Failed to apply template updates: {e}")
                raise

        return new_presentation_id

    def _prepare_element_update_requests(self, element_info: Dict[str, Any], value: Any) -> List[Dict[str, Any]]:
        """
        Prepare update requests for a single element with improved image validation.

        Args:
            element_info: Element information from template
            value: New value to set

        Returns:
            List of update requests
        """
        requests = []
        element_type = element_info['element_type']
        element_id = element_info['element_id']

        try:
            if element_type == 'text' and isinstance(value, str):
                # Clear existing text
                requests.append({
                    "deleteText": {
                        "objectId": element_id,
                        "textRange": {"type": "ALL"}
                    }
                })

                # Add new text with markdown support
                text_requests = self.markdown_processor.create_slides_requests(element_id, value)
                requests.extend(text_requests)

            elif element_type == 'image' and isinstance(value, str):
                if value.startswith(('http://', 'https://')):
                    # Validate image URL before adding to requests
                    if self._validate_image_url(value):
                        requests.append({
                            "replaceImage": {
                                "imageObjectId": element_id,
                                "url": value
                            }
                        })
                    else:
                        logger.warning(f"Skipping invalid/inaccessible image URL: {value}")

        except Exception as e:
            logger.warning(f"Could not prepare update for element {element_id}: {e}")

        return requests

    def _validate_image_url(self, url: str) -> bool:
        """
        Validate that an image URL is accessible and valid.

        Args:
            url: Image URL to validate

        Returns:
            True if URL appears to be valid and accessible
        """
        import urllib.request
        import urllib.error

        if not url or not url.startswith(('http://', 'https://')):
            return False

        # Check for common image extensions
        valid_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp')
        url_lower = url.lower()

        # Allow URLs with parameters that might contain image extensions
        if not any(ext in url_lower for ext in valid_extensions):
            # If no obvious image extension, try a quick HEAD request
            try:
                req = urllib.request.Request(url, method='HEAD')
                req.add_header('User-Agent', 'Mozilla/5.0 (compatible; Google-Slides-Templater/1.0)')

                with urllib.request.urlopen(req, timeout=5) as response:
                    content_type = response.headers.get('Content-Type', '')
                    return content_type.startswith('image/')

            except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError):
                logger.warning(f"Could not validate image URL: {url}")
                return False

        return True

    def replace_image(self, presentation_id: str, image_id: str, new_url: str):
        """
        Replace image by URL with validation.

        Args:
            presentation_id: Presentation ID
            image_id: Image ID
            new_url: New image URL
        """
        if not new_url.startswith(('http://', 'https://')):
            raise ValueError("Image URL must start with http:// or https://")

        # Validate URL before attempting replacement
        if not self._validate_image_url(new_url):
            raise ValueError(f"Image URL is not accessible or invalid: {new_url}")

        request = {
            "replaceImage": {
                "imageObjectId": image_id,
                "url": new_url
            }
        }
        self.batch_update(presentation_id, [request])

    def apply_template(self, template: Dict[str, Any], data: Dict[str, Any],
                       title: str = None) -> str:
        """
        Apply template with data substitution using batch processing with improved error handling.

        Args:
            template: Template configuration
            data: Data for substitution (supports Markdown)
            title: New presentation title

        Returns:
            Created presentation ID
        """
        # Validate template
        if not template.get('source_presentation_id'):
            raise TemplateValidationError("Template missing source_presentation_id")

        validation_result = self.validate_template_data(template, data)
        if not validation_result['valid']:
            missing = validation_result.get('missing_placeholders', [])
            invalid = validation_result.get('invalid_types', [])
            error_msg = f"Template validation failed. Missing: {missing}, Invalid: {invalid}"
            raise TemplateValidationError(error_msg)

        new_title = title or f"Copy of {template['title']} - {self._get_timestamp()}"
        new_presentation_id = self.copy_presentation(
            template['source_presentation_id'],
            new_title
        )

        # Separate text and image updates for better error handling
        text_requests = []
        image_requests = []

        for slide_info in template['slides']:
            for element_info in slide_info['replaceable_elements']:
                placeholder_name = element_info['placeholder_name']

                if placeholder_name in data:
                    value = data[placeholder_name]
                    element_type = element_info['element_type']

                    if element_type == 'text':
                        requests = self._prepare_element_update_requests(element_info, value)
                        text_requests.extend(requests)
                    elif element_type == 'image':
                        requests = self._prepare_element_update_requests(element_info, value)
                        image_requests.extend(requests)

        # Process text updates first (they're more reliable)
        if text_requests:
            try:
                self.batch_update(new_presentation_id, text_requests)
                logger.info(f"Applied {len(text_requests)} text updates")
            except Exception as e:
                logger.error(f"Failed to apply text updates: {e}")
                raise

        # Process image updates with individual error handling
        successful_image_updates = 0
        failed_image_updates = 0

        for request in image_requests:
            try:
                self.batch_update(new_presentation_id, [request])
                successful_image_updates += 1
            except Exception as e:
                failed_image_updates += 1
                image_url = request.get('replaceImage', {}).get('url', 'unknown')
                logger.warning(f"Failed to update image {image_url}: {e}")

        logger.info(f"Image updates: {successful_image_updates} successful, {failed_image_updates} failed")

        if successful_image_updates == 0 and len(image_requests) > 0:
            logger.warning("All image updates failed - check image URLs are accessible")

        return new_presentation_id

    def _is_replaceable_element(self, element: Dict[str, Any]) -> bool:
        """Determine if element is replaceable."""
        if 'image' in element:
            return True

        if 'shape' in element:
            shape = element['shape']
            if 'text' in shape:
                text_elements = shape['text'].get('textElements', [])
                for te in text_elements:
                    if 'textRun' in te:
                        content = te['textRun'].get('content', '').strip()
                        if content and len(content) > 1:
                            return True

        if 'table' in element:
            return True

        return False

    def _get_element_type(self, element: Dict[str, Any]) -> str:
        """Determine element type."""
        if 'image' in element:
            return 'image'
        elif 'shape' in element:
            return 'text'
        elif 'table' in element:
            return 'table'
        elif 'video' in element:
            return 'video'
        else:
            return 'unknown'

    def _extract_content(self, element: Dict[str, Any]) -> str:
        """Extract content from element."""
        element_type = self._get_element_type(element)

        if element_type == 'text':
            return self._extract_text_content(element)
        elif element_type == 'image':
            image_data = element.get('image', {})
            return image_data.get('sourceUrl', 'https://example.com/image.jpg')
        else:
            return f"Sample {element_type} content"

    def _extract_text_content(self, element: Dict[str, Any]) -> str:
        """Extract text content from text element."""
        if 'shape' not in element:
            return ""

        shape = element['shape']
        if 'text' not in shape:
            return ""

        text_data = shape['text']
        text_elements = text_data.get('textElements', [])

        result = ""
        for te in text_elements:
            if 'textRun' in te:
                content = te['textRun'].get('content', '')
                result += content

        return result.strip()

    def _convert_element_to_markdown(self, element: Dict[str, Any], content: str) -> str:
        """
        Convert Google Slides element to Markdown format.

        Args:
            element: Element from Google Slides API
            content: Extracted content

        Returns:
            Content in Markdown format
        """
        element_type = self._get_element_type(element)

        if element_type == 'text':
            if 'shape' in element and 'text' in element['shape']:
                text_elements = element['shape']['text'].get('textElements', [])
                if text_elements:
                    return self.markdown_processor.slides_elements_to_markdown(text_elements)

        elif element_type == 'image':
            image_data = element.get('image', {})
            url = image_data.get('sourceUrl', 'https://example.com/image.jpg')
            description = image_data.get('title', 'Image')
            return f"![{description}]({url})"

        return content

    def _generate_placeholder_name(self, element: Dict[str, Any],
                                   slide_idx: int, counter: int) -> str:
        """
        Generate meaningful placeholder name.

        Args:
            element: Presentation element
            slide_idx: Slide index
            counter: Element counter

        Returns:
            Placeholder name
        """
        element_type = self._get_element_type(element)

        if element_type == 'text':
            text_content = self._extract_text_content(element)
            if text_content:
                # Clean text and extract meaningful words - берем больше слов для лучших имен
                clean_text = re.sub(r'[#*`~\[\]()]+', '', text_content)
                clean_text = re.sub(r'<[^>]+>', '', clean_text)
                words = re.findall(r'\w+', clean_text.lower())[:4]  # Увеличено до 4 слов
                if words and len(words[0]) > 2:
                    return '_'.join(words)

        return f"slide_{slide_idx + 1}_{element_type}_{counter}"

    def _get_placeholder_description(self, element_type: str) -> str:
        """Get placeholder description by element type."""
        descriptions = {
            'text': 'Text content (supports Markdown formatting)',
            'image': 'Image URL (must be publicly accessible)',
            'table': 'Table data (list of rows or dictionary)',
            'video': 'Video URL'
        }
        return descriptions.get(element_type, f'{element_type.title()} type content')

    def _get_timestamp(self) -> str:
        """Get current timestamp."""
        return time.strftime('%Y-%m-%d %H:%M:%S')

    def save_template(self, template: Dict[str, Any], filename: str):
        """
        Save template to JSON file with path validation.

        Args:
            template: Template configuration
            filename: File name to save
        """
        if not _is_safe_path(filename):
            raise SlidesAPIError("Unsafe file path")

        safe_path = Path(filename).resolve()
        safe_path.parent.mkdir(parents=True, exist_ok=True)

        try:
            with open(safe_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            logger.info(f"Template saved to {safe_path}")
        except Exception as e:
            logger.error(f"Failed to save template to {safe_path}: {e}")
            raise SlidesAPIError(f"Failed to save template: {e}")

    def load_template(self, filename: str) -> Dict[str, Any]:
        """
        Load template from JSON file with path validation.

        Args:
            filename: Template file name

        Returns:
            Template configuration
        """
        if not _is_safe_path(filename):
            raise SlidesAPIError("Unsafe file path")

        safe_path = Path(filename).resolve()

        try:
            with open(safe_path, 'r', encoding='utf-8') as f:
                template = json.load(f)
            logger.info(f"Template loaded from {safe_path}")
            return template
        except FileNotFoundError:
            raise SlidesAPIError(f"Template file not found: {safe_path}")
        except json.JSONDecodeError as e:
            raise SlidesAPIError(f"Invalid JSON in template file: {e}")
        except Exception as e:
            logger.error(f"Failed to load template from {safe_path}: {e}")
            raise SlidesAPIError(f"Failed to load template: {e}")

    def get_template_info(self, template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get detailed template information.

        Args:
            template: Template configuration

        Returns:
            Template information
        """
        placeholders = template.get('placeholders', {})
        slides = template.get('slides', [])

        element_types = {}
        for placeholder_info in placeholders.values():
            elem_type = placeholder_info['type']
            element_types[elem_type] = element_types.get(elem_type, 0) + 1

        return {
            'name': template.get('name', 'Unnamed'),
            'title': template.get('title', 'Untitled'),
            'created_at': template.get('created_at', 'Unknown'),
            'source_presentation_id': template.get('source_presentation_id'),
            'slide_size': template.get('slide_size', {}),
            'total_slides': len(slides),
            'total_placeholders': len(placeholders),
            'element_types': element_types,
            'placeholder_names': list(placeholders.keys())
        }

    def validate_template_data(self, template: Dict[str, Any], data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Validate data for template application.

        Args:
            template: Template configuration
            data: Data to check

        Returns:
            Validation result
        """
        placeholders = template.get('placeholders', {})

        validation_result = {
            'valid': True,
            'missing_placeholders': [],
            'extra_data': [],
            'invalid_types': [],
            'warnings': []
        }

        # Check for missing placeholders
        for placeholder_name, placeholder_info in placeholders.items():
            if placeholder_name not in data:
                validation_result['missing_placeholders'].append(placeholder_name)
                validation_result['valid'] = False

        # Check for extra data
        for data_key in data.keys():
            if data_key not in placeholders:
                validation_result['extra_data'].append(data_key)
                validation_result['warnings'].append(f"Data '{data_key}' does not match any placeholder")

        # Validate data types
        for placeholder_name, value in data.items():
            if placeholder_name in placeholders:
                expected_type = placeholders[placeholder_name]['type']

                if expected_type == 'text' and not isinstance(value, str):
                    validation_result['invalid_types'].append(
                        f"{placeholder_name}: expected text, got {type(value).__name__}")
                elif expected_type == 'image' and not isinstance(value, str):
                    validation_result['invalid_types'].append(
                        f"{placeholder_name}: expected image URL, got {type(value).__name__}")
                elif expected_type == 'image' and isinstance(value, str) and not value.startswith(
                        ('http://', 'https://')):
                    validation_result['warnings'].append(
                        f"{placeholder_name}: image URL should start with http:// or https://")

        if validation_result['invalid_types']:
            validation_result['valid'] = False

        return validation_result

    def preview_template_application(self, template: Dict[str, Any], data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Preview template application.

        Args:
            template: Template configuration
            data: Data to apply

        Returns:
            Preview of changes
        """
        placeholders = template.get('placeholders', {})
        preview = {
            'changes': [],
            'unchanged_placeholders': []
        }

        for placeholder_name, placeholder_info in placeholders.items():
            slide_index = placeholder_info.get('slide_index', 0) + 1
            element_type = placeholder_info.get('type', 'unknown')
            position = placeholder_info.get('position', {})

            if placeholder_name in data:
                old_value = placeholder_info.get('example', '')
                new_value = data[placeholder_name]

                preview['changes'].append({
                    'placeholder': placeholder_name,
                    'slide': slide_index,
                    'type': element_type,
                    'position': position,
                    'layer': placeholder_info.get('layer', 0),
                    'old_value': str(old_value),
                    'new_value': str(new_value)
                })
            else:
                preview['unchanged_placeholders'].append({
                    'placeholder': placeholder_name,
                    'slide': slide_index,
                    'type': element_type,
                    'position': position,
                    'layer': placeholder_info.get('layer', 0),
                    'current_value': placeholder_info.get('example', '')
                })

        return preview

    def clone_presentation(self, presentation_id: str, title: str = None) -> str:
        """
        Clone presentation (alias for copy_presentation).

        Args:
            presentation_id: Source presentation ID
            title: New presentation title

        Returns:
            Cloned presentation ID
        """
        if title is None:
            original = self.get_presentation(presentation_id)
            title = f"Copy of {original.get('title', 'Untitled')}"

        return self.copy_presentation(presentation_id, title)

    def get_presentation_info(self, presentation_id: str) -> Dict[str, Any]:
        """
        Get brief presentation information.

        Args:
            presentation_id: Presentation ID

        Returns:
            Presentation information
        """
        presentation = self.get_presentation(presentation_id)
        slides = presentation.get('slides', [])

        total_elements = 0
        element_types = {}

        for slide in slides:
            for element in slide.get('pageElements', []):
                total_elements += 1
                elem_type = self._get_element_type(element)
                element_types[elem_type] = element_types.get(elem_type, 0) + 1

        return {
            'id': presentation_id,
            'title': presentation.get('title', 'Untitled'),
            'total_slides': len(slides),
            'total_elements': total_elements,
            'element_types': element_types,
            'url': self.get_presentation_url(presentation_id)
        }

    def create_sample_data(self, template: Dict[str, Any]) -> Dict[str, Any]:
        """
        Create sample data for template.

        Args:
            template: Template configuration

        Returns:
            Dictionary with sample data
        """
        placeholders = template.get('placeholders', {})
        sample_data = {}

        for placeholder_name, placeholder_info in placeholders.items():
            element_type = placeholder_info['type']

            if element_type == 'text':
                sample_data[placeholder_name] = f"""# New header for {placeholder_name}

## Subheader

This is **updated content** with various formatting:

- List with *italic*
- Item with `code`
- ~~Strikethrough~~ corrected text

### Additional information

***Bold italic*** and regular text.

> Quote with important information"""

            elif element_type == 'image':
                sample_data[placeholder_name] = "https://via.placeholder.com/600x400/4285f4/ffffff?text=Sample+Image"

            else:
                sample_data[placeholder_name] = f"Sample data for {element_type}"

        return sample_data


def create_templater(auth_config: Optional[AuthConfig] = None, **kwargs) -> SlidesTemplater:
    """
    Create SlidesTemplater instance with configuration.

    Args:
        auth_config: Authentication configuration
        **kwargs: Additional configuration parameters

    Returns:
        Configured SlidesTemplater
    """
    return SlidesTemplater.from_credentials(auth_config, **kwargs)
