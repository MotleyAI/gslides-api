import logging
from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any, Literal, Optional, Union

if TYPE_CHECKING:
    pass  # Forward reference types handled via string annotations

from pydantic import BaseModel, Field, PrivateAttr, model_validator

from gslides_api.agnostic.domain import ImageData
from gslides_api.agnostic.element import MarkdownTableElement
from gslides_api.agnostic.units import OutputUnit

from gslides_api.agnostic.element_size import ElementSizeMeta

logger = logging.getLogger(__name__)

if TYPE_CHECKING:
    pass


def _extract_font_size_pt(styles: list[Any] | None) -> float:
    """Extract the dominant font size (in points) from a shape element's text styles.

    Handles both GSlides (RichStyle with font_size_pt) and PPTX
    (StyleInfo dict with font_size that has a .pt attribute).

    Returns:
        Font size in points, or 12.0 as fallback.
    """
    if not styles:
        return 12.0

    font_sizes = []
    for style in styles:
        if isinstance(style, dict):
            # PPTX StyleInfo dict
            fs = style.get("font_size")
            if fs is not None and hasattr(fs, "pt"):
                font_sizes.append(fs.pt)
        else:
            # GSlides RichStyle object
            if hasattr(style, "font_size_pt") and style.font_size_pt is not None:
                font_sizes.append(style.font_size_pt)

    return max(font_sizes) if font_sizes else 12.0


def _extract_font_size_from_table(element: "AbstractTableElement") -> float:
    """Extract the dominant font size (in points) from a table element's first cell.

    Handles both GSlides (TableElement with tableRows) and PPTX
    (GraphicFrame with .table accessor).

    Returns:
        Font size in points, or 10.0 as fallback.
    """
    try:
        if hasattr(element, "pptx_element") and element.pptx_element is not None:
            # PPTX path
            table = element.pptx_element.table
            cell = table.cell(0, 0)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size is not None and hasattr(run.font.size, "pt"):
                        return run.font.size.pt
        elif hasattr(element, "gslides_element") and element.gslides_element is not None:
            # GSlides path — access table rows from the underlying gslides-api element
            gslides_table = element.gslides_element
            if hasattr(gslides_table, "table") and gslides_table.table is not None:
                table_data = gslides_table.table
                if table_data.tableRows:
                    first_row = table_data.tableRows[0]
                    if first_row.tableCells:
                        cell = first_row.tableCells[0]
                        # Cell text content has styles
                        if hasattr(cell, "text") and hasattr(cell.text, "textElements"):
                            for te in cell.text.textElements:
                                if hasattr(te, "textRun") and te.textRun is not None:
                                    ts = te.textRun.style
                                    if ts and hasattr(ts, "fontSize") and ts.fontSize:
                                        return ts.fontSize.magnitude
    except Exception:
        pass
    return 10.0


# Supporting data classes
class AbstractThumbnail(BaseModel):
    contentUrl: str
    width: int
    height: int
    mime_type: str
    content: Optional[bytes] = None
    file_size: Optional[int] = None


class AbstractSlideProperties(BaseModel):
    isSkipped: bool = False


class AbstractAltText(BaseModel):
    title: str | None = None
    description: str | None = None


class AbstractSpeakerNotes(BaseModel, ABC):
    @abstractmethod
    def read_text(self, as_markdown: bool = True) -> str:
        pass

    @abstractmethod
    def write_text(self, api_client: "AbstractSlidesAPIClient", content: str):
        pass


class AbstractCredentials(BaseModel):
    token: Optional[str] = None
    refresh_token: Optional[str] = None
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    token_uri: Optional[str] = None


class AbstractSize(BaseModel):
    width: float = 0.0
    height: float = 0.0


class AbstractPreprocessedSlide(BaseModel):
    gslide: "AbstractSlide"
    raw_metadata: str = ""
    metadata: Optional[list[dict]] = None
    named_elements: dict[str, "AbstractElement"] = Field(default_factory=dict)


# Enums
class AbstractThumbnailSize:
    MEDIUM = "MEDIUM"


class AbstractElementKind:
    IMAGE = "IMAGE"
    SHAPE = "SHAPE"
    TABLE = "TABLE"


# Core abstract classes
class AbstractSlidesAPIClient(ABC):
    auto_flush: bool = True

    # TODO: remembering to call this is a chore, should we make this into a context manager?
    @abstractmethod
    def flush_batch_update(self):
        pass

    @abstractmethod
    def copy_presentation(
        self, presentation_id: str, copy_title: str, folder_id: Optional[str] = None
    ) -> dict:
        pass

    @abstractmethod
    def create_folder(
        self, name: str, ignore_existing: bool = True, parent_folder_id: Optional[str] = None
    ) -> dict:
        pass

    @abstractmethod
    def delete_file(self, file_id: str):
        pass

    def trash_file(self, file_id: str):
        """Move a file to trash. Defaults to delete_file for non-GSlides adapters."""
        self.delete_file(file_id)

    @abstractmethod
    def set_credentials(self, credentials: AbstractCredentials):
        pass

    @abstractmethod
    def get_credentials(self) -> Optional[AbstractCredentials]:
        pass

    @abstractmethod
    def replace_text(
        self, slide_ids: list[str], match_text: str, replace_text: str, presentation_id: str
    ):
        pass

    @classmethod
    def get_default_api_client(cls) -> "AbstractSlidesAPIClient":
        """Get the default API client wrapped in concrete implementation."""
        # TODO: This is a horrible, non-generalizable hack, will need to fix later
        from gslides_api.adapters.gslides_adapter import GSlidesAPIClient

        return GSlidesAPIClient.get_default_api_client()

    @abstractmethod
    def get_presentation_as_pdf(self, presentation_id: str) -> bytes:
        pass


class AbstractElement(BaseModel, ABC):
    objectId: str = ""
    presentation_id: str = ""
    slide_id: str = ""
    alt_text: AbstractAltText = Field(default_factory=AbstractAltText)
    type: str = ""

    # Parent references - not serialized, populated by parent validators
    _parent_slide: Optional["AbstractSlide"] = PrivateAttr(default=None)
    _parent_presentation: Optional["AbstractPresentation"] = PrivateAttr(default=None)

    def __eq__(self, other: object) -> bool:
        """Custom equality that excludes parent references to avoid circular comparison."""
        if not isinstance(other, AbstractElement):
            return False
        # Compare only the public fields, not parent references
        return self.model_dump() == other.model_dump()

    def __hash__(self) -> int:
        """Hash based on objectId for use in sets/dicts."""
        return hash(self.objectId)

    @abstractmethod
    def absolute_size(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        pass

    # @abstractmethod
    # def element_properties(self) -> dict:
    #     pass

    @abstractmethod
    def absolute_position(self, units: OutputUnit = OutputUnit.IN) -> tuple[float, float]:
        pass

    @abstractmethod
    def create_image_element_like(
        self, api_client: AbstractSlidesAPIClient
    ) -> "AbstractImageElement":
        pass

    @abstractmethod
    def set_alt_text(
        self,
        api_client: AbstractSlidesAPIClient,
        title: str | None = None,
        description: str | None = None,
    ):
        pass


class AbstractShapeElement(AbstractElement):
    type: Literal[AbstractElementKind.SHAPE] = AbstractElementKind.SHAPE

    @property
    @abstractmethod
    def has_text(self) -> bool:
        pass

    @abstractmethod
    def write_text(
        self, api_client: AbstractSlidesAPIClient, content: str, autoscale: bool = False
    ):
        pass

    @abstractmethod
    def read_text(self, as_markdown: bool = True) -> str:
        pass

    @abstractmethod
    def styles(self, skip_whitespace: bool = True) -> list[Any] | None:
        pass


class AbstractImageElement(AbstractElement):
    type: Literal[AbstractElementKind.IMAGE] = AbstractElementKind.IMAGE

    # @abstractmethod
    # def replace_image(self, url: str, api_client: Optional[AbstractSlidesAPIClient] = None):
    #     pass

    @abstractmethod
    def replace_image(
        self,
        api_client: AbstractSlidesAPIClient,
        file: str | None = None,
        url: str | None = None,
    ):
        pass


class AbstractTableElement(AbstractElement):
    type: Literal[AbstractElementKind.TABLE] = AbstractElementKind.TABLE

    @abstractmethod
    def resize(
        self,
        api_client: AbstractSlidesAPIClient,
        rows: int,
        cols: int,
        fix_width: bool = True,
        fix_height: bool = True,
        target_height_in: float | None = None,
    ) -> float:
        """Resize the table.

        Returns:
            Font scale factor (1.0 if no scaling, < 1.0 if rows were added with fix_height)
        """
        pass

    @abstractmethod
    def update_content(
        self,
        api_client: AbstractSlidesAPIClient,
        markdown_content: MarkdownTableElement,
        check_shape: bool = True,
        font_scale_factor: float = 1.0,
    ):
        pass

    @abstractmethod
    def to_markdown_element(self, name: str | None = None) -> MarkdownTableElement:
        pass

    @abstractmethod
    def get_horizontal_border_weight(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Get weight of a single horizontal border in specified units."""
        pass

    @abstractmethod
    def get_row_count(self) -> int:
        """Get current number of rows."""
        pass

    @abstractmethod
    def get_column_count(self) -> int:
        """Get current number of columns."""
        pass

    def get_total_height_including_borders(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Get total table height including borders.

        Returns:
            Total height: sum of row heights + all horizontal border heights.
        """
        _, row_heights_total = self.absolute_size(units=units)
        border_weight = self.get_horizontal_border_weight(units=units)
        num_borders = self.get_row_count() + 1
        return row_heights_total + (border_weight * num_borders)

    def get_max_height(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Calculate max allowed height based on elements below this table.

        Returns:
            Max height in specified units.

        Raises:
            RuntimeError: If parent references are not set (programming error).
        """
        if self._parent_slide is None or self._parent_presentation is None:
            raise RuntimeError(
                f"Element {self.objectId} missing parent references. "
                f"_parent_slide={self._parent_slide}, _parent_presentation={self._parent_presentation}. "
                "This is a programming error - parent references should be set during slide creation."
            )

        slide = self._parent_slide
        presentation = self._parent_presentation

        # Get this table's position and size
        table_x, table_top_y = self.absolute_position(units=units)
        table_width, table_height = self.absolute_size(units=units)
        table_bottom_y = table_top_y + table_height

        # Find minimum Y of elements below the table
        slide_height = presentation.slide_height(units=units)
        min_y_below = slide_height

        for element in slide.page_elements_flat:
            if element.objectId == self.objectId:
                continue

            elem_x, elem_y = element.absolute_position(units=units)

            # Element is "below" if its top is at or below table's bottom
            if elem_y >= table_bottom_y:
                min_y_below = min(min_y_below, elem_y)

        return min_y_below - table_top_y


class AbstractSlide(BaseModel, ABC):
    elements: list[AbstractElement] = Field(
        description="The elements of the slide", default_factory=list
    )
    objectId: str = ""
    slideProperties: AbstractSlideProperties = Field(default_factory=AbstractSlideProperties)
    speaker_notes: Optional[AbstractSpeakerNotes] = None

    # Parent reference for this slide
    _parent_presentation: Optional["AbstractPresentation"] = PrivateAttr(default=None)

    def __eq__(self, other: object) -> bool:
        """Custom equality that excludes parent references to avoid circular comparison."""
        if not isinstance(other, AbstractSlide):
            return False
        # Compare only the public fields, not parent references
        return self.model_dump() == other.model_dump()

    def __hash__(self) -> int:
        """Hash based on objectId for use in sets/dicts."""
        return hash(self.objectId)

    @model_validator(mode="after")
    def _populate_element_parent_refs(self) -> "AbstractSlide":
        """Populate parent references on elements after creation/deserialization."""
        for element in self.elements:
            element._parent_slide = self
            # _parent_presentation is set by presentation validator
        return self

    @property
    def name(self) -> str:
        """Get the slide name from the speaker notes."""
        if not self.speaker_notes:
            return ""
        try:
            return self.speaker_notes.read_text()
        except Exception:
            return ""

    @property
    def page_elements_flat(self) -> list[AbstractElement]:
        """Flatten the elements tree into a list."""
        return self.elements

    def markdown(self) -> str:
        """Return a markdown representation of this slide's layout and content.

        Metadata (element type, position, size, char capacity) is embedded in
        HTML comments following the gslides-api MarkdownSlideElement convention.
        Text and table content appears as regular markdown below each comment.
        """
        parts = []
        for element in self.page_elements_flat:
            name = element.alt_text.title or element.objectId
            x, y = element.absolute_position()
            w, h = element.absolute_size()

            if isinstance(element, AbstractShapeElement) and element.has_text:
                text = element.read_text(as_markdown=True)
                try:
                    font_pt = _extract_font_size_pt(element.styles(skip_whitespace=True))
                except Exception:
                    font_pt = 12.0
                meta = ElementSizeMeta(
                    box_width_inches=w, box_height_inches=h, font_size_pt=font_pt,
                )
                parts.append(
                    f"<!-- text: {name} | pos=({x:.1f},{y:.1f}) size=({w:.1f},{h:.1f}) "
                    f"| ~{meta.approx_char_capacity} chars -->\n{text}"
                )
            elif isinstance(element, AbstractTableElement):
                md_elem = element.to_markdown_element(name=name)
                table_md = md_elem.content.to_markdown() if md_elem and md_elem.content else ""
                _rows, cols = md_elem.shape if md_elem else (0, 0)
                # Estimate per-column char capacity (equal-width approximation)
                col_chars_str = ""
                if cols > 0:
                    font_pt = _extract_font_size_from_table(element)
                    col_width = w / cols
                    chars_per_col = int(col_width / (font_pt * 0.5 / 72))
                    col_chars_str = f" | ~{chars_per_col} chars/col"
                parts.append(
                    f"<!-- table: {name} | pos=({x:.1f},{y:.1f}) size=({w:.1f},{h:.1f})"
                    f"{col_chars_str} -->\n{table_md}"
                )
            elif isinstance(element, AbstractImageElement):
                parts.append(
                    f"<!-- image: {name} | pos=({x:.1f},{y:.1f}) size=({w:.1f},{h:.1f}) -->"
                )
            else:
                parts.append(
                    f"<!-- {element.type}: {name} | pos=({x:.1f},{y:.1f}) size=({w:.1f},{h:.1f}) -->"
                )

        return "\n\n".join(parts)

    @abstractmethod
    def thumbnail(
        self, api_client: AbstractSlidesAPIClient, size: str, include_data: bool = False
    ) -> AbstractThumbnail:
        pass

    def get_elements_by_alt_title(self, title: str) -> list[AbstractElement]:
        return [e for e in self.page_elements_flat if e.alt_text.title == title]

    def __getitem__(self, item: str):
        """Get element by alt title."""
        elements = self.get_elements_by_alt_title(item)
        if not elements:
            raise KeyError(f"Element with alt title {item} not found")
        if len(elements) > 1:
            raise KeyError(f"Multiple elements with alt title {item} found")
        return elements[0]


class AbstractPresentation(BaseModel, ABC):
    slides: list[AbstractSlide] = Field(default_factory=list)
    presentationId: str | None = None
    revisionId: str | None = None
    title: str | None = None

    @model_validator(mode="after")
    def _populate_slide_parent_refs(self) -> "AbstractPresentation":
        """Populate parent references on slides/elements after creation/deserialization."""
        for slide in self.slides:
            slide._parent_presentation = self
            for element in slide.elements:
                element._parent_presentation = self
        return self

    @property
    @abstractmethod
    def url(self) -> str:
        pass

    @abstractmethod
    def slide_height(self, units: OutputUnit = OutputUnit.IN) -> float:
        """Return slide height in specified units."""
        pass

    @abstractmethod
    def save(self, api_client: "AbstractSlidesAPIClient") -> None:
        """Save/persist all changes made to this presentation."""
        pass

    def __getitem__(self, item: str):
        """Get slide by name."""
        for slide in self.slides:
            if slide.name == item:
                return slide
        raise KeyError(f"Slide with name {item} not found")

    @classmethod
    @abstractmethod
    def from_id(
        cls, api_client: AbstractSlidesAPIClient, presentation_id: str
    ) -> "AbstractPresentation":
        from gslides_api.adapters.gslides_adapter import GSlidesAPIClient, GSlidesPresentation
        from gslides_api.adapters.html_adapter import HTMLAPIClient, HTMLPresentation
        from gslides_api.adapters.pptx_adapter import PowerPointAPIClient, PowerPointPresentation

        if isinstance(api_client, GSlidesAPIClient):
            return GSlidesPresentation.from_id(api_client, presentation_id)
        elif isinstance(api_client, PowerPointAPIClient):
            return PowerPointPresentation.from_id(api_client, presentation_id)
        elif isinstance(api_client, HTMLAPIClient):
            return HTMLPresentation.from_id(api_client, presentation_id)
        else:
            raise NotImplementedError("Only gslides, pptx, and html clients are supported")

    @abstractmethod
    def copy_via_drive(
        self,
        api_client: AbstractSlidesAPIClient,
        copy_title: str,
        folder_id: Optional[str] = None,
    ) -> "AbstractPresentation":
        pass

    @abstractmethod
    def sync_from_cloud(self, api_client: AbstractSlidesAPIClient):
        pass

    @abstractmethod
    def insert_copy(
        self,
        source_slide: "AbstractSlide",
        api_client: AbstractSlidesAPIClient,
        insertion_index: int | None = None,
    ) -> "AbstractSlide":
        pass

    @abstractmethod
    def delete_slide(self, slide: Union["AbstractSlide", int], api_client: AbstractSlidesAPIClient):
        """Delete a slide from the presentation by reference or index."""
        pass

    @abstractmethod
    def delete_slides(
        self, slides: list[Union["AbstractSlide", int]], api_client: AbstractSlidesAPIClient
    ):
        """Delete multiple slides from the presentation by reference or index."""
        pass

    @abstractmethod
    def move_slide(
        self,
        slide: Union["AbstractSlide", int],
        insertion_index: int,
        api_client: AbstractSlidesAPIClient,
    ):
        """Move a slide to a new position within the presentation."""
        pass

    @abstractmethod
    def duplicate_slide(
        self, slide: Union["AbstractSlide", int], api_client: AbstractSlidesAPIClient
    ) -> "AbstractSlide":
        """Duplicate a slide within the presentation."""
        pass

    async def get_slide_thumbnails(
        self,
        api_client: "AbstractSlidesAPIClient",
        slides: Optional[list["AbstractSlide"]] = None,
    ) -> list["AbstractThumbnail"]:
        """Get thumbnails for slides in this presentation.

        Default implementation loops through each slide's thumbnail() method.
        Subclasses can override for more efficient batch implementations
        (e.g., HTML using single Playwright session, PPTX using single PDF conversion).

        Args:
            api_client: The API client for slide operations
            slides: Optional list of slides to get thumbnails for. If None, uses all slides.

        Returns:
            List of AbstractThumbnail objects with image data
        """
        target_slides = slides if slides is not None else self.slides
        thumbnails = []

        for slide in target_slides:
            thumb = slide.thumbnail(
                api_client=api_client,
                size=AbstractThumbnailSize.MEDIUM,
                include_data=True,
            )

            # Ensure file_size is set if content is available
            if thumb.content and thumb.file_size is None:
                thumb = AbstractThumbnail(
                    contentUrl=thumb.contentUrl,
                    width=thumb.width,
                    height=thumb.height,
                    mime_type=thumb.mime_type,
                    content=thumb.content,
                    file_size=len(thumb.content),
                )

            thumbnails.append(thumb)

        return thumbnails


# class AbstractLayoutMatcher(ABC):
#     """Abstract matcher for finding slide layouts that match given criteria."""
#
#     @abstractmethod
#     def __init__(self, presentation: AbstractPresentation, matching_rule: Optional[str] = None):
#         pass
#
#     @abstractmethod
#     def match(self, layout, matching_rule: Optional[str] = None) -> list[AbstractPreprocessedSlide]:
#         pass
