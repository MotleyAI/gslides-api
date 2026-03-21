from dataclasses import dataclass
from typing import Optional
import logging

from langchain_core.language_models import BaseLanguageModel

from gslides_api.adapters.abstract_slides import (
    AbstractElement,
    AbstractImageElement,
    AbstractPresentation,
    AbstractShapeElement,
    AbstractSlide,
    AbstractSlidesAPIClient,
    AbstractTableElement,
    _extract_font_size_from_table,
    _extract_font_size_pt,
)
from gslides_api.agnostic.element_size import ElementSizeMeta
from gslides_api.agnostic.units import OutputUnit
from motleycrew.utils.image_utils import is_this_a_chart

logger = logging.getLogger(__name__)


@dataclass
class SlideElementNames:
    """Names of different types of elements in a slide."""

    image_names: list[str]
    text_names: list[str]
    chart_names: list[str]
    table_names: list[str]

    @classmethod
    def empty(cls) -> "SlideElementNames":
        """Create an empty SlideElementNames instance."""
        return cls(image_names=[], text_names=[], chart_names=[], table_names=[])


def name_slides(
    presentation_id: str,
    name_elements: bool = True,
    api_client: AbstractSlidesAPIClient | None = None,
    skip_empty_text_boxes: bool = False,
    llm: Optional[BaseLanguageModel] = None,
    check_success: bool = False,
    min_image_size_cm: float = 4.0,
) -> dict[str, SlideElementNames]:
    """
    Name slides in a presentation based on their speaker notes.
    If name_elements is True, also name the elements in the slides."""

    if api_client is None:
        raise ValueError("API client is required")
    """Name slides in a presentation based on their speaker notes., enforcing unique names"""
    # api_client = api_client or AbstractSlidesAPIClient.get_default_api_client()
    if llm is None:
        logger.warning("No LLM provided, will not attempt to distinguish charts from images")

    presentation = AbstractPresentation.from_id(
        api_client=api_client, presentation_id=presentation_id
    )
    slide_names = {}
    for i, slide in enumerate(presentation.slides):
        if slide.slideProperties.isSkipped:
            logger.info(
                f"Skipping slide {i+1}: {slide.objectId} as it is marked as skipped in Google Slides"
            )
            continue

        speaker_notes = slide.speaker_notes.read_text().strip()
        if speaker_notes:
            slide_name = speaker_notes.split("\n")[0]
            if slide_name in slide_names:
                slide_name = f"Slide_{i+1}"
        else:
            slide_name = f"Slide_{i+1}"

        logger.info(f"Naming slide {i+1}: {slide.objectId} as {slide_name}")

        slide.speaker_notes.write_text(api_client=api_client, content=slide_name)

        if name_elements:
            slide_names[slide_name] = name_slide_elements(
                slide,
                skip_empty_text_boxes=skip_empty_text_boxes,
                slide_name=slide_name,
                llm=llm,
                api_client=api_client,
                min_image_size_cm=min_image_size_cm,
            )
        else:
            # Just get the existing names
            text_names = [
                e.alt_text.title
                for e in slide.page_elements_flat
                if isinstance(e, AbstractShapeElement)
            ]
            image_names = [
                e.alt_text.title
                for e in slide.page_elements_flat
                if isinstance(e, AbstractImageElement)
            ]
            table_names = [
                e.alt_text.title
                for e in slide.page_elements_flat
                if isinstance(e, AbstractTableElement)
            ]

            slide_names[slide_name] = SlideElementNames(
                image_names=image_names,
                text_names=text_names,
                chart_names=[],
                table_names=table_names,
            )

    presentation.save(api_client=api_client)
    if check_success:
        presentation.sync_from_cloud(api_client=api_client)
        for name, slide in zip(slide_names, presentation.slides):
            assert name
            assert slide.speaker_notes.read_text() == name
            # TODO: check element names

    return slide_names


def name_if_empty(
    element: AbstractElement,
    value: str,
    api_client: AbstractSlidesAPIClient,
    names_so_far: list[str] | None = None,
):
    if names_so_far is None:
        names_so_far = []

    if element.alt_text.title is not None:
        # Google API doesn't support setting alt text to empty string,
        # so we use space instead to indicate "empty"
        # And whitespaces aren't valid variable names anyway
        # Also names in a single slide must be unique
        current_name = element.alt_text.title.strip()
        if current_name and current_name not in names_so_far:
            return current_name

    element.set_alt_text(api_client=api_client, title=value)
    return value


def delete_slide_names(
    presentation_id: str, api_client: AbstractSlidesAPIClient
):  # | None = None):
    # api_client = api_client or AbstractSlidesAPIClient.get_default_api_client()
    presentation = AbstractPresentation.from_id(
        api_client=api_client, presentation_id=presentation_id
    )
    for slide in presentation.slides:
        slide.speaker_notes.write_text(" ", api_client=api_client)


def delete_alt_titles(presentation_id: str, api_client: AbstractSlidesAPIClient):  # | None = None):
    # api_client = api_client or AbstractSlidesAPIClient.get_default_api_client()
    presentation = AbstractPresentation.from_id(
        api_client=api_client, presentation_id=presentation_id
    )
    for slide in presentation.slides:
        for element in slide.page_elements_flat:
            if (
                isinstance(element, (AbstractShapeElement, AbstractImageElement))
                and element.alt_text.title
            ):
                logger.info(f"Deleting alt title {element.alt_text.title}")
                # Unfortunately, Google API doesn't support setting alt text to empty string, so use space instead
                element.set_alt_text(api_client=api_client, title=" ")

    presentation.save(api_client=api_client)


def name_elements(
    elements: list[AbstractElement],
    root_name: str,
    api_client: AbstractSlidesAPIClient,
    names_so_far: list[str] | None = None,
) -> list[str]:
    if names_so_far is None:
        names_so_far = []
    names = []
    if len(elements) == 1:
        names.append(
            name_if_empty(
                element=elements[0],
                value=root_name,
                api_client=api_client,
                names_so_far=names_so_far + names,
            )
        )

    elif len(elements) > 1:
        for i, e in enumerate(elements):
            names.append(
                name_if_empty(
                    element=e,
                    value=f"{root_name}_{i+1}",
                    api_client=api_client,
                    names_so_far=names_so_far + names,
                )
            )

    logger.info(f"Named {len(names)} {root_name.lower()}s: {names}")
    return names


def _is_pptx_chart_element(element: AbstractElement) -> bool:
    """Check if an element is a PowerPoint chart (GraphicFrame with embedded chart data)."""
    if not hasattr(element, "pptx_element") or element.pptx_element is None:
        return False
    return getattr(element.pptx_element, "has_chart", False)


def name_slide_elements(
    slide: AbstractSlide,
    slide_name: str = "",
    skip_empty_text_boxes: bool = False,
    min_image_size_cm: float = 4.0,
    api_client: AbstractSlidesAPIClient | None = None,
    llm: Optional[BaseLanguageModel] = None,
) -> SlideElementNames:
    """
    Name the elements in a slide.
    :param slide:
    :param slide_name: Only used for clearer logging, for the case when the caller has changed the name
    :param skip_empty_text_boxes:
    :param api_client:
    :param llm:
    :return:
    """
    if api_client is None:
        raise ValueError("API client is required")
    # api_client = api_client or AbstractSlidesAPIClient.get_default_api_client()
    all_images = [e for e in slide.page_elements_flat if isinstance(e, AbstractImageElement)]

    # Also find PowerPoint chart shapes (GraphicFrames with embedded charts)
    # These are separate from images in PowerPoint (unlike Google Slides where charts are images)
    pptx_chart_shapes = [
        e
        for e in slide.page_elements_flat
        if _is_pptx_chart_element(e) and not isinstance(e, AbstractImageElement)
    ]
    if pptx_chart_shapes:
        logger.info(
            f"Found {len(pptx_chart_shapes)} PowerPoint chart shape(s) in slide {slide_name}"
        )

    # Sort first by y, then by x
    all_images.sort(key=lambda x: x.absolute_position()[::-1])
    pptx_chart_shapes.sort(key=lambda x: x.absolute_position()[::-1])

    images = []
    charts = []
    for i, image in enumerate(all_images):
        if image.alt_text.title and image.alt_text.title.strip():  # already named
            continue

        if min(image.absolute_size(units=OutputUnit.CM)) < min_image_size_cm:
            logger.info(f"Skipping image number {i+1} in slide {slide_name} as it is too small")
            continue

        if llm is not None:
            image_data = image.get_image_data()
            if is_this_a_chart(
                image_bytes=image_data.content, mime_type=image_data.mime_type, llm=llm
            ):
                logger.info(f"Identified image number {i+1} in slide {slide_name} as a chart")
                charts.append(image)
            else:
                logger.info(f"Identified image number {i+1} in slide {slide_name} as an image")
                images.append(image)
        else:
            logger.info(
                f"Assuming image number {i+1} in slide {slide_name} is a chart as no LLM provided"
            )
            charts.append(image)

    # Add PowerPoint chart shapes to the charts list
    # Always add them - name_if_empty will handle deduplication if they already have names
    for chart_shape in pptx_chart_shapes:
        charts.append(chart_shape)
        logger.info(f"Adding PowerPoint chart shape to charts list in slide {slide_name}")

    image_names = name_elements(images, "Image", api_client)
    chart_names = name_elements(charts, "Chart", api_client)

    table_names = name_elements(
        [e for e in slide.page_elements_flat if isinstance(e, AbstractTableElement)],
        "Table",
        api_client,
    )

    text_names = []

    text_boxes = [e for e in slide.page_elements_flat if isinstance(e, AbstractShapeElement)]

    if skip_empty_text_boxes:
        text_boxes = [e for e in text_boxes if e.read_text().strip()]

    if not text_boxes:
        return SlideElementNames(
            image_names=image_names,
            text_names=text_names,
            chart_names=chart_names,
            table_names=table_names,
        )

    # Sort first by y, then by x
    text_boxes.sort(key=lambda x: x.absolute_position()[::-1])
    top_box = text_boxes[0]
    text_names.append(name_if_empty(top_box, "Title", api_client, names_so_far=text_names))

    other_boxes = text_boxes[1:]
    text_names.extend(name_elements(other_boxes, "Text", api_client, names_so_far=text_names))

    return SlideElementNames(
        image_names=image_names,
        text_names=text_names,
        chart_names=chart_names,
        table_names=table_names,
    )


def _extract_font_size_from_element(element: AbstractShapeElement) -> float:
    """Extract the dominant font size (in points) from a shape element's text styles.

    Thin wrapper around _extract_font_size_pt from abstract_slides.
    """
    try:
        return _extract_font_size_pt(element.styles(skip_whitespace=True))
    except Exception:
        return 12.0


def _extract_font_size_from_table_element(element: AbstractTableElement) -> float:
    """Extract the dominant font size (in points) from a table element's first cell.

    Delegates to _extract_font_size_from_table from abstract_slides.
    """
    return _extract_font_size_from_table(element)


def _build_element_size_meta(
    element: AbstractElement,
    font_size_pt: float,
) -> ElementSizeMeta | None:
    """Build ElementSizeMeta from an element's absolute size and a font size.

    Returns None if the element has zero or negative dimensions.
    """
    try:
        element_size = element.absolute_size(units=OutputUnit.IN)
        width, height = element_size[0], element_size[1]
        if width > 0 and height > 0:
            return ElementSizeMeta(
                box_width_inches=width,
                box_height_inches=height,
                font_size_pt=font_size_pt,
            )
    except Exception:
        pass
    return None
