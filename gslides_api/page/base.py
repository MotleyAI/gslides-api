from enum import Enum
from typing import List, Optional

from pydantic import Field, model_validator

from gslides_api.domain import ColorScheme, PageBackgroundFill
from gslides_api.domain import GSlidesBaseModel
from gslides_api.element.element import PageElement
from gslides_api.element.base import ElementKind
from gslides_api.client import api_client, GoogleAPIClient


class PageType(Enum):
    """Enumeration of possible page types."""

    SLIDE = "SLIDE"
    MASTER = "MASTER"
    LAYOUT = "LAYOUT"
    NOTES = "NOTES"
    NOTES_MASTER = "NOTES_MASTER"


class PageProperties(GSlidesBaseModel):
    """Represents properties of a page."""

    pageBackgroundFill: Optional[PageBackgroundFill] = None
    colorScheme: Optional[ColorScheme] = None


def unroll_group_elements(elements: List["PageElement"]) -> List[PageElement]:
    out = []
    for element in elements:
        out.append(element)
        if element.type == ElementKind.GROUP:
            out.extend(unroll_group_elements(element.elementGroup.children))
    return out


class BasePage(GSlidesBaseModel):
    """Base class for all page types in a presentation."""

    objectId: Optional[str] = None
    pageElements: Optional[List[PageElement]] = (
        None  # Make optional to preserve original JSON exactly
    )
    revisionId: Optional[str] = None
    pageProperties: Optional[PageProperties] = None
    pageType: Optional[PageType] = Field(default=None, exclude=True)

    # Store the presentation ID for reference but exclude from model_dump
    presentation_id: Optional[str] = Field(default=None, exclude=True)

    def _propagate_presentation_id(self, presentation_id: Optional[str] = None) -> None:
        """Helper method to set presentation_id on all pageElements."""
        target_id = presentation_id if presentation_id is not None else self.presentation_id
        if target_id is not None and self.pageElements is not None:
            for element in self.pageElements:
                element.presentation_id = target_id
            # This recurses into GroupElement children
            for element in self.page_elements_flat:
                element.presentation_id = target_id

        if hasattr(self, "slideProperties") and self.slideProperties.notesPage is not None:
            self.slideProperties.notesPage.presentation_id = target_id

    @property
    def page_elements_flat(self) -> List[PageElement]:
        if self.pageElements is None:
            return []
        return unroll_group_elements(self.pageElements)

    @model_validator(mode="after")
    def set_element_parent_id(self) -> "BasePage":
        for e in self.page_elements_flat:
            e.slide_id = self.objectId
        return self

    @model_validator(mode="after")
    def set_presentation_id_on_elements(self) -> "BasePage":
        """Automatically set presentation_id on all pageElements after model creation."""
        self._propagate_presentation_id()
        return self

    def __setattr__(self, name: str, value) -> None:
        """Override setattr to propagate presentation_id when it's set directly."""
        super().__setattr__(name, value)
        # If presentation_id was just set, propagate it to pageElements
        if name == "presentation_id" and hasattr(self, "pageElements"):
            self._propagate_presentation_id(value)

    @classmethod
    def from_ids(
        cls, presentation_id: str, slide_id: str, api_client: Optional[GoogleAPIClient] = None
    ) -> "BasePage":
        # To avoid circular imports
        client = api_client or globals()["api_client"]
        json = client.get_slide_json(presentation_id, slide_id)
        new_slide = cls.model_validate(json)
        new_slide.presentation_id = presentation_id
        return new_slide

    def select_elements(self, kind: ElementKind) -> List[PageElement]:
        if self.pageElements is None:
            return []
        return [e for e in self.page_elements_flat if e.type == kind]

    @property
    def image_elements(self):
        return self.select_elements(ElementKind.IMAGE)

    def get_element_by_id(self, element_id: str) -> PageElement | None:
        if self.pageElements is None:
            return None
        return next((e for e in self.page_elements_flat if e.objectId == element_id), None)

    def get_element_by_alt_title(self, title: str) -> PageElement | None:
        if self.pageElements is None:
            return None
        return next(
            (
                e
                for e in self.page_elements_flat
                if isinstance(e.title, str) and e.title.strip() == title
            ),
            None,
        )


# Rebuild models to resolve forward references
BasePage.model_rebuild()
