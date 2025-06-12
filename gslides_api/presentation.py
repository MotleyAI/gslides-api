import copy
import logging
from typing import List, Optional, Dict, Any

from gslides_api import Size, Dimension
from gslides_api.execute import (
    copy_presentation,
    create_presentation,
    delete_object,
    get_presentation_json,
    batch_update,
)
from gslides_api.page import Page
from gslides_api.domain import GSlidesBaseModel

logger = logging.getLogger(__name__)


class Presentation(GSlidesBaseModel):
    """Represents a Google Slides presentation."""

    presentationId: Optional[str]
    pageSize: Size
    slides: Optional[List[Page]] = None
    title: Optional[str] = None
    locale: Optional[str] = None
    revisionId: Optional[str] = None
    masters: Optional[List[Page]] = None
    layouts: Optional[List[Page]] = None
    notesMaster: Optional[Page] = None

    @classmethod
    def create_blank(cls, title: str = "New Presentation") -> "Presentation":
        """Create a blank presentation in Google Slides."""
        new_id = create_presentation({"title": title})
        return cls.from_id(new_id)

    @classmethod
    def from_json(cls, json_data: Dict[str, Any]) -> "Presentation":
        """
        Convert a JSON representation of a presentation into a Presentation object.

        Args:
            json_data: The JSON data representing a Google Slides presentation

        Returns:
            A Presentation object populated with the data from the JSON
        """
        # Use Pydantic's model_validate to parse the processed JSON
        out = cls.model_validate(json_data)

        # Set presentation_id on slides
        if out.slides is not None:
            for s in out.slides:
                s.presentation_id = out.presentationId

        return out

    @classmethod
    def from_id(cls, presentation_id: str) -> "Presentation":
        presentation_json = get_presentation_json(presentation_id)
        return cls.from_json(presentation_json)

    def copy_via_domain_objects(self) -> "Presentation":
        """Clone a presentation in Google Slides."""
        config = self.to_api_format()
        config.pop("presentationId", None)
        config.pop("revisionId", None)
        new_id = create_presentation(config)
        return self.from_id(new_id)

    def copy_via_drive(self, copy_title: Optional[str] = None):
        copy_title = copy_title or f"Copy of {self.title}"
        new = copy_presentation(self.presentationId, copy_title)
        return self.from_id(new["id"])

    def sync_from_cloud(self):
        re_p = Presentation.from_id(self.presentationId)
        self.__dict__ = re_p.__dict__

    def slide_from_id(self, slide_id: str) -> Optional[Page]:
        match = [s for s in self.slides if s.objectId == slide_id]
        if len(match) == 0:
            logger.error(
                f"Slide with id {slide_id} not found in presentation {self.presentationId}"
            )
            return None
        return match[0]

    def delete_slide(self, slide_id: str):
        delete_object(slide_id, self.presentationId)

    @property
    def url(self):
        if self.presentationId is None:
            return None
        return f"https://docs.google.com/presentation/d/{self.presentationId}/edit"
