import copy
import logging
from typing import List, Optional, Dict, Any

from gslides_api.page.page import Page
from gslides_api.domain import Size
from gslides_api.client import api_client, GoogleAPIClient
from gslides_api.domain import GSlidesBaseModel
from gslides_api.page.slide import Slide
from gslides_api.page.page import Layout, Master, NotesMaster

logger = logging.getLogger(__name__)


class Presentation(GSlidesBaseModel):
    """Represents a Google Slides presentation."""

    presentationId: Optional[str]
    pageSize: Size
    slides: Optional[List[Slide]] = None
    title: Optional[str] = None
    locale: Optional[str] = None
    revisionId: Optional[str] = None
    masters: Optional[List[Master]] = None
    layouts: Optional[List[Layout]] = None
    notesMaster: Optional[NotesMaster] = None

    @classmethod
    def create_blank(
        cls, title: str = "New Presentation", api_client: Optional[GoogleAPIClient] = None
    ) -> "Presentation":
        """Create a blank presentation in Google Slides."""
        client = api_client or globals()["api_client"]
        new_id = client.create_presentation({"title": title})
        return cls.from_id(new_id, api_client=api_client)

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
    def from_id(
        cls, presentation_id: str, api_client: Optional[GoogleAPIClient] = None
    ) -> "Presentation":
        client = api_client or globals()["api_client"]
        presentation_json = client.get_presentation_json(presentation_id)
        return cls.from_json(presentation_json)

    def copy_via_domain_objects(
        self, api_client: Optional[GoogleAPIClient] = None
    ) -> "Presentation":
        """Clone a presentation in Google Slides."""
        client = api_client or globals()["api_client"]
        config = self.to_api_format()
        config.pop("presentationId", None)
        config.pop("revisionId", None)
        new_id = client.create_presentation(config)
        return self.from_id(new_id, api_client=api_client)

    def copy_via_drive(
        self,
        copy_title: Optional[str] = None,
        api_client: Optional[GoogleAPIClient] = None,
        folder_id: Optional[str] = None,
    ):
        client = api_client or globals()["api_client"]
        copy_title = copy_title or f"Copy of {self.title}"
        new = client.copy_presentation(self.presentationId, copy_title, folder_id=folder_id)
        return self.from_id(new["id"], api_client=api_client)

    def sync_from_cloud(self, api_client: Optional[GoogleAPIClient] = None):
        re_p = Presentation.from_id(self.presentationId, api_client=api_client)
        self.__dict__ = re_p.__dict__

    def slide_from_id(self, slide_id: str) -> Optional[Page]:
        match = [s for s in self.slides if s.objectId == slide_id]
        if len(match) == 0:
            logger.error(
                f"Slide with id {slide_id} not found in presentation {self.presentationId}"
            )
            return None
        return match[0]

    def delete_slide(self, slide_id: str, api_client: Optional[GoogleAPIClient] = None):
        client = api_client or globals()["api_client"]
        client.delete_object(slide_id, self.presentationId)

    @property
    def url(self):
        if self.presentationId is None:
            return None
        return f"https://docs.google.com/presentation/d/{self.presentationId}/edit"
