from typing import Any, Dict, List

from gslides_api.domain.domain import GSlidesBaseModel


class GSlidesAPIRequest(GSlidesBaseModel):
    """Base class for all requests to the Google Slides API."""

    def to_request(self) -> List[Dict[str, Any]]:
        """Convert to the format expected by the Google Slides API."""
        request_name = self.__class__.__name__.replace("Request", "")
        # make first letter lowercase
        request_name = request_name[0].lower() + request_name[1:]

        return [{request_name: self.to_api_format()}]
