"""Utility for normalizing Google Slides presentation IDs from URLs."""

import logging
import re
from urllib.parse import urlparse

logger = logging.getLogger(__name__)


def normalize_presentation_id(presentation_id_or_url: str) -> str:
    """
    Extract presentation ID from a presentation ID or URL (e.g. "https://docs.google.com/presentation/d/1234567890/edit?slide=id.p1").
    """
    presentation_id_or_url = presentation_id_or_url.strip()

    if presentation_id_or_url.startswith("https://docs.google.com/presentation/d/"):
        try:
            parsed_url = urlparse(presentation_id_or_url)
            parts = parsed_url.path.split("/")
            idx = parts.index("presentation")
            assert parts[idx + 1] == "d"
            return parts[idx + 2]
        except (TypeError, AssertionError, ValueError, IndexError) as e:
            logger.warning(f"Error extracting presentation ID from {presentation_id_or_url}: {e}")

    # check if a valid presentation ID is provided
    if re.match(r"^[a-zA-Z0-9_-]{25,}$", presentation_id_or_url):
        return presentation_id_or_url

    raise ValueError(f"Invalid presentation ID or URL: {presentation_id_or_url}")
