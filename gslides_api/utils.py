import logging
from typing import Any, Dict, List

logger = logging.getLogger(__name__)


def dict_to_dot_separated_field_list(x: Dict[str, Any]) -> List[str]:
    """Convert a dictionary to a list of dot-separated fields."""
    out = []
    for k, v in x.items():
        if isinstance(v, dict):
            out += [f"{k}.{i}" for i in dict_to_dot_separated_field_list(v)]
        else:
            out.append(k)
    return out


def image_url_is_valid(url: str) -> bool:
    """
    Validate that an image URL is accessible and valid.

    Args:
        url: Image URL to validate

    Returns:
        True if URL appears to be valid and accessible
    """
    import urllib.error
    import urllib.request

    if not url or not url.startswith(("http://", "https://")):
        return False

    # Check for common image extensions
    valid_extensions = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp")
    url_lower = url.lower()

    # Allow URLs with parameters that might contain image extensions
    if not any(ext in url_lower for ext in valid_extensions):
        # If no obvious image extension, try a quick HEAD request
        try:
            req = urllib.request.Request(url, method="HEAD")
            req.add_header("User-Agent", "Mozilla/5.0 (compatible; Google-Slides-Templater/1.0)")

            with urllib.request.urlopen(req, timeout=5) as response:
                content_type = response.headers.get("Content-Type", "")
                return content_type.startswith("image/")

        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError):
            logger.warning(f"Could not validate image URL: {url}")
            return False

    return True
