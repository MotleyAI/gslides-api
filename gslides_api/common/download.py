import logging

import httpx

from gslides_api.common.log_time import log_time
from gslides_api.common.retry import retry

logger = logging.getLogger(__name__)


@log_time
def download_file(url: str, timeout: int = 30, **retry_kwargs) -> httpx.Response:
    """
    Downloads file from a URL with retry mechanism and error handling.

    Args:
        url: URL to download the file from
        timeout: Timeout in seconds for the request (default: 30)
        **retry_kwargs: Additional retry parameters (max_attempts, initial_delay, max_delay,
                       exponential_base, jitter, exceptions). Defaults: max_attempts=3,
                       initial_delay=1.0, max_delay=10.0

    Returns:
        The downloaded file as an httpx.Response object

    Raises:
        ValueError: If the URL is invalid or empty
        httpx.RequestError: If the download fails after all retries
    """
    if not url or not url.strip():
        raise ValueError("URL cannot be empty")

    # Set default retry parameters
    retry_params = {
        "max_attempts": 3,
        "initial_delay": 1.0,
        "max_delay": 10.0,
        "exceptions": (
            httpx.RequestError,
            httpx.TimeoutException,
        ),
    }
    # Override with user-provided kwargs
    retry_params.update(retry_kwargs)

    @retry(**retry_params)
    def _download() -> httpx.Response:
        response = httpx.get(url, timeout=timeout)
        response.raise_for_status()
        return response

    return _download()


def download_text_file(url: str, timeout: int = 30, **retry_kwargs) -> str:
    """
    Calls download_file and returns the text content of the file.

    Args:
        url: URL to download the file from
        timeout: Timeout in seconds for the request (default: 30)
        **retry_kwargs: Additional retry parameters passed to download_file

    Returns:
        The text content of the file
    """
    response = download_file(url, timeout=timeout, **retry_kwargs)
    return response.text


@log_time
def download_binary_file(url: str, timeout: int = 30, **retry_kwargs) -> tuple[bytes, int | None]:
    """
    Calls download_file and returns the binary content and file size of the file.

    Args:
        url: URL to download the file from
        timeout: Timeout in seconds for the request (default: 30)
        **retry_kwargs: Additional retry parameters passed to download_file

    Returns:
        A tuple of (binary_content, file_size_from_headers)
        file_size_from_headers may be None if Content-Length header is not present
    """
    response = download_file(url, timeout=timeout, **retry_kwargs)
    content_length = response.headers.get("content-length")
    file_size = int(content_length) if content_length else None
    return response.content, file_size
