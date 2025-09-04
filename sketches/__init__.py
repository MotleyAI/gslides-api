"""
Google Slides Templater
"""

from .auth import (AuthConfig, AuthenticationError, CredentialManager,
                   Credentials, SlidesAPIError, TokenRefreshError,
                   authenticate, check_credentials_file, create_oauth_template,
                   create_service_account_template, get_credentials_info,
                   setup_oauth_flow, validate_credentials)
from .core import (ElementPosition, LayoutConfig, MaxRetriesExceededError,
                   RateLimitExceededError, SlideElement, SlidesConfig,
                   SlidesTemplater, TemplateValidationError, create_templater)
from .markdown_processor import (MarkdownConfig, MarkdownProcessingError,
                                 MarkdownProcessor, UnsafeContentError,
                                 clean_markdown_for_slides,
                                 markdown_to_slides_elements,
                                 slides_elements_to_markdown)

__all__ = [
    "SlidesTemplater",
    "MarkdownProcessor",
    "CredentialManager",
    "Credentials",
    "SlidesConfig",
    "LayoutConfig",
    "MarkdownConfig",
    "AuthConfig",
    "ElementPosition",
    "SlideElement",
    "create_templater",
    "authenticate",
    "setup_oauth_flow",
    "markdown_to_slides_elements",
    "slides_elements_to_markdown",
    "clean_markdown_for_slides",
    "validate_credentials",
    "get_credentials_info",
    "create_service_account_template",
    "create_oauth_template",
    "check_credentials_file",
    # Исключения
    "SlidesAPIError",
    "AuthenticationError",
    "TokenRefreshError",
    "RateLimitExceededError",
    "MaxRetriesExceededError",
    "TemplateValidationError",
    "MarkdownProcessingError",
    "UnsafeContentError",
]
