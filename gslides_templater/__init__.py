"""
Google Slides Templater
"""

from .core import SlidesTemplater, create_templater
from .markdown_processor import (
    MarkdownProcessor,
    MarkdownConfig,
    MarkdownProcessingError,
    UnsafeContentError,
    markdown_to_slides_elements,
    slides_elements_to_markdown,
    clean_markdown_for_slides
)
from .auth import (
    authenticate,
    AuthConfig,
    CredentialManager,
    Credentials,
    SlidesAPIError,
    AuthenticationError,
    TokenRefreshError,
    setup_oauth_flow,
    validate_credentials,
    get_credentials_info,
    create_service_account_template,
    create_oauth_template,
    check_credentials_file
)

# Конфигурационные классы из core
from .core import (
    SlidesConfig,
    LayoutConfig,
    ElementPosition,
    SlideElement,
    RateLimitExceededError,
    MaxRetriesExceededError,
    TemplateValidationError
)

__all__ = [
    # Основные классы
    'SlidesTemplater',
    'MarkdownProcessor',
    'CredentialManager',
    'Credentials',

    # Конфигурации
    'SlidesConfig',
    'LayoutConfig',
    'MarkdownConfig',
    'AuthConfig',
    'ElementPosition',
    'SlideElement',

    # Функции
    'create_templater',
    'authenticate',
    'setup_oauth_flow',
    'markdown_to_slides_elements',
    'slides_elements_to_markdown',
    'clean_markdown_for_slides',
    'validate_credentials',
    'get_credentials_info',
    'create_service_account_template',
    'create_oauth_template',
    'check_credentials_file',

    # Исключения
    'SlidesAPIError',
    'AuthenticationError',
    'TokenRefreshError',
    'RateLimitExceededError',
    'MaxRetriesExceededError',
    'TemplateValidationError',
    'MarkdownProcessingError',
    'UnsafeContentError',
]