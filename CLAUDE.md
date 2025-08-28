# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Testing
```bash
poetry run pytest                           # Run all tests
poetry run pytest tests/test_specific.py    # Run specific test file
poetry run pytest -v                        # Verbose test output
poetry run pytest -k "test_name"            # Run tests matching pattern
```

### Code Quality
```bash
poetry run black gslides_api/        # Format code with black
poetry run isort gslides_api/         # Sort imports with isort
```

### Package Management
```bash
poetry install                       # Install all dependencies
poetry install -E image              # Install with image extras (Pillow)
poetry build                         # Build distribution packages
```

### Debug/Development Scripts
```bash
# Use the debug script with credentials for testing
GSLIDES_CREDENTIALS_PATH=/path/to/credentials poetry run python debug_elements.py
```

## Architecture Overview

This is a Python library that provides a type-safe, Pydantic-based wrapper for the Google Slides API. The codebase follows a layered architecture:

### Core Components

1. **Domain Models** (`gslides_api/domain.py`): Pydantic models representing Google Slides API objects
   - All models inherit from `GSlidesBaseModel` 
   - Models include Size, Color, Shape properties, Image properties, etc.
   - Models provide `to_api_format()` method for API serialization

2. **Client Layer** (`gslides_api/client.py`): 
   - `GoogleAPIClient` manages Google API credentials and connections
   - Handles batch request queuing with `pending_batch_requests`
   - Manages connections to Slides, Sheets, and Drive APIs
   - Use `initialize_credentials()` function to set up authentication

3. **Presentation Model** (`gslides_api/presentation.py`):
   - Main entry point with `Presentation` class
   - Factory methods: `create_blank()`, `from_id()`, `from_json()`
   - Contains slides, masters, layouts, and notes master

4. **Page Hierarchy** (`gslides_api/page/`):
   - `BasePage` - abstract base for all page types
   - `Slide` - individual presentation slides
   - `Layout` - slide layout templates  
   - `Master` - master slide templates
   - `Notes` - speaker notes pages

5. **Element System** (`gslides_api/element/`):
   - `PageElementBase` - base class for all slide elements
   - `ShapeElement` - text boxes, shapes, etc.
   - Element types determined by `ElementKind` enum
   - Elements use discriminated unions for type-safe handling

6. **Request System** (`gslides_api/request/`):
   - Type-safe request builders for Google Slides API operations
   - All requests inherit from `GSlidesAPIRequest`
   - Supports batch operations through the client
   - Common requests: CreateShapeRequest, InsertTextRequest, UpdateTextStyleRequest

7. **Markdown Integration** (`gslides_api/markdown/`):
   - `from_markdown.py` - converts Markdown to Google Slides elements
   - `to_markdown.py` - converts Google Slides content back to Markdown
   - Uses `marko` library for Markdown parsing
   - Supports lists, text formatting, and structural elements

### Key Patterns

- **Pydantic Models**: All data structures use Pydantic v2 for validation and serialization
- **Type Safety**: Extensive use of Union types and discriminated unions for element handling
- **API Batching**: Client supports queuing multiple requests for efficient batch execution
- **Factory Methods**: Presentations and elements use factory methods for creation
- **Forward References**: Models use `model_rebuild()` to resolve circular dependencies

### Authentication Setup

The library requires Google API credentials. See `CREDENTIALS.md` for setup instructions. 
Use the `initialize_credentials("/path/to/credentials/")` function to configure authentication before making API calls.

## Development Patterns

### Model Design Principles

- **Avoid Dict Fields**: When creating Pydantic models, use `List[HelperClass]` instead of `Dict[str, Value]` patterns
- **Forward References**: Call `model_rebuild()` after all imports to resolve circular dependencies (see `__init__.py:170-174`)
- **API Serialization**: All models provide `to_api_format()` method for Google API compatibility
- **Discriminated Unions**: Element types use discriminated unions with `ElementKind` enum for type safety

### Testing Patterns

- Tests are organized by functionality (e.g., `test_markdown_features.py`, `test_element_kind.py`)
- Use `pytest` fixtures for common setup patterns
- Test both API format serialization and domain model validation
- Integration tests often require credentials setup for real API calls

### Code Organization

- **Domain Models**: Core Pydantic models in `domain.py` with extensive type annotations
- **Element Hierarchy**: Elements inherit from `PageElementBase` with kind-based discrimination
- **Request Builders**: Type-safe request objects that validate before API calls
- **Page Abstraction**: All page types (Slide, Layout, Master, Notes) inherit from `BasePage`

### Important Dependencies

- **Python 3.11+**: Required for latest Pydantic v2 features
- **Pydantic v2**: Core validation and serialization framework
- **Google API Client**: Official Google API Python library
- **marko**: Markdown parsing for conversion features
- **Pillow**: Optional image processing (install with `-E image`)

### API Integration Notes

- Client manages batch requests for efficiency - queue operations then execute
- All Google API responses are deserialized into type-safe Pydantic models
- Authentication supports both OAuth2 and Service Account workflows
- Enable both Google Slides API and Google Sheets API in your Google Cloud project