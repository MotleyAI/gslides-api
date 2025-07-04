# gslides-api

A Python library for working with Google Slides API using Pydantic domain objects.

## Overview

This library provides a Pythonic interface to the Google Slides API with:

- **Pydantic domain objects** that match the JSON structure returned by the Google Slides API
- **Type-safe operations** with full type hints support
- **Easy-to-use methods** for creating, reading, and manipulating Google Slides presentations
- **Comprehensive coverage** of Google Slides API features

## Installation

```bash
pip install gslides-api
```

## Quick Start

### Authentication

First, set up your Google API credentials. See [CREDENTIALS.md](CREDENTIALS.md) for detailed instructions.

```python
from gslides_api import initialize_credentials

# Initialize with your credentials directory
initialize_credentials("/path/to/your/credentials/")
```

### Basic Usage

```python
from gslides_api import Presentation

# Load an existing presentation
presentation = Presentation.from_id("your-presentation-id")

# Create a new blank presentation
new_presentation = Presentation.create_blank("My New Presentation")

# Access slides
for slide in presentation.slides:
    print(f"Slide ID: {slide.objectId}")
    
# Create a new slide
new_slide = presentation.add_slide()
```



## Features

- **Domain Objects**: Complete Pydantic models for all Google Slides API objects
- **Presentations**: Create, load, copy, and manipulate presentations
- **Slides**: Add, remove, duplicate, and reorder slides
- **Elements**: Work with text boxes, shapes, images, and other slide elements
- **Layouts**: Access and use slide layouts and masters
- **Requests**: Type-safe request builders for batch operations
- **Markdown Support**: Convert between Markdown and Google Slides content

## API Coverage

The library covers most Google Slides API functionality including:

- Presentations and slides management
- Text elements and formatting
- Shapes and images
- Tables and charts
- Page layouts and masters
- Batch update operations

## Requirements

- Python 3.8+
- Google API credentials (OAuth2 or Service Account)

## Dependencies

- `google-api-python-client` - Google API client library
- `google-auth-oauthlib` - OAuth2 authentication
- `pydantic` - Data validation and serialization
- `marko` - Markdown processing
- `protobuf` - Protocol buffer support

## Development

### Running Tests

```bash
pip install -e ".[test]"
pytest
```

### Code Formatting

```bash
pip install -e ".[dev]"
black gslides_api/
isort gslides_api/
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Related Projects

- [md2googleslides](https://github.com/ShoggothAI/md2googleslides) - TypeScript library for creating slides from Markdown
- [gslides](https://github.com/michael-gracie/gslides) - Python library focused on charts and tables
- [gslides-maker](https://github.com/vilmacio/gslides-maker) - Generate slides from Wikipedia content

## Acknowledgments

This library is built on top of the excellent Google API Python client and leverages the power of Pydantic for type-safe data handling.
