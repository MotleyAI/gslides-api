# gslides-api MCP Server

The gslides-api MCP (Model Context Protocol) server exposes Google Slides operations as tools for AI assistants. This enables AI systems to read, modify, and manipulate Google Slides presentations programmatically.

## Table of Contents

- [Installation](#installation)
- [Credential Setup](#credential-setup)
- [Server Startup](#server-startup)
- [Tool Reference](#tool-reference)
- [Output Formats](#output-formats)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)

## Installation

Install gslides-api with the MCP extra:

```bash
pip install gslides-api[mcp]
# or with poetry
poetry add gslides-api -E mcp
```

## Credential Setup

See [CREDENTIALS.md](CREDENTIALS.md) for detailed instructions on setting up Google API credentials.

## Server Startup

### Command Line

```bash
# Using the CLI argument
python -m gslides_api.mcp.server --credential-path /path/to/credentials

# Using environment variable
export GSLIDES_CREDENTIALS_PATH=/path/to/credentials
python -m gslides_api.mcp.server

# With custom default output format
python -m gslides_api.mcp.server --credential-path /path/to/credentials --default-format outline

# Using the installed script
gslides-mcp --credential-path /path/to/credentials
```

### CLI Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `--credential-path` | Path to Google API credentials directory | `GSLIDES_CREDENTIALS_PATH` env var |
| `--default-format` | Default output format: `raw`, `domain`, or `outline` | `raw` |

### Environment Variables

| Variable | Description |
|----------|-------------|
| `GSLIDES_CREDENTIALS_PATH` | Alternative to `--credential-path` (CLI arg takes precedence) |

### MCP Configuration (`.mcp.json`)

Add to your `.mcp.json` configuration:

```json
{
  "gslides": {
    "type": "stdio",
    "command": "python",
    "args": ["-m", "gslides_api.mcp.server"]
  }
}
```

The server reads credentials from the `GSLIDES_CREDENTIALS_PATH` environment variable. Use `--credential-path` to override:

```json
{
  "gslides": {
    "type": "stdio",
    "command": "python",
    "args": ["-m", "gslides_api.mcp.server", "--credential-path", "/path/to/credentials"]
  }
}
```

## Tool Reference

### Query Tools

#### `get_presentation`

Get a full presentation by URL or deck ID.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `how` | string | No | Server default | Output format: `raw`, `domain`, or `outline` |

**Example:**
```json
{
  "presentation_id_or_url": "https://docs.google.com/presentation/d/1abc123/edit",
  "how": "outline"
}
```

---

#### `get_slide`

Get a single slide by name (first line of speaker notes).

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name (first line of speaker notes) |
| `how` | string | No | Server default | Output format: `raw`, `domain`, or `outline` |

**Example:**
```json
{
  "presentation_id_or_url": "1abc123",
  "slide_name": "Introduction",
  "how": "domain"
}
```

---

#### `get_element`

Get a single element by slide name and element name (alt-title).

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name |
| `element_name` | string | Yes | - | Element name (from alt-text title) |
| `how` | string | No | Server default | Output format |

---

#### `get_slide_thumbnail`

Get a slide thumbnail image, optionally with black borders around text boxes.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name |
| `add_text_box_borders` | boolean | No | `false` | Add 1pt black outlines to text boxes |
| `size` | string | No | `LARGE` | Thumbnail size: `SMALL` (200px), `MEDIUM` (800px), `LARGE` (1600px) |

**Returns:** JSON with base64-encoded PNG image data.

**Example Response:**
```json
{
  "success": true,
  "slide_name": "Introduction",
  "slide_id": "p1",
  "width": 1600,
  "height": 900,
  "mime_type": "image/png",
  "image_base64": "iVBORw0KGgo..."
}
```

---

### Markdown Tools

#### `read_element_markdown`

Read the text content of a shape element as markdown.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name |
| `element_name` | string | Yes | - | Element name (text box alt-title) |

**Returns:** Markdown content preserving bold, italic, bullets, etc.

---

#### `write_element_markdown`

Write markdown content to a shape element (text box).

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name |
| `element_name` | string | Yes | - | Element name (text box alt-title) |
| `markdown` | string | Yes | - | Markdown content to write |

---

### Image Tools

#### `replace_element_image`

Replace an image element with a new image from URL.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name |
| `element_name` | string | Yes | - | Element name (image alt-title) |
| `image_url` | string | Yes | - | URL of new image |

---

### Slide Manipulation Tools

#### `copy_slide`

Duplicate a slide within the presentation.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name to copy |
| `insertion_index` | integer | No | After original | Position for new slide (0-indexed) |

**Returns:** New slide info including `new_slide_id`.

---

#### `move_slide`

Move a slide to a new position in the presentation.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name to move |
| `insertion_index` | integer | Yes | - | New position (0-indexed) |

---

#### `delete_slide`

Delete a slide from the presentation.

**Arguments:**
| Name | Type | Required | Default | Description |
|------|------|----------|---------|-------------|
| `presentation_id_or_url` | string | Yes | - | Google Slides URL or presentation ID |
| `slide_name` | string | Yes | - | Slide name to delete |

---

## Output Formats

The `how` parameter controls the output format for query tools:

### `raw`

Returns the raw JSON response from the Google Slides API. Most verbose, includes all API fields.

```json
{
  "presentationId": "1abc123",
  "title": "My Presentation",
  "slides": [...],
  "masters": [...],
  "layouts": [...]
}
```

### `domain`

Returns the gslides-api domain object serialized via `model_dump()`. Structured Pydantic models with Python-friendly field names.

```json
{
  "presentationId": "1abc123",
  "title": "My Presentation",
  "slides": [...],
  "pageSize": {"width": {...}, "height": {...}}
}
```

### `outline`

Returns a condensed structure optimized for AI consumption. Includes slide names, element names, alt-descriptions, and text content as markdown.

```json
{
  "presentation_id": "1abc123",
  "title": "My Presentation",
  "slides": [
    {
      "slide_name": "Introduction",
      "slide_id": "p1",
      "elements": [
        {
          "element_name": "Title",
          "element_id": "e1",
          "type": "shape",
          "alt_description": "Main title",
          "content_markdown": "# Welcome to the Presentation"
        },
        {
          "element_name": "Hero Image",
          "element_id": "e2",
          "type": "image",
          "alt_description": "Company logo"
        }
      ]
    }
  ]
}
```

## Slide and Element Naming

### Slide Names

Slide names are derived from the **first line of the speaker notes**, stripped of whitespace. If a slide has no speaker notes or the first line is empty, it cannot be referenced by name (use the outline format to discover slide IDs).

### Element Names

Element names come from the **alt-text title** of each element. In Google Slides:
1. Right-click an element
2. Select "Alt text"
3. Enter a name in the "Title" field

Elements without alt-text titles cannot be referenced by name.

## Error Handling

All tools return detailed error responses when failures occur:

```json
{
  "error": true,
  "error_type": "SlideNotFound",
  "message": "No slide found with name 'Introduction'",
  "details": {
    "presentation_id": "1abc123",
    "searched_slide_name": "Introduction",
    "available_slides": ["Cover", "Overview", "Conclusion"]
  }
}
```

### Error Types

| Error Type | Description |
|------------|-------------|
| `ValidationError` | Invalid input parameter |
| `SlideNotFound` | Specified slide name not found |
| `ElementNotFound` | Specified element name not found on slide |
| `PresentationError:*` | Google API error (includes specific exception type) |

## Troubleshooting

### "API client not initialized"

Ensure you've provided a valid credential path via `--credential-path` or the `GSLIDES_CREDENTIALS_PATH` environment variable.

### OAuth flow opens repeatedly

Delete `token.json` in your credentials directory and re-authenticate.

### "Rate limit exceeded"

The server implements exponential backoff for rate limits. If you continue to see errors, reduce the frequency of API calls.

### Slide/Element not found

- Use `get_presentation` with `how=outline` to see all available slide and element names
- Ensure slides have speaker notes with content on the first line
- Ensure elements have alt-text titles set

### Import errors for MCP

Ensure you've installed the MCP extra:
```bash
pip install gslides-api[mcp]
```
