# Google Slides Templater

A powerful Python library for working with Google Slides API. Create presentations, templates, and automate slide generation with full Markdown support and precise element positioning.

## Features

- **Create and edit presentations** programmatically
- **Full Markdown support** - convert Markdown to Google Slides formatting with proper symbol removal
- **Smart template system** with automatic placeholder detection and position tracking
- **Precise element positioning** - EMU coordinate system with layer information
- **Advanced image handling** - URL validation and error recovery
- **Batch presentation processing** with improved error handling
- **Multiple authentication methods** - Service Account, OAuth, Application Default Credentials
- **Comprehensive error handling** and logging

## Installation

```bash
pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client pydantic
```

Optional (for enhanced Markdown support):
```bash
pip install markdown
```

## Quick Start

### 1. Authentication Setup

Choose one of the authentication methods:

#### Option A: Service Account (Recommended for automation)
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing
3. Enable Google Slides API and Google Drive API
4. Create a Service Account and download JSON credentials
5. Save as `service_account.json`

#### Option B: OAuth (Recommended for personal use)
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create OAuth 2.0 Client ID (Desktop application)
3. Download credentials and save as `credentials.json`

### 2. Basic Usage

```python
from sketches import create_templater, AuthConfig

# Initialize with AuthConfig
auth_config = AuthConfig(
    service_account_file='service_account.json'
    # or credentials_path='credentials.json', token_path='token.json'
)
templater = create_templater(auth_config=auth_config)

# Create a new presentation
presentation_id = templater.create_presentation("My Report")

# Add a slide with Markdown content
slide_id = templater.add_markdown_slide(presentation_id, '''
# Sales Report Q4 2024

## Key Metrics

Sales increased by **25%** this quarter thanks to:
- New product launches
- Team expansion
- *Excellent* customer feedback

### Next Steps
- Expand to new markets
- Hire additional staff
- Continue innovation
''')

# Get presentation URL
url = templater.get_presentation_url(presentation_id)
print(f"Presentation: {url}")
```

### 3. Create Presentation from Markdown

```python
markdown_content = '''
# Introduction
Welcome to our quarterly review.

# Key Achievements
We exceeded all targets:
- Revenue: **$2.5M** (+25%)
- Customers: **10,000** (+40%)
- Satisfaction: **4.8/5** stars

# Future Plans
Next quarter we will focus on:
- Product development
- Market expansion
- Team growth
'''

# Each level 1 header (#) becomes a new slide
presentation_id = templater.create_presentation_from_markdown(
    markdown_content, 
    title="Q4 2024 Review"
)
```

## Template System with Positioning

### Creating Templates with Position Data

Extract templates from existing presentations with precise positioning:

```python
# Create template from existing presentation
template = templater.create_template(
    presentation_id="your_presentation_id",
    template_name="quarterly_report",
    debug=True  # Enable position extraction debugging
)

# Save template to file (includes position and layer data)
templater.save_template(template, "quarterly_report.json")

# Template now includes:
# - Element positions (x, y, width, height in EMU units)
# - Layer information for proper stacking
# - Slide dimensions and layout configuration
```

### Using Templates with Enhanced Data

```python
# Load template with position data
template = templater.load_template("quarterly_report.json")

# Prepare data (supports full Markdown formatting)
quarterly_data = {
    "title": "# **Q1 2025 Report** 🚀\n\n*Executive Summary*",
    "summary": """**Outstanding Results** this quarter:

- Revenue growth: **30%** 📈
- Customer satisfaction: ***Excellent*** ratings
- `Key metrics` exceeded all expectations

> This is our best quarter yet!""",
    "metrics": """
## Key Performance Indicators

1. **Financial Metrics**
   - Revenue: **$3.2M** (+30% YoY)
   - Profit: *$980K* (+45% YoY)
   - Expenses: ~~$2.2M~~ **$2.1M** (optimized)

2. **Customer Metrics**
   - Total customers: `12,500`
   - Net promoter score: **9.2/10**
   - Retention rate: ***95%***
""",
    "chart_image": "https://via.placeholder.com/800x600/4285f4/ffffff?text=Q1+Chart"
}

# Apply template with enhanced positioning
new_presentation_id = templater.apply_template(
    template=template, 
    data=quarterly_data, 
    title="Q1 2025 Final Report"
)
```

### Template Validation and Previews

```python
# Comprehensive validation
validation = templater.validate_template_data(template, quarterly_data)

if validation['valid']:
    print("✅ Data is valid!")
    if validation['warnings']:
        print("⚠️ Warnings:")
        for warning in validation['warnings']:
            print(f"  - {warning}")
else:
    print("❌ Issues found:")
    for issue in validation['missing_placeholders']:
        print(f"  - Missing: {issue}")
    for issue in validation['invalid_types']:
        print(f"  - Invalid type: {issue}")

# Preview changes with position info
preview = templater.preview_template_application(template, quarterly_data)
for change in preview['changes']:
    pos = change['position']
    print(f"Slide {change['slide']}: {change['placeholder']}")
    print(f"  Position: [{pos['x']:.0f},{pos['y']:.0f}] Layer: {change['layer']}")
    print(f"  Change: {change['old_value'][:30]}... → {change['new_value'][:30]}...")
```

## Advanced Markdown Support

The library now properly processes Markdown with symbol removal:

```python
markdown_text = '''
# Header 1 **with bold**
## Header 2 *with italic*
### Header 3 `with code`

**Bold text** and *italic text* and ***bold italic***

~~Strikethrough~~ text for corrections

`Inline code` formatting with proper colors

- Bullet lists
  - With nested items
  - And **formatting**
- Multiple levels supported

1. Numbered lists
2. Are also supported
3. With *proper* formatting

> Important blockquotes
> With multiple lines
> And **formatting**

Mixed formatting: **Bold with *nested italic*** text.
'''

# Markdown symbols are properly removed, formatting is applied
templater.add_markdown_slide(presentation_id, markdown_text)
```

**Key Improvements:**
- Markdown symbols (`**`, `*`, `#`, etc.) are properly removed from final text
- Proper text insertion with subsequent formatting application
- Support for nested formatting combinations
- Enhanced quote and code block processing

## Configuration Options

### Slides Configuration

```python
from sketches import SlidesConfig, LayoutConfig, MarkdownConfig

slides_config = SlidesConfig(
    max_retries=5,
    rate_limit_delay=0.2,
    batch_size=50,
    max_total_requests=1000,
    request_timeout=60.0
)

layout_config = LayoutConfig(
    slide_width=720,
    slide_height=540,
    margin_x=50,
    margin_y=50,
    default_width=620,
    default_height=200
)

markdown_config = MarkdownConfig(
    header_sizes={1: 36, 2: 28, 3: 24, 4: 20, 5: 18, 6: 16},
    max_text_length=1000000
)

# Create templater with custom configurations
auth_config = AuthConfig(service_account_file='service_account.json')
templater = create_templater(
    auth_config=auth_config,
    slides_config=slides_config,
    layout_config=layout_config,
    markdown_config=markdown_config
)
```

## Error Handling and Recovery

### Image URL Validation

```python
# Automatic image URL validation
try:
    templater.replace_image(presentation_id, image_id, "https://example.com/image.jpg")
except ValueError as e:
    print(f"Image URL invalid: {e}")

# Template application with image error recovery
template_data = {
    "title": "# Report Title",
    "chart": "https://unreachable-url.com/chart.png",  # This will be skipped
    "photo": "https://valid-url.com/photo.jpg"        # This will work
}

# Images are processed individually - valid ones succeed, invalid ones are skipped
new_presentation_id = templater.apply_template(template, template_data)
```

### Comprehensive Error Handling

```python
from sketches import (
    SlidesAPIError,
    AuthenticationError,
    TemplateValidationError,
    MarkdownProcessingError,
    RateLimitExceededError
)

try:
    presentation_id = templater.create_presentation("Test")
    templater.apply_template(template, data)
except AuthenticationError as e:
    print(f"Authentication failed: {e}")
except TemplateValidationError as e:
    print(f"Template validation failed: {e}")
except MarkdownProcessingError as e:
    print(f"Markdown processing error: {e}")
except RateLimitExceededError as e:
    print(f"Rate limit exceeded: {e}")
except SlidesAPIError as e:
    print(f"Google Slides API error: {e}")
```

## Advanced Features

### Element Positioning

```python
from sketches import ElementPosition

# Add text box with precise positioning
position = ElementPosition(
    x=100.0,  # X coordinate in points
    y=200.0,  # Y coordinate in points  
    width=400.0,  # Width in points
    height=100.0,  # Height in points
    layer=1  # Stacking layer
)

text_box_id = templater.add_text_box(
    presentation_id=presentation_id,
    slide_id=slide_id,
    text="**Precisely positioned text**",
    position=position
)
```

### Batch Operations with Error Recovery

```python
# Enhanced batch processing
template = templater.load_template("report_template.json")

quarterly_reports = []
for quarter in ["Q1", "Q2", "Q3", "Q4"]:
    try:
        data = {
            "title": f"# {quarter} 2024 Report",
            "revenue": f"$3.{quarter[1]}M",
            "chart": f"https://charts.example.com/{quarter.lower()}.png"
        }
        
        pres_id = templater.apply_template(template, data, f"{quarter} 2024 Report")
        quarterly_reports.append(pres_id)
        print(f"✅ Created {quarter} report: {pres_id}")
        
    except Exception as e:
        print(f"❌ Failed to create {quarter} report: {e}")
        continue

print(f"Successfully created {len(quarterly_reports)} reports")
```

## Template Structure and Metadata

### Enhanced Template Information

```python
# Get comprehensive template details
template_info = templater.get_template_info(template)
print(f"Template: {template_info['name']}")
print(f"Source: {template_info['source_presentation_id']}")
print(f"Created: {template_info['created_at']}")
print(f"Slides: {template_info['total_slides']}")
print(f"Placeholders: {template_info['total_placeholders']}")
print(f"Element types: {template_info['element_types']}")

# Template includes positioning metadata
slide_size = template['slide_size']
print(f"Slide dimensions: {slide_size['width']}x{slide_size['height']} {slide_size['units']}")

layout = template['layout_config']
print(f"Margins: {layout['margin_x']}x{layout['margin_y']} points")

# Position information for each placeholder
for name, info in template['placeholders'].items():
    position = info['position']
    if position:
        print(f"{name}: [{position['x']:.0f},{position['y']:.0f}] "
              f"size:{position['width']:.0f}x{position['height']:.0f} "
              f"layer:{position['layer']}")
```

## Authentication Methods

### Authentication Methods

```python
from sketches import AuthConfig, authenticate

# Method 1: AuthConfig (Recommended)
auth_config = AuthConfig(
    service_account_file='service_account.json',
    # or
    credentials_path='credentials.json',
    token_path='token.json',
    scopes=['https://www.googleapis.com/auth/presentations',
            'https://www.googleapis.com/auth/drive']
)

templater = create_templater(auth_config=auth_config)

# Method 2: Direct authentication
credentials = authenticate(auth_config)
templater = SlidesTemplater(credentials)
```

## API Reference

### Core Classes

#### SlidesTemplater
Main class with enhanced positioning and error handling.

#### AuthConfig
Configuration for authentication:
- `service_account_file`: Path to service account JSON
- `credentials_path`: Path to OAuth credentials JSON  
- `token_path`: Path to save OAuth token
- `scopes`: List of OAuth scopes
- `use_application_default`: Whether to try Application Default Credentials

#### SlidesConfig
Configuration for API behavior:
- `max_retries`: Maximum retry attempts (default: 3)
- `rate_limit_delay`: Delay between requests (default: 0.1s)
- `batch_size`: Maximum requests per batch (default: 50)
- `request_timeout`: Request timeout (default: 30s)

#### LayoutConfig  
Configuration for slide layout:
- `slide_width/height`: Slide dimensions in points
- `margin_x/y`: Default margins
- `default_width/height`: Default element sizes

#### MarkdownConfig
Configuration for Markdown processing:
- `header_sizes`: Font sizes for each header level
- `max_text_length`: Maximum text length for processing

### Enhanced Methods

#### Template Methods (Updated)
- `create_template(presentation_id, template_name, debug=False)` - Enhanced with position extraction
- `apply_template(template, data, title=None)` - Improved error handling and image validation
- `validate_template_data(template, data)` - Enhanced validation with warnings
- `preview_template_application(template, data)` - Preview with position information

#### New Utility Methods
- `get_credentials_info(credentials)` - Get credential information
- `check_credentials_file(file_path)` - Validate credentials file
- `create_sample_data(template)` - Generate sample data for template testing

## Migration Guide

### Updated Template Structure
Templates now include:
- Position coordinates in EMU units
- Layer information for element stacking  
- Slide size and layout configuration
- Enhanced placeholder metadata

### Improved Error Handling
- Individual image processing (failures don't stop entire batch)
- Better validation with warnings
- Specific exception types for different error categories

## Examples

Check the `examples/` directory for complete examples:

- `example_1_create_template.py` - Create template with positioning data
- `example_2_use_template.py` - Apply template with enhanced error handling

## Troubleshooting

### Common Issues

1. **Markdown symbols still visible**: Ensure you're using the latest version and check browser cache
2. **Image replacement fails**: URLs must be publicly accessible; use image validation
3. **Position data missing**: Use `debug=True` when creating templates to see extraction details
4. **Authentication errors**: Check credentials file format with `check_credentials_file()`

### Debug Mode

```python
# Enable debugging for template creation
template = templater.create_template(
    presentation_id="your_id",
    template_name="debug_template", 
    debug=True  # Shows detailed position extraction
)

# Enable logging
import logging
logging.basicConfig(level=logging.DEBUG)
```