---
name: prepare-presentation
description: Convert an existing presentation (Google Slides or PPTX) into a prepared template by copying it, naming elements by layout position, and replacing all content with placeholders.
---

# Prepare Presentation Skill

Convert an existing presentation (Google Slides or PPTX) into a prepared template by copying it,
naming elements by layout position, and replacing all content with placeholders.

## Workflow

### Step 1: Copy the source presentation

```bash
poetry run python .claude/skills/prepare-presentation/prepare.py copy --source <url_or_path> [--title <title>]
```

Creates a working copy of the presentation so the original is not modified.
Prints the new presentation's ID or file path.

### Step 2: Inspect slides and decide on names

For each slide, run:

```bash
poetry run python .claude/skills/prepare-presentation/prepare.py inspect --source <copied_id_or_path> [--slide <index>]
```

This prints:
- The path to a saved thumbnail image (read it to see the visual layout)
- The slide's `markdown()` output with all element metadata in HTML comments

Examine each slide's thumbnail and markdown to determine position-based names for elements and slides.

### Step 3: Apply names

```bash
poetry run python .claude/skills/prepare-presentation/prepare.py name --source <copied_id_or_path> --mapping '<json>'
```

The `--mapping` JSON has this structure:

```json
{
  "slides": [
    {
      "index": 0,
      "slide_name": "title_slide",
      "elements": {"Title": "title", "Text_1": "subtitle"}
    }
  ]
}
```

The element keys are the **display names from the inspect output** (alt_text title or objectId), and the values are the new names to assign. Display names are unique per slide.

### Step 4: Replace content with placeholders

```bash
poetry run python .claude/skills/prepare-presentation/prepare.py templatize --source <copied_id_or_path>
```

This replaces:
- **Named images** (min dimension >= 4 cm): replaced with a gray placeholder image
- **Named text shapes**: replaced with "Example Text" (or "# Example Header Style\n\nExample body style" for multi-style elements)
- **Named tables**: all cells replaced with "Text"

## Naming Guidelines

Names must describe **layout geometry and relative position only** — never content, topic, or semantic meaning. Imagine the slide with all text/images blanked out; names should still make sense from position alone.

### Element names (purely positional)
- Use position on the slide: "top_left", "top_right", "center", "bottom_left", etc.
- For multiple text elements in the same region, number them by reading order: "left_text_1", "left_text_2"
- All images/charts regardless of type -> "chart" (or "chart_1", "chart_2" if multiple)
- Tables -> "table" (or "table_1", "table_2" if multiple)
- The topmost/largest text element is typically "title"
- A text element directly below the title -> "subtitle"
- Small decorative elements (< 1cm) can be skipped

### Slide names
- Name by layout structure, not content: "title_fullwidth", "two_column", "chart_right", "grid_2x3", etc.
- First slide -> "title_slide"
- Last slide (if simple/closing) -> "closing"
- For repeated layouts, number them: "chart_right_1", "chart_right_2"

### What NOT to do
- Do NOT use content-derived names like "customer_name", "insights", "performance", "weekly"
- Do NOT name elements after what they currently say (e.g., "sidebar" because it says "Key Achievements")
- DO describe where the element sits: "left_text", "right_chart", "top_bar", "bottom_row_1"

## Notes

- Minimum image size for replacement: 4 cm in the smallest dimension
- Tables get "Text" in every cell, preserving the table shape (rows x cols)
- `write_text` handles markdown-to-style mapping automatically
- For Google Slides, credentials must be initialized before running
- For PPTX files, just pass the file path as `--source`
