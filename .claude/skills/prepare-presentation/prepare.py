"""Prepare-presentation skill: copy, inspect, name, and templatize a presentation.

Usage:
    poetry run python .claude/skills/prepare-presentation/prepare.py copy --source <url_or_path> [--title <title>]
    poetry run python .claude/skills/prepare-presentation/prepare.py inspect --source <url_or_path> [--slide <index>]
    poetry run python .claude/skills/prepare-presentation/prepare.py name --source <url_or_path> --mapping <json>
    poetry run python .claude/skills/prepare-presentation/prepare.py templatize --source <url_or_path>
"""

import argparse
import json
import logging
import os
import tempfile

from gslides_api.adapters.abstract_slides import (
    AbstractImageElement,
    AbstractPresentation,
    AbstractShapeElement,
    AbstractTableElement,
    AbstractThumbnailSize,
)
from gslides_api.agnostic.element import MarkdownTableElement, TableData
from gslides_api.agnostic.units import OutputUnit
from gslides_api.common.presentation_id import normalize_presentation_id

logger = logging.getLogger(__name__)

# Minimum image dimension (in cm) to qualify for replacement
MIN_IMAGE_SIZE_CM = 4.0

# Path to placeholder image (relative to this script)
PLACEHOLDER_PATH = os.path.join(os.path.dirname(__file__), "placeholder.png")


def _resolve_source(source: str):
    """Determine adapter type from source string and return (api_client, presentation_id).

    - If source looks like a Google Slides URL or ID -> GSlidesAPIClient
    - If source is a file path ending in .pptx -> PowerPointAPIClient
    """
    source = source.strip()

    # PPTX file path
    if source.lower().endswith(".pptx") or os.path.isfile(source):
        from gslides_api.adapters.pptx_adapter import PowerPointAPIClient

        api_client = PowerPointAPIClient()
        return api_client, source

    # Google Slides URL or ID
    from gslides_api.adapters.gslides_adapter import GSlidesAPIClient

    presentation_id = normalize_presentation_id(source)
    if presentation_id:
        api_client = GSlidesAPIClient.get_default_api_client()
        credential_location = os.getenv("GSLIDES_CREDENTIALS_PATH")
        if credential_location:
            api_client.initialize_credentials(credential_location)
        return api_client, presentation_id

    # Fallback: try as GSlides ID anyway
    api_client = GSlidesAPIClient.get_default_api_client()
    credential_location = os.getenv("GSLIDES_CREDENTIALS_PATH")
    if credential_location:
        api_client.initialize_credentials(credential_location)
    return api_client, source


def _load_presentation(source: str):
    """Load a presentation from source. Returns (api_client, presentation)."""
    api_client, presentation_id = _resolve_source(source)
    presentation = AbstractPresentation.from_id(
        api_client=api_client, presentation_id=presentation_id
    )
    return api_client, presentation


def cmd_copy(args):
    """Copy a presentation and print the new ID/path."""
    api_client, presentation = _load_presentation(args.source)
    title = args.title or f"Template - {presentation.title or 'Untitled'}"

    copied = presentation.copy_via_drive(api_client=api_client, copy_title=title)
    presentation_id = copied.presentationId or ""

    # For PPTX, the presentationId is the file path
    print(f"Copied presentation: {presentation_id}")
    if hasattr(copied, "url"):
        try:
            print(f"URL: {copied.url}")
        except Exception:
            pass


def cmd_inspect(args):
    """Inspect slides: print thumbnails and markdown for each slide."""
    api_client, presentation = _load_presentation(args.source)

    if args.slide is not None:
        slides_to_inspect = [(args.slide, presentation.slides[args.slide])]
    else:
        slides_to_inspect = list(enumerate(presentation.slides))

    for i, slide in slides_to_inspect:
        print(f"\n{'='*60}")
        print(f"SLIDE {i} (objectId: {slide.objectId})")
        print(f"{'='*60}")

        # Get and save thumbnail
        try:
            thumb = slide.thumbnail(
                api_client=api_client,
                size=AbstractThumbnailSize.MEDIUM,
                include_data=True,
            )
            if thumb.content:
                ext = ".png" if "png" in thumb.mime_type else ".jpg"
                tmp = tempfile.NamedTemporaryFile(
                    delete=False, suffix=ext, prefix=f"slide_{i}_"
                )
                tmp.write(thumb.content)
                tmp.close()
                print(f"Thumbnail: {tmp.name}")
            elif thumb.contentUrl and thumb.contentUrl.startswith("file://"):
                print(f"Thumbnail: {thumb.contentUrl.replace('file://', '')}")
            else:
                print(f"Thumbnail URL: {thumb.contentUrl}")
        except Exception as e:
            print(f"Thumbnail error: {e}")

        # Print markdown representation
        print(f"\nMarkdown:\n")
        print(slide.markdown())
        print()


def cmd_name(args):
    """Apply naming to slides and elements from a JSON mapping."""
    api_client, presentation = _load_presentation(args.source)
    mapping = json.loads(args.mapping)

    for slide_mapping in mapping["slides"]:
        idx = slide_mapping["index"]
        slide = presentation.slides[idx]

        # Name the slide via speaker notes
        slide_name = slide_mapping.get("slide_name")
        if slide_name and slide.speaker_notes:
            slide.speaker_notes.write_text(api_client=api_client, content=slide_name)
            print(f"Slide {idx}: named '{slide_name}'")

        # Name elements via alt text (keys are display names from inspect output)
        elements_mapping = slide_mapping.get("elements", {})
        for old_name, new_name in elements_mapping.items():
            # Find element by display name (alt_text.title or objectId)
            found = False
            for element in slide.page_elements_flat:
                display_name = element.alt_text.title or element.objectId
                if display_name == old_name:
                    element.set_alt_text(api_client=api_client, title=new_name)
                    print(f"  Element '{old_name}' -> '{new_name}'")
                    found = True
                    break

            if not found:
                print(f"  WARNING: Element '{old_name}' not found on slide {idx}")

    presentation.save(api_client=api_client)
    print("\nNames applied and saved.")


def cmd_templatize(args):
    """Replace all named content with placeholders."""
    api_client, presentation = _load_presentation(args.source)

    for i, slide in enumerate(presentation.slides):
        print(f"Processing slide {i}...")

        for element in slide.page_elements_flat:
            name = element.alt_text.title
            if not name or not name.strip():
                continue  # Skip unnamed elements

            # Images: replace with placeholder if large enough
            if isinstance(element, AbstractImageElement):
                min_dim = min(element.absolute_size(units=OutputUnit.CM))
                if min_dim >= MIN_IMAGE_SIZE_CM:
                    element.replace_image(api_client=api_client, file=PLACEHOLDER_PATH)
                    print(f"  Replaced image '{name}' with placeholder")
                else:
                    print(f"  Skipped small image '{name}' ({min_dim:.1f} cm)")

            # Text shapes: replace with example text
            elif isinstance(element, AbstractShapeElement) and element.has_text:
                try:
                    styles = element.styles(skip_whitespace=True)
                except Exception:
                    styles = None

                if styles and len(styles) >= 2:
                    # Multi-style: provide header + body example
                    element.write_text(
                        api_client=api_client,
                        content="# Example Header Style\n\nExample body style",
                    )
                else:
                    element.write_text(api_client=api_client, content="Example Text")
                print(f"  Replaced text '{name}' with placeholder")

            # Tables: replace all cells with "Text"
            elif isinstance(element, AbstractTableElement):
                md_elem = element.to_markdown_element(name=name)
                if md_elem and md_elem.content:
                    rows, cols = md_elem.shape
                    # Build replacement table: all cells = "Text"
                    if rows > 0 and cols > 0:
                        new_headers = ["Text"] * cols
                        new_rows = [["Text"] * cols for _ in range(max(0, rows - 1))]
                        new_table_data = TableData(headers=new_headers, rows=new_rows)
                        new_md_elem = MarkdownTableElement(
                            name=name, content=new_table_data
                        )
                        element.update_content(
                            api_client=api_client,
                            markdown_content=new_md_elem,
                            check_shape=False,
                        )
                        print(f"  Replaced table '{name}' ({rows}x{cols}) with 'Text' placeholders")

    presentation.save(api_client=api_client)
    print("\nTemplatization complete.")


def main():
    parser = argparse.ArgumentParser(
        description="Prepare a presentation as a template"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # copy
    copy_parser = subparsers.add_parser("copy", help="Copy a presentation")
    copy_parser.add_argument("--source", required=True, help="Source presentation URL, ID, or file path")
    copy_parser.add_argument("--title", help="Title for the copy")

    # inspect
    inspect_parser = subparsers.add_parser("inspect", help="Inspect slides")
    inspect_parser.add_argument("--source", required=True, help="Presentation URL, ID, or file path")
    inspect_parser.add_argument("--slide", type=int, help="Specific slide index to inspect")

    # name
    name_parser = subparsers.add_parser("name", help="Apply names to slides and elements")
    name_parser.add_argument("--source", required=True, help="Presentation URL, ID, or file path")
    name_parser.add_argument("--mapping", required=True, help="JSON mapping of names")

    # templatize
    templatize_parser = subparsers.add_parser("templatize", help="Replace content with placeholders")
    templatize_parser.add_argument("--source", required=True, help="Presentation URL, ID, or file path")

    args = parser.parse_args()

    commands = {
        "copy": cmd_copy,
        "inspect": cmd_inspect,
        "name": cmd_name,
        "templatize": cmd_templatize,
    }
    commands[args.command](args)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    main()
