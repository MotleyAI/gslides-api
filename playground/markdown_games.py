import os

import marko

from gslides_api import Presentation, Slide, initialize_credentials

here = os.path.dirname(os.path.abspath(__file__))
credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
initialize_credentials(credential_location)

md = """This is a ***very*** *important* report with **bold** text.

* It illustrates **bullet points**
  * With nested sub-points
  * And even more `code` blocks
    * Third level nesting
* And even `code` blocks
* Plus *italic* formatting
  * Nested italic *emphasis*
  * With **bold** nested items

Here's a [link to Google](https://google.com) for testing hyperlinks.

Some ~~strikethrough~~ text to test deletion formatting.

Ordered list example:
1. First numbered item
   1. Nested numbered sub-item
   2. Another nested item with **bold**
      1. Third level numbering
2. Second with `inline code`
   1. Nested under third
   2. Final nested item

Mixed content with [links](https://example.com) and ~~crossed out~~ text.
"""

out = marko.parse(md)
print(out)
print("Markdown converted to elements!")


# Setup, choose presentation
presentation_id = (
    "1bW53VB1GqljfLEt8qS3ZaFiq47lgF9iMpossptwuato"  # "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"
)
source_presentation = Presentation.from_id(presentation_id)

s = source_presentation.slides[8]
new_slide = s.write_copy(9)
new_slide.pageElements[3].write_text(md, as_markdown=True)
new_slide = Slide.from_ids(presentation_id, new_slide.objectId)

# Save the API response for testing
import json
import os

api_response = json.loads(
    new_slide.pageElements[3].model_dump_json()
)  # Convert to JSON-serializable format
test_data_dir = os.path.join(os.path.dirname(__file__), "..", "tests", "test_data")
os.makedirs(test_data_dir, exist_ok=True)
test_data_file = os.path.join(test_data_dir, "markdown_api_response.json")

test_data = {"original_markdown": md.strip(), "api_response": api_response}
#
# with open(test_data_file, "w") as f:
#     json.dump(test_data, f, indent=2)

print(f"API response saved to: {test_data_file}")
print(api_response)
print("Kind of a copy written!")

# Test markdown reconstruction
print("\n" + "=" * 50)
print("TESTING MARKDOWN RECONSTRUCTION")
print("=" * 50)

original_markdown = md.strip()
print(f"Original markdown:\n{repr(original_markdown)}")

reconstructed_markdown = new_slide.pageElements[3].to_markdown()
print(f"\nReconstructed markdown:\n{repr(reconstructed_markdown)}")

# Test if reconstruction is successful
if reconstructed_markdown:
    print(f"\nOriginal formatted:\n{original_markdown}")
    print(f"\nReconstructed formatted:\n{reconstructed_markdown}")

    # Simple comparison (ignoring minor whitespace differences)
    original_normalized = " ".join(original_markdown.split())
    reconstructed_normalized = (
        " ".join(reconstructed_markdown.split()) if reconstructed_markdown else ""
    )

    if original_normalized == reconstructed_normalized:
        print("\n✅ SUCCESS: Markdown reconstruction matches original!")
    else:
        print("\n❌ DIFFERENCE: Markdown reconstruction differs from original")
        print(f"Original normalized: {repr(original_normalized)}")
        print(f"Reconstructed normalized: {repr(reconstructed_normalized)}")

        # Show character-by-character differences
        import difflib

        diff = list(
            difflib.unified_diff(
                original_markdown.splitlines(keepends=True),
                reconstructed_markdown.splitlines(keepends=True) if reconstructed_markdown else [],
                fromfile="original",
                tofile="reconstructed",
                lineterm="",
            )
        )
        if diff:
            print("\nDetailed differences:")
            for line in diff:
                print(line.rstrip())
else:
    print("\n❌ FAILED: No markdown was reconstructed")
