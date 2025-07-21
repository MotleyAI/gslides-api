import json
import os

from gslides_api import Presentation, Slide, initialize_credentials

here = os.path.dirname(os.path.abspath(__file__))
credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
initialize_credentials(credential_location)

complex_md = """This is a ***very*** *important* report with **bold** text.

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

# Setup, choose presentation
presentation_id = "1FHbC3ZXsEDUUNtQbxyyDQ3EFjwwt13_WovJAiYxhmOU"
source_presentation = Presentation.from_id(presentation_id)

s = source_presentation.slides[1]
# test delete text with bullets
new_slide = s.duplicate()
new_slide.get_element_by_alt_title("text_1").delete_text()

new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("text_1").read_text()
assert re_md == ""
new_slide.delete()

# test delete text without bullets
new_slide = s.duplicate()
new_slide.get_element_by_alt_title("title_1").delete_text()

new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("title_1").read_text()
assert re_md == ""
new_slide.delete()

# test write text with bullets
new_slide = s.duplicate()


md = "Oh what a text\n* Bullet points\n* And more\n1. Numbered items\n2. And more"
new_slide.get_element_by_alt_title("text_1").write_text(md, as_markdown=True)
new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("text_1").read_text()
assert re_md == md
new_slide.delete()

# Test write complicated markdown
medium_md = """This is a ***very*** *important* report with **bold** text.

* It illustrates **bullet points**
  * With nested sub-points
  * And even more `code` blocks"""
new_slide = s.duplicate()
new_slide.get_element_by_alt_title("text_1").write_text(medium_md, as_markdown=True)
new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("text_1").read_text()
# assert re_md == medium_md
new_slide.delete()


new_slide = s.duplicate()
new_slide.get_element_by_alt_title("text_1").write_text(complex_md, as_markdown=True)
new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("text_1").read_text()
assert re_md == complex_md
new_slide.delete()

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
