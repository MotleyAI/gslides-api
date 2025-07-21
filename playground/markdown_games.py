import json
import os

from gslides_api import Presentation, Slide, initialize_credentials

here = os.path.dirname(os.path.abspath(__file__))
credential_location = os.getenv("GSLIDES_CREDENTIALS_PATH")
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

# Test write medium markdown
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

# Test write complicated markdown
new_slide = s.duplicate()
new_slide.get_element_by_alt_title("text_1").write_text(complex_md, as_markdown=True)
new_slide.sync_from_cloud()
re_md = new_slide.get_element_by_alt_title("text_1").read_text()
assert re_md == complex_md
new_slide.delete()
