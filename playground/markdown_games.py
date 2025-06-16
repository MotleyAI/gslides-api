import os

import marko

from gslides_api import Presentation, initialize_credentials


here = os.path.dirname(os.path.abspath(__file__))
credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
initialize_credentials(credential_location)

md = """
# Title
Video Analysis Report for {client_name} - Period ending {end_date}
## Subtitle *yes* **yes**
This is a *very important* report.
* It illustrates **bullet points**
* And *deep* bullet points
* And *more* bullet points
* And even `code` blocks
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
print("Kind of a copy written!")
