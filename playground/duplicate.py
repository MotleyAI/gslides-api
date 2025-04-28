import logging
from typing import Any

import gslides
from gslides import creds as re_creds

from gslides_api import Presentation, initialize_credentials


presentation_id = "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

# If modifying these scopes, delete the file token.json.

credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"


logger = logging.getLogger(__name__)

initialize_credentials(credential_location)

presentation = Presentation.from_id(presentation_id)
print(f"\nSuccessfully loaded presentation with ID: {presentation.presentationId}")
print(f"Number of slides: {len(presentation.slides)}")
print(
    f"First slide title: {presentation.slides[0].pageElements[0].shape.text.textElements[1].textRun.content if presentation.slides[0].pageElements[0].shape.text else 'No title'}"
)


new_p = Presentation.create_blank("New Presentation")
print("New presentation created successfully")
print(f"https://docs.google.com/presentation/d/{new_p.presentationId}/edit")
new_p.slides[0].delete()
new_p.sync_from_cloud()


# for i, slide in enumerate(presentation.slides):
#     if i > 13:
#         slide.write_copy(presentation_id=new_p.presentationId)

print("yay!")
