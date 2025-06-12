from gslides_api import Presentation, initialize_credentials
from gslides_api.execute import copy_presentation
from gslides_api.json_diff import json_diff

credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
presentation_id = "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

initialize_credentials(credential_location)

presentation = Presentation.from_id(presentation_id)

clone = presentation.clone()
copy = copy_presentation(presentation_id, "Copy of Presentation")
print(clone.url)


def check_duplication(p: Presentation, index: int, duplicate: bool = False) -> list:
    slide = p.slides[index]
    if duplicate:
        # This passed all checks
        new_slide = slide.duplicate()
    else:
        new_slide = slide.write_copy(insertion_index=index + 1)
    diff = json_diff(
        slide.to_api_format(),
        new_slide.to_api_format(),
        ignored_keys=["objectId", "speakerNotesObjectId", "contentUrl"],
    )
    if len(diff) > 0:
        print(diff)
    new_slide.delete()
    return diff


for i in range(len(presentation.slides)):
    assert len(check_duplication(presentation, i)) == 0


print("Slide written successfully")
