from gslides_api import Presentation, initialize_credentials
from gslides_api.domain.domain import MimeType, ThumbnailProperties, ThumbnailSize

credential_location = "/home/james/Dropbox/PyCharmProjects/gslides-playground/"
presentation_id = "1bj3qEcf1P6NhShY8YC0UyEwpc_bFdrxxtijqz8hBbXM"

initialize_credentials(credential_location)

presentation = Presentation.from_id(presentation_id)

slide = presentation.slides[7]
thumbnail = slide.thumbnail(ThumbnailSize.LARGE)
print(thumbnail.mime_type)
thumbnail.save("slide_thumbnail.png")

new_slide_2 = slide.write_copy(insertion_index=8)
new_slide_2.delete()
new_slide = slide.duplicate()
new_slide.delete()
slide.move(insertion_index=10)

slide.move(insertion_index=7)


print("Slide written successfully")
