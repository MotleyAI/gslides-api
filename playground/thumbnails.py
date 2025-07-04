# This is how to get thumbnails from slides

from gslides_api import Presentation, initialize_credentials
from gslides_api.domain import ThumbnailSize


presentation = Presentation.from_id(presentation_id, api_client=api_client)

slide = presentation.slides[7]
thumbnail = slide.thumbnail(ThumbnailSize.LARGE)
print(thumbnail.mime_type)
thumbnail.save("slide_thumbnail.png")
