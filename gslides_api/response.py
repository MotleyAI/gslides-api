import os
import requests
import imghdr

from gslides_api.domain import GSlidesBaseModel


class ImageThumbnail(GSlidesBaseModel):
    """Represents a response to an image request received from
    https://developers.google.com/workspace/slides/api/reference/rest/v1/presentations.pages/getThumbnail
    """

    contentUrl: str
    width: int
    height: int
    _payload: bytes = None

    @property
    def payload(self):
        if self._payload is None:
            self._payload = requests.get(self.contentUrl).content
        return self._payload

    @property
    def mime_type(self):
        return imghdr.what(None, h=self.payload)

    def save(self, file_path: str):
        # Get file extension and convert to expected format name
        file_extension = os.path.splitext(file_path)[1].lower().lstrip(".")

        if file_extension:
            # Detect the actual image format from the payload

            # Handle common extension aliases
            expected_format = "jpeg" if file_extension in ("jpg", "jpeg") else file_extension

            if self.mime_type and self.mime_type != expected_format:
                raise ValueError(
                    f"Image format mismatch: file extension '.{file_extension}' suggests "
                    f"'{expected_format}' format, but payload contains '{self.mime_type}' format"
                )

        with open(file_path, "wb") as f:
            f.write(self.payload)

    def to_ipython_image(self):
        try:
            from IPython.display import Image
        except ImportError:
            raise ImportError("IPython is not installed. Please install it to use this method.")
        from IPython.display import Image

        return Image(self.payload)
