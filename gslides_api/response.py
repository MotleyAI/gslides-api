import os
import requests
from io import BytesIO

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
        """Detect the image format from the payload.

        Uses Pillow (PIL) as the primary method, with a fallback to basic
        header detection for common formats.
        """
        try:
            from PIL import Image
            with BytesIO(self.payload) as img_buffer:
                with Image.open(img_buffer) as img:
                    # PIL format names to standard format names
                    format_mapping = {
                        'JPEG': 'jpeg',
                        'PNG': 'png',
                        'GIF': 'gif',
                        'BMP': 'bmp',
                        'WEBP': 'webp',
                        'TIFF': 'tiff'
                    }
                    return format_mapping.get(img.format, img.format.lower() if img.format else None)
        except ImportError:
            # Fallback: basic header detection for common formats
            return self._detect_format_from_header()
        except Exception:
            # If PIL fails to open the image, try fallback
            return self._detect_format_from_header()

    def _detect_format_from_header(self):
        """Fallback method to detect image format from file headers."""
        if not self.payload:
            return None

        # Check common image format headers
        if self.payload.startswith(b'\xff\xd8\xff'):
            return 'jpeg'
        elif self.payload.startswith(b'\x89PNG\r\n\x1a\n'):
            return 'png'
        elif self.payload.startswith(b'GIF87a') or self.payload.startswith(b'GIF89a'):
            return 'gif'
        elif self.payload.startswith(b'BM'):
            return 'bmp'
        elif self.payload.startswith(b'RIFF') and b'WEBP' in self.payload[:12]:
            return 'webp'
        elif self.payload.startswith(b'II*\x00') or self.payload.startswith(b'MM\x00*'):
            return 'tiff'
        else:
            return None

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
