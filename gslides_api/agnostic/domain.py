import mimetypes
import os
from typing import Optional

from pydantic import BaseModel


class ImageData(BaseModel):
    """Container for retrieved image data with metadata."""

    content: bytes
    """Raw image data as bytes."""

    mime_type: str
    """MIME type of the image (e.g., 'image/jpeg', 'image/png')."""

    filename: Optional[str] = None
    """Optional filename hint for saving."""

    @classmethod
    def from_file(cls, path: str):
        """Read image data from a file."""
        with open(path, "rb") as f:
            content = f.read()
        mime_type = mimetypes.guess_type(path)[0]
        return cls(content=content, mime_type=mime_type, filename=os.path.abspath(path))

    def save_to_file(self, path: str) -> str:
        """Save image data to a file.

        Args:
            path: File path to save to. If it's a directory, uses filename hint.

        Returns:
            str: The actual path where the file was saved.

        Raises:
            ValueError: If path is a directory but no filename hint is available.
            OSError: If file cannot be written.
        """
        if os.path.isdir(path):
            if not self.filename:
                # Generate filename from MIME type
                ext = mimetypes.guess_extension(self.mime_type) or ".bin"
                filename = f"image{ext}"
            else:
                filename = self.filename
            file_path = os.path.join(path, filename)
        else:
            file_path = path

        # Ensure parent directory exists
        parent_dir = os.path.dirname(file_path)
        if parent_dir:
            try:
                os.makedirs(parent_dir, exist_ok=True)
            except OSError as e:
                raise OSError(f"Cannot create directory {parent_dir}: {e}") from e

        # Write the file
        try:
            with open(file_path, "wb") as f:
                f.write(self.content)
        except OSError as e:
            raise OSError(f"Cannot write file {file_path}: {e}") from e

        return file_path

    def get_extension(self) -> str:
        """Get file extension based on MIME type."""
        ext = mimetypes.guess_extension(self.mime_type)
        return ext or ".bin"
