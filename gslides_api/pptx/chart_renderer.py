"""Render PowerPoint charts to images using LibreOffice.

This module provides functionality to extract chart images from native PowerPoint
charts (GraphicFrame elements with embedded chart data) that cannot be directly
exported as images via python-pptx.
"""

import asyncio
import io
import logging
import os
import shutil
import subprocess
import tempfile
from typing import Optional, Tuple

from PIL import Image
from pptx import Presentation
from pptx.util import Emu

from gslides_api.agnostic.domain import ImageData

from gslides_api.common.log_time import log_time

logger = logging.getLogger(__name__)


@log_time
def render_slide_to_image(
    presentation_path: str,
    slide_index: int,
    crop_bounds: Optional[Tuple[int, int, int, int]] = None,
) -> Optional[bytes]:
    """Render a single slide to a PNG image using LibreOffice.

    Args:
        presentation_path: Path to the PPTX file
        slide_index: Zero-based index of the slide to render
        crop_bounds: Optional (left, top, right, bottom) in EMUs to crop the image

    Returns:
        PNG image bytes, or None if rendering fails
    """
    soffice_path = shutil.which("soffice")
    if not soffice_path:
        logger.warning("LibreOffice (soffice) not available for chart rendering")
        return None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # First, convert PPTX to PDF (LibreOffice only exports first slide to PNG)
            pdf_result = subprocess.run(
                [
                    soffice_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    tmpdir,
                    presentation_path,
                ],
                capture_output=True,
                timeout=60,
            )

            if pdf_result.returncode != 0:
                logger.warning(f"LibreOffice PDF conversion failed: {pdf_result.stderr.decode()}")
                return None

            # Find the PDF file
            base_name = os.path.splitext(os.path.basename(presentation_path))[0]
            pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")

            if not os.path.exists(pdf_path):
                logger.warning("LibreOffice produced no PDF output")
                return None

            # Use pdftoppm to extract the specific page as PNG
            pdftoppm_path = shutil.which("pdftoppm")
            if pdftoppm_path:
                png_prefix = os.path.join(tmpdir, "slide")
                page_num = slide_index + 1  # pdftoppm uses 1-based indexing

                pdftoppm_result = subprocess.run(
                    [
                        pdftoppm_path,
                        "-png",
                        "-f",
                        str(page_num),
                        "-l",
                        str(page_num),
                        "-r",
                        "150",  # 150 DPI for good quality
                        pdf_path,
                        png_prefix,
                    ],
                    capture_output=True,
                    timeout=30,
                )

                if pdftoppm_result.returncode != 0:
                    logger.warning(f"pdftoppm conversion failed: {pdftoppm_result.stderr.decode()}")
                    return None

                # Find the output PNG (pdftoppm adds page number suffix)
                png_files = [
                    f for f in os.listdir(tmpdir) if f.startswith("slide") and f.endswith(".png")
                ]
                if not png_files:
                    logger.warning("pdftoppm produced no PNG output")
                    return None

                png_path = os.path.join(tmpdir, png_files[0])
            else:
                # Fallback: use ImageMagick convert
                convert_path = shutil.which("convert")
                if not convert_path:
                    logger.warning("Neither pdftoppm nor ImageMagick available for PDF to PNG")
                    return None

                png_path = os.path.join(tmpdir, "slide.png")
                convert_result = subprocess.run(
                    [
                        convert_path,
                        "-density",
                        "150",
                        f"{pdf_path}[{slide_index}]",  # Page number in square brackets
                        png_path,
                    ],
                    capture_output=True,
                    timeout=30,
                )

                if convert_result.returncode != 0:
                    logger.warning(
                        f"ImageMagick conversion failed: {convert_result.stderr.decode()}"
                    )
                    return None

            if not os.path.exists(png_path):
                logger.warning("No PNG file produced")
                return None

            # Read and optionally crop the image
            with Image.open(png_path) as img:
                if crop_bounds:
                    # Convert EMU bounds to pixel coordinates
                    # EMU to inches: 914400 EMU per inch
                    # Then multiply by DPI (150)
                    dpi = 150
                    emu_per_inch = 914400

                    left_px = int(crop_bounds[0] / emu_per_inch * dpi)
                    top_px = int(crop_bounds[1] / emu_per_inch * dpi)
                    right_px = int(crop_bounds[2] / emu_per_inch * dpi)
                    bottom_px = int(crop_bounds[3] / emu_per_inch * dpi)

                    # Ensure bounds are within image
                    left_px = max(0, left_px)
                    top_px = max(0, top_px)
                    right_px = min(img.width, right_px)
                    bottom_px = min(img.height, bottom_px)

                    img = img.crop((left_px, top_px, right_px, bottom_px))

                # Convert to PNG bytes
                png_buffer = io.BytesIO()
                img.save(fp=png_buffer, format="PNG")
                return png_buffer.getvalue()

    except subprocess.TimeoutExpired:
        logger.warning("LibreOffice/conversion command timed out")
        return None
    except Exception as e:
        logger.warning(f"Chart rendering failed: {e}")
        return None


@log_time
async def render_all_slides_to_images(
    presentation_path: str,
    dpi: int = 150,
) -> list[bytes]:
    """Render all slides to PNG images using LibreOffice (async).

    This is more efficient than calling render_slide_to_image() for each slide
    because it converts the PPTX to PDF only once and extracts all pages at once.

    Args:
        presentation_path: Path to the PPTX file
        dpi: Resolution for rendering (default 150)

    Returns:
        List of PNG image bytes, one per slide. Empty list if rendering fails.
    """
    soffice_path = shutil.which("soffice")
    if not soffice_path:
        logger.warning("LibreOffice (soffice) not available for slide rendering")
        return []

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Convert PPTX to PDF using async subprocess
            pdf_process = await asyncio.create_subprocess_exec(
                soffice_path,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                presentation_path,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )

            try:
                _, stderr = await asyncio.wait_for(pdf_process.communicate(), timeout=60)
            except asyncio.TimeoutError:
                pdf_process.kill()
                logger.warning("LibreOffice PDF conversion timed out")
                return []

            if pdf_process.returncode != 0:
                logger.warning(f"LibreOffice PDF conversion failed: {stderr.decode()}")
                return []

            # Find the PDF file
            base_name = os.path.splitext(os.path.basename(presentation_path))[0]
            pdf_path = os.path.join(tmpdir, f"{base_name}.pdf")

            if not os.path.exists(pdf_path):
                logger.warning("LibreOffice produced no PDF output")
                return []

            # Use pdftoppm to extract ALL pages as PNG (no -f/-l flags)
            pdftoppm_path = shutil.which("pdftoppm")
            if not pdftoppm_path:
                logger.warning("pdftoppm not available for PDF to PNG conversion")
                return []

            png_prefix = os.path.join(tmpdir, "slide")

            pdftoppm_process = await asyncio.create_subprocess_exec(
                pdftoppm_path,
                "-png",
                "-r",
                str(dpi),
                pdf_path,
                png_prefix,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )

            try:
                _, stderr = await asyncio.wait_for(pdftoppm_process.communicate(), timeout=60)
            except asyncio.TimeoutError:
                pdftoppm_process.kill()
                logger.warning("pdftoppm conversion timed out")
                return []

            if pdftoppm_process.returncode != 0:
                logger.warning(f"pdftoppm conversion failed: {stderr.decode()}")
                return []

            # Find all output PNGs (pdftoppm names them slide-01.png, slide-02.png, etc.)
            png_files = sorted(
                [f for f in os.listdir(tmpdir) if f.startswith("slide") and f.endswith(".png")]
            )

            if not png_files:
                logger.warning("pdftoppm produced no PNG output")
                return []

            # Read all PNG files
            png_bytes_list = []
            for png_file in png_files:
                png_path = os.path.join(tmpdir, png_file)
                with open(png_path, "rb") as f:
                    png_bytes_list.append(f.read())

            logger.info(f"Rendered {len(png_bytes_list)} slides from {presentation_path}")
            return png_bytes_list

    except Exception as e:
        logger.warning(f"Batch slide rendering failed: {e}")
        return []


@log_time
def render_chart_element_to_image(
    presentation: Presentation,
    slide_index: int,
    shape_left: int,
    shape_top: int,
    shape_width: int,
    shape_height: int,
) -> Optional[ImageData]:
    """Render a chart shape from a presentation to an ImageData object.

    This saves the presentation to a temp file, renders the slide,
    and crops to the shape's bounding box.

    Args:
        presentation: python-pptx Presentation object
        slide_index: Zero-based index of the slide containing the chart
        shape_left: Shape left position in EMUs
        shape_top: Shape top position in EMUs
        shape_width: Shape width in EMUs
        shape_height: Shape height in EMUs

    Returns:
        ImageData with PNG bytes, or None if rendering fails
    """
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            temp_path = tmp.name

        # Save presentation to temp file
        presentation.save(temp_path)

        # Calculate crop bounds
        crop_bounds = (
            shape_left,
            shape_top,
            shape_left + shape_width,
            shape_top + shape_height,
        )

        # Render and crop
        png_bytes = render_slide_to_image(
            presentation_path=temp_path,
            slide_index=slide_index,
            crop_bounds=crop_bounds,
        )

        if png_bytes is None:
            return None

        return ImageData(
            content=png_bytes,
            mime_type="image/png",
        )

    except Exception as e:
        logger.warning(f"Failed to render chart element: {e}")
        return None
    finally:
        # Clean up temp file
        if temp_path and os.path.exists(temp_path):
            os.unlink(temp_path)


@log_time
def render_pptx_chart_to_image(
    pptx_element: "GraphicFrame",
    presentation: Presentation,
    slide_index: int,
) -> Optional[ImageData]:
    """Render a native PowerPoint chart to an ImageData object.

    Convenience wrapper that extracts position/size from the shape.

    Args:
        pptx_element: python-pptx GraphicFrame containing the chart
        presentation: python-pptx Presentation object
        slide_index: Zero-based index of the slide containing the chart

    Returns:
        ImageData with PNG bytes, or None if rendering fails
    """
    return render_chart_element_to_image(
        presentation=presentation,
        slide_index=slide_index,
        shape_left=pptx_element.left,
        shape_top=pptx_element.top,
        shape_width=pptx_element.width,
        shape_height=pptx_element.height,
    )


@log_time
def render_chart_from_file(
    presentation_path: str,
    slide_index: int,
    shape_left: int,
    shape_top: int,
    shape_width: int,
    shape_height: int,
) -> Optional[ImageData]:
    """Render a chart from a PPTX file directly without re-saving.

    This is more efficient than render_chart_element_to_image when you already
    have the presentation file path (e.g., during ingestion).

    Args:
        presentation_path: Path to the PPTX file
        slide_index: Zero-based index of the slide containing the chart
        shape_left: Shape left position in EMUs
        shape_top: Shape top position in EMUs
        shape_width: Shape width in EMUs
        shape_height: Shape height in EMUs

    Returns:
        ImageData with PNG bytes, or None if rendering fails
    """
    crop_bounds = (
        shape_left,
        shape_top,
        shape_left + shape_width,
        shape_top + shape_height,
    )

    png_bytes = render_slide_to_image(
        presentation_path=presentation_path,
        slide_index=slide_index,
        crop_bounds=crop_bounds,
    )

    if png_bytes is None:
        return None

    return ImageData(
        content=png_bytes,
        mime_type="image/png",
    )
