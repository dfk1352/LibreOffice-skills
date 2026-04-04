"""Writer snapshot module for page-level PNG export."""

from dataclasses import dataclass
from pathlib import Path

from uno_bridge import uno_context
from writer.exceptions import (
    DocumentNotFoundError,
    FilterError,
    InvalidPageError,
    SnapshotError,
)


@dataclass
class SnapshotResult:
    """Result metadata from a snapshot export.

    Attributes:
        file_path: Path to the exported PNG file.
        width: Image width in pixels.
        height: Image height in pixels.
        dpi: Export resolution in dots per inch.
    """

    file_path: str
    width: int
    height: int
    dpi: int


def snapshot_page(
    doc_path: str,
    output_path: str,
    page: int = 1,
    dpi: int = 150,
) -> SnapshotResult:
    """Capture a specific page from a Writer document as PNG.

    Args:
        doc_path: Path to the Writer document.
        output_path: File path for the PNG output.
        page: 1-indexed page number to capture.
        dpi: Export resolution in dots per inch.

    Returns:
        SnapshotResult with file path, dimensions, and dpi.

    Raises:
        DocumentNotFoundError: If the document file does not exist.
        InvalidPageError: If the page number is out of bounds.
        FilterError: If the PNG export fails.
    """
    file_path = Path(doc_path)
    if not file_path.exists():
        raise DocumentNotFoundError(f"Document not found: {doc_path}")

    if page < 1:
        raise InvalidPageError(f"Page must be >= 1, got {page}")

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with uno_context() as desktop:
        import uno  # type: ignore[import-not-found]

        doc = desktop.loadComponentFromURL(
            file_path.resolve().as_uri(), "_blank", 0, ()
        )
        if doc is None:
            raise SnapshotError(f"Failed to open document for snapshot: {doc_path}")
        try:
            page_count = _page_count(doc)
            if page > page_count:
                raise InvalidPageError(
                    f"Page {page} out of range (document has {page_count} pages)"
                )

            # Read page geometry from document style
            width_inches = 8.27  # A4 fallback
            height_inches = 11.69
            try:
                styles = doc.StyleFamilies.getByName("PageStyles")
                # Try common names for the default page style
                page_style = None
                for name in ("Default Page Style", "Standard"):
                    try:
                        page_style = styles.getByName(name)
                        break
                    except Exception:
                        continue
                if page_style is not None:
                    width_inches = page_style.Width / 2540.0
                    height_inches = page_style.Height / 2540.0
            except Exception:
                pass  # Fall back to A4

            filter_data = []

            fd_width = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fd_width.Name = "PixelWidth"
            fd_width.Value = int(dpi * width_inches)
            filter_data.append(fd_width)

            fd_height = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fd_height.Name = "PixelHeight"
            fd_height.Value = int(dpi * height_inches)
            filter_data.append(fd_height)

            fd_page = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fd_page.Name = "PageRange"
            fd_page.Value = str(page)
            filter_data.append(fd_page)

            props = []

            p_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p_filter.Name = "FilterName"
            p_filter.Value = "writer_png_Export"
            props.append(p_filter)

            p_data = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p_data.Name = "FilterData"
            p_data.Value = uno.Any(
                "[]com.sun.star.beans.PropertyValue", tuple(filter_data)
            )
            props.append(p_data)

            try:
                doc.storeToURL(output.resolve().as_uri(), tuple(props))
            except Exception as e:
                raise FilterError(f"PNG export failed: {e}") from e

        finally:
            doc.close(True)

    width, height = _read_png_dimensions(output)

    return SnapshotResult(
        file_path=str(output),
        width=width,
        height=height,
        dpi=dpi,
    )


def _read_png_dimensions(path: Path) -> tuple[int, int]:
    """Read width and height from a PNG file.

    Args:
        path: Path to the PNG file.

    Returns:
        Tuple of (width, height) in pixels.
    """
    try:
        from PIL import Image

        with Image.open(path) as img:
            return img.size
    except ImportError:
        import struct

        with open(path, "rb") as f:
            f.read(16)
            w, h = struct.unpack(">II", f.read(8))
        return w, h


def _page_count(doc: object) -> int:
    if hasattr(doc, "supportsService") and doc.supportsService(
        "com.sun.star.text.TextDocument"
    ):
        try:
            from typing import cast, Any

            return int(cast(Any, doc).getPropertyValue("PageCount"))
        except Exception:
            pass

    controller = getattr(doc, "getCurrentController", lambda: None)()
    controller_count = getattr(controller, "PageCount", None)
    if isinstance(controller_count, int) and controller_count > 0:
        return controller_count

    try:
        draw_count = int(getattr(doc, "DrawPages").Count)
        if draw_count > 0:
            return draw_count
    except Exception:
        pass

    return 1
