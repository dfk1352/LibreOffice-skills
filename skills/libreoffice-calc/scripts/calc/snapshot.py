"""Calc snapshot module for area-level PNG export."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from dataclasses import dataclass
from pathlib import Path

from calc.exceptions import (
    DocumentNotFoundError,
    FilterError,
    InvalidAreaError,
    InvalidSheetError,
    SnapshotError,
)
from uno_bridge import uno_context


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


def snapshot_area(
    doc_path: str,
    output_path: str,
    sheet: str = "Sheet1",
    row: int = 0,
    col: int = 0,
    width: int | None = None,
    height: int | None = None,
    dpi: int = 150,
) -> SnapshotResult:
    """Capture a cell-anchored area from a Calc spreadsheet as PNG.

    Args:
        doc_path: Path to the Calc spreadsheet.
        output_path: File path for the PNG output.
        sheet: Name of the sheet to capture.
        row: Zero-based row of the top-left anchor cell.
        col: Zero-based column of the top-left anchor cell.
        width: Pixel width of the capture area (None for full sheet).
        height: Pixel height of the capture area (None for full sheet).
        dpi: Export resolution in dots per inch.

    Returns:
        SnapshotResult with file path, dimensions, and dpi.

    Raises:
        DocumentNotFoundError: If the document file does not exist.
        InvalidSheetError: If the sheet name does not exist.
        InvalidAreaError: If row, col, width, or height are negative.
        FilterError: If the PNG export fails.
    """
    file_path = Path(doc_path)
    if not file_path.exists():
        raise DocumentNotFoundError(f"Document not found: {doc_path}")

    if row < 0:
        raise InvalidAreaError(f"Row must be >= 0, got {row}")
    if col < 0:
        raise InvalidAreaError(f"Col must be >= 0, got {col}")
    if width is not None and width < 0:
        raise InvalidAreaError(f"Width must be >= 0, got {width}")
    if height is not None and height < 0:
        raise InvalidAreaError(f"Height must be >= 0, got {height}")

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with uno_context() as desktop:
        import uno

        doc = desktop.loadComponentFromURL(
            file_path.resolve().as_uri(), "_blank", 0, ()
        )
        if doc is None:
            raise SnapshotError(f"Failed to open document for snapshot: {doc_path}")
        try:
            sheets = doc.Sheets
            if not sheets.hasByName(sheet):
                available = [sheets.getByIndex(i).Name for i in range(sheets.Count)]
                raise InvalidSheetError(
                    f"Sheet '{sheet}' not found. Available: {available}"
                )

            sheet_obj = sheets.getByName(sheet)
            controller = doc.getCurrentController()
            controller.setActiveSheet(sheet_obj)

            if width is not None and height is not None:
                # Determine cell range using actual column/row geometry
                end_col = col
                end_row = row
                try:
                    accum_w = 0
                    # Convert pixel width to 1/100 mm (assume 96 dpi)
                    target_w_mm100 = width * 2540.0 / 96.0
                    for c in range(col, col + 200):
                        cw = sheet_obj.Columns.getByIndex(c).Width
                        accum_w += cw
                        end_col = c
                        if accum_w >= target_w_mm100:
                            break
                    accum_h = 0
                    target_h_mm100 = height * 2540.0 / 96.0
                    for r in range(row, row + 500):
                        rh = sheet_obj.Rows.getByIndex(r).Height
                        accum_h += rh
                        end_row = r
                        if accum_h >= target_h_mm100:
                            break
                except Exception:
                    # Fallback heuristic
                    end_row = row + max(1, height // 20)
                    end_col = col + max(1, width // 80)
                cell_range = sheet_obj.getCellRangeByPosition(
                    col, row, end_col, end_row
                )
            else:
                cell_range = sheet_obj.getCellRangeByPosition(
                    col, row, max(col + 10, col), max(row + 20, row)
                )
                try:
                    cursor = sheet_obj.createCursor()
                    cursor.gotoStartOfUsedArea(False)
                    cursor.gotoEndOfUsedArea(True)
                    used = cursor.getRangeAddress()
                    end_col = max(col, int(used.EndColumn))
                    end_row = max(row, int(used.EndRow))
                    cell_range = sheet_obj.getCellRangeByPosition(
                        col,
                        row,
                        end_col,
                        end_row,
                    )
                except Exception:
                    pass
            controller.select(cell_range)

            pixel_w = width if width is not None else int(dpi * 8.27)
            pixel_h = height if height is not None else int(dpi * 11.69)

            filter_data = []

            fd_width = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fd_width.Name = "PixelWidth"
            fd_width.Value = pixel_w
            filter_data.append(fd_width)

            fd_height = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            fd_height.Name = "PixelHeight"
            fd_height.Value = pixel_h
            filter_data.append(fd_height)

            props = []

            p_filter = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p_filter.Name = "FilterName"
            p_filter.Value = "calc_png_Export"
            props.append(p_filter)

            p_data = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p_data.Name = "FilterData"
            p_data.Value = uno.Any(
                "[]com.sun.star.beans.PropertyValue", tuple(filter_data)
            )
            props.append(p_data)

            p_sel = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            p_sel.Name = "Selection"
            p_sel.Value = cell_range
            props.append(p_sel)

            try:
                doc.storeToURL(output.resolve().as_uri(), tuple(props))
            except Exception as e:
                raise FilterError(f"PNG export failed: {e}") from e

        finally:
            doc.close(True)

    actual_width, actual_height = _read_png_dimensions(output)

    return SnapshotResult(
        file_path=str(output),
        width=actual_width,
        height=actual_height,
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
