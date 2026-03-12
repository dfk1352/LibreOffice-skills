"""Shared helpers for Writer tests."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from PIL import Image

from uno_bridge import uno_context


@contextmanager
def open_uno_doc(doc_path: Path | str) -> Iterator[Any]:
    """Open a Writer document through UNO for test assertions."""
    path = Path(doc_path)
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            yield doc
        finally:
            doc.close(True)


def create_test_image(
    path: Path,
    size: tuple[int, int] = (40, 40),
    color: str = "red",
) -> Path:
    """Create a small raster image for Writer image tests."""
    image = Image.new("RGB", size, color=color)
    image.save(path)
    return path


def get_table_names(doc_path: Path | str) -> list[str]:
    """Return Writer text table names in document order."""
    with open_uno_doc(doc_path) as doc:
        tables = doc.getTextTables()
        return list(tables.getElementNames())


def get_table_dimensions(doc_path: Path | str, table_name: str) -> tuple[int, int]:
    """Return row and column counts for a named table."""
    with open_uno_doc(doc_path) as doc:
        table = doc.getTextTables().getByName(table_name)
        return (table.Rows.Count, table.Columns.Count)


def get_table_cell_value(doc_path: Path | str, table_name: str, cell_name: str) -> str:
    """Return the string value of a Writer table cell."""
    with open_uno_doc(doc_path) as doc:
        table = doc.getTextTables().getByName(table_name)
        return table.getCellByName(cell_name).getString()


def get_graphic_names(doc_path: Path | str) -> list[str]:
    """Return Writer graphic object names in document order."""
    with open_uno_doc(doc_path) as doc:
        graphics = doc.getGraphicObjects()
        return list(graphics.getElementNames())


def get_graphic_size(doc_path: Path | str, graphic_name: str) -> tuple[int, int]:
    """Return width and height for a named Writer graphic."""
    with open_uno_doc(doc_path) as doc:
        graphic = doc.getGraphicObjects().getByName(graphic_name)
        size = graphic.Size
        return (size.Width, size.Height)
