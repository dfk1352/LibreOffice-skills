"""Shared helpers for Writer tests."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from PIL import Image

from uno_bridge import uno_context

ARABIC_NUMBERING_TYPE = 4
BULLET_NUMBERING_TYPE = 6


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


def get_text_properties(doc_path: Path | str, text: str) -> dict[str, Any]:
    """Return UNO formatting properties for one unique matched text span."""
    with open_uno_doc(doc_path) as doc:
        match = _require_unique_text_match(doc, text)
        cursor = doc.Text.createTextCursorByRange(match.Start)
        cursor.gotoRange(match.End, True)
        return {
            "char_weight": cursor.CharWeight,
            "char_posture": str(cursor.CharPosture),
            "char_underline": cursor.CharUnderline,
            "font_name": cursor.CharFontName,
            "font_size": cursor.CharHeight,
            "color": cursor.CharColor,
            "align": cursor.ParaAdjust,
        }


def assert_text_formatting(doc_path: Path | str, text: str, **expected: Any) -> None:
    """Assert UNO formatting properties on one unique text span."""
    properties = get_text_properties(doc_path, text)
    for key, value in expected.items():
        assert properties[key] == value


def get_list_paragraphs(doc_path: Path | str) -> list[dict[str, Any]]:
    """Return list paragraph metadata in document order."""
    with open_uno_doc(doc_path) as doc:
        paragraphs = doc.Text.createEnumeration()
        items: list[dict[str, Any]] = []
        while paragraphs.hasMoreElements():
            paragraph = paragraphs.nextElement()
            style_name = getattr(paragraph, "NumberingStyleName", "")
            if not style_name:
                continue
            level = int(getattr(paragraph, "NumberingLevel", 0))
            items.append(
                {
                    "text": paragraph.getString(),
                    "level": level,
                    "style_name": style_name,
                    "numbering_type": _get_numbering_type(paragraph, level),
                }
            )
        return items


def assert_list_items(
    doc_path: Path | str,
    expected_texts: list[str],
    *,
    expected_levels: list[int] | None = None,
    expected_numbering_type: int | None = None,
) -> None:
    """Assert list paragraph structure in document order."""
    items = get_list_paragraphs(doc_path)
    assert [item["text"] for item in items] == expected_texts
    if expected_levels is not None:
        assert [item["level"] for item in items] == expected_levels
    if expected_numbering_type is not None:
        assert all(item["numbering_type"] == expected_numbering_type for item in items)


def _require_unique_text_match(doc: Any, text: str) -> Any:
    matches = _find_text_matches(doc, text)
    assert len(matches) == 1, f'Expected exactly one match for "{text}"'
    return matches[0]


def _find_text_matches(doc: Any, needle: str) -> list[Any]:
    search = doc.createSearchDescriptor()
    search.SearchString = needle
    matches: list[Any] = []
    found = doc.findFirst(search)
    while found is not None:
        matches.append(found)
        found = doc.findNext(found.End, search)
    return matches


def _get_numbering_type(paragraph: Any, level: int) -> int | None:
    rules = getattr(paragraph, "NumberingRules", None)
    if rules is None:
        return None
    try:
        properties = rules.getByIndex(level)
    except Exception:
        try:
            properties = rules.getByIndex(0)
        except Exception:
            return None

    for property_value in properties:
        if getattr(property_value, "Name", None) == "NumberingType":
            try:
                return int(property_value.Value)
            except Exception:
                return None
    return None
