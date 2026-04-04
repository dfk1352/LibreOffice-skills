# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import struct
import wave
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from PIL import Image

from uno_bridge import uno_context

BLANK_LAYOUT = 20
TITLE_AND_CONTENT_LAYOUT = 1
TITLE_ONLY_LAYOUT = 19
TITLE_SLIDE_LAYOUT = 0
ARABIC_NUMBERING_TYPE = 4
UNORDERED_NUMBERING_STYLE = "List 1"
ORDERED_NUMBERING_STYLE = "Numbering 123"


@contextmanager
def open_impress_doc(doc_path: Path | str) -> Iterator[Any]:
    """Open an Impress document through UNO for test assertions."""
    path = Path(doc_path)
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            yield doc
        finally:
            doc.close(True)


def create_test_image(
    path: Path,
    size: tuple[int, int] = (48, 48),
    color: str = "blue",
) -> Path:
    """Create a small raster image for Impress tests."""
    image = Image.new("RGB", size, color=color)
    image.save(path)
    return path


def create_test_audio(path: Path) -> Path:
    """Create a minimal valid WAV file for Impress media tests."""
    with wave.open(str(path), "w") as wav_file:
        wav_file.setnchannels(1)
        wav_file.setsampwidth(2)
        wav_file.setframerate(44100)
        wav_file.writeframes(struct.pack("<h", 0) * 128)
    return path


def create_test_video(path: Path) -> Path:
    """Create a minimal placeholder file for Impress media tests."""
    path.write_bytes(b"\x00" * 256)
    return path


def append_slide(doc: Any, layout: int = BLANK_LAYOUT) -> Any:
    """Append one slide with the requested layout and return it."""
    pages = doc.DrawPages
    index = pages.Count
    pages.insertNewByIndex(index)
    slide = pages.getByIndex(index)
    slide.Layout = layout
    return slide


def add_text_shape(
    doc: Any,
    slide: Any,
    text: str,
    *,
    name: str | None = None,
    x_cm: float = 1.0,
    y_cm: float = 1.0,
    width_cm: float = 10.0,
    height_cm: float = 3.0,
) -> Any:
    """Add a named text shape to a slide."""
    import uno  # type: ignore[import-not-found]

    shape = doc.createInstance("com.sun.star.drawing.TextShape")
    position = uno.createUnoStruct("com.sun.star.awt.Point")
    position.X = int(x_cm * 1000)
    position.Y = int(y_cm * 1000)
    shape.Position = position

    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width = int(width_cm * 1000)
    size.Height = int(height_cm * 1000)
    shape.Size = size
    slide.add(shape)
    shape.setString(text)
    if name is not None:
        _assign_shape_name(shape, name)
    return shape


def set_shape_text(shape: Any, text: str) -> None:
    """Replace the full text content of a shape."""
    shape.setString(text)


def set_notes_text(slide: Any, text: str) -> None:
    """Replace the speaker notes text for a slide."""
    notes_shape = find_notes_text_shape(slide.getNotesPage())
    if notes_shape is None:
        raise AssertionError("Notes text shape not found")
    notes_shape.setString(text)


def apply_list_to_paragraphs(
    text_shape: Any,
    paragraph_texts: list[str],
    *,
    ordered: bool,
) -> None:
    """Apply list styling to matching paragraphs within a text shape."""
    enumeration = text_shape.getText().createEnumeration()
    while enumeration.hasMoreElements():
        paragraph = enumeration.nextElement()
        if paragraph.getString() not in paragraph_texts:
            continue
        cursor = _paragraph_cursor(paragraph)
        cursor.NumberingLevel = 0
        try:
            cursor.NumberingIsNumber = ordered
        except Exception:
            pass


def get_slide_shapes(doc_path: Path | str, slide_index: int) -> list[dict[str, Any]]:
    """Return shape summaries for one slide."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        summaries = [
            _shape_summary(slide.getByIndex(index), index)
            for index in range(slide.Count)
        ]
        deduped: dict[tuple[str, str], dict[str, Any]] = {}
        ordered: list[dict[str, Any]] = []
        for summary in summaries:
            normalized_name = _normalize_name(summary["name"])
            shape_type = str(summary["type"])
            if not normalized_name:
                ordered.append(summary)
                continue
            key = (normalized_name, shape_type)
            existing = deduped.get(key)
            if existing is None:
                deduped[key] = summary
                ordered.append(summary)
                continue
            existing_score = (
                int(existing["width_cm"] * 1000),
                int(existing["height_cm"] * 1000),
                int(existing["x_cm"] * 1000),
                int(existing["y_cm"] * 1000),
            )
            summary_score = (
                int(summary["width_cm"] * 1000),
                int(summary["height_cm"] * 1000),
                int(summary["x_cm"] * 1000),
                int(summary["y_cm"] * 1000),
            )
            if summary_score > existing_score:
                deduped[key] = summary
                ordered[ordered.index(existing)] = summary
        return ordered


def get_shape_text(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str | None = None,
    index: int | None = None,
    placeholder: str | None = None,
    notes: bool = False,
) -> str:
    """Return the text from one slide or notes shape."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        if notes:
            shape = find_notes_text_shape(slide.getNotesPage())
        else:
            shape = resolve_shape(
                slide, name=name, index=index, placeholder=placeholder
            )
        if shape is None:
            raise AssertionError("Requested shape was not found")
        return _shape_text(shape)


def get_text_properties(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str | None = None,
    placeholder: str | None = None,
) -> dict[str, Any]:
    """Return UNO formatting properties for one text-bearing shape."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        shape = resolve_shape(slide, name=name, placeholder=placeholder)
        if shape is None:
            raise AssertionError("Requested text shape was not found")
        cursor = shape.getText().createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        return {
            "char_weight": cursor.CharWeight,
            "char_posture": _normalize_posture(cursor.CharPosture),
            "char_underline": int(cursor.CharUnderline),
            "font_name": str(cursor.CharFontName),
            "font_size": float(cursor.CharHeight),
            "color": int(cursor.CharColor),
            "align": int(cursor.ParaAdjust),
        }


def _collect_list_metadata(shape: Any) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    enumeration = shape.getText().createEnumeration()
    while enumeration.hasMoreElements():
        paragraph = enumeration.nextElement()
        text = paragraph.getString()
        style_name = _paragraph_numbering_style(paragraph)
        if not style_name:
            continue
        level = _paragraph_numbering_level(paragraph)
        items.append(
            {
                "text": _normalize_list_item_text(text),
                "level": level,
                "style_name": style_name,
                "numbering_type": _get_numbering_type(paragraph, level),
            }
        )
    return items


def get_list_paragraphs(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str | None = None,
    placeholder: str | None = None,
    notes: bool = False,
) -> list[dict[str, Any]]:
    """Return list paragraph metadata from one text-bearing object."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        if notes:
            shape = find_notes_text_shape(slide.getNotesPage())
        else:
            shape = resolve_shape(slide, name=name, placeholder=placeholder)
        if shape is None:
            raise AssertionError("Requested list shape was not found")
        return _collect_list_metadata(shape)


def get_table_matrix(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str,
) -> list[list[str]]:
    """Return table cell strings as a rectangular matrix."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        shape = resolve_shape(slide, name=name)
        if shape is None:
            raise AssertionError(f'Table shape "{name}" not found')
        model = shape.Model
        rows = model.Rows.Count
        cols = model.Columns.Count
        return [
            [model.getCellByPosition(col, row).getString() for col in range(cols)]
            for row in range(rows)
        ]


def get_shape_geometry(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str,
) -> dict[str, int]:
    """Return one shape's raw UNO geometry."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        normalized_expected = _normalize_name(name)
        candidates: list[dict[str, int]] = []
        for index in range(slide.Count):
            shape = slide.getByIndex(index)
            normalized_actual = _normalize_name(_shape_name(shape))
            if (
                normalized_actual != normalized_expected
                and not _matches_uno_duplicate_name(
                    normalized_actual,
                    normalized_expected,
                )
            ):
                continue
            candidates.append(
                {
                    "x": int(getattr(shape.Position, "X", 0)),
                    "y": int(getattr(shape.Position, "Y", 0)),
                    "width": int(getattr(shape.Size, "Width", 0)),
                    "height": int(getattr(shape.Size, "Height", 0)),
                }
            )
        if not candidates:
            raise AssertionError(f'Shape "{name}" not found')
        return max(
            candidates,
            key=lambda candidate: (
                candidate["width"] > 100,
                candidate["height"] > 100,
                candidate["x"],
                candidate["width"],
                candidate["height"],
                candidate["y"],
            ),
        )


def get_chart_details(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str,
) -> dict[str, Any]:
    """Return persisted details for one chart shape."""
    details: dict[str, Any] = get_shape_geometry(doc_path, slide_index, name=name)
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        normalized_expected = _normalize_name(name)
        candidates = []
        for index in range(slide.Count):
            candidate = slide.getByIndex(index)
            normalized_actual = _normalize_name(_shape_name(candidate))
            if (
                normalized_actual != normalized_expected
                and not _matches_uno_duplicate_name(
                    normalized_actual,
                    normalized_expected,
                )
            ):
                continue
            if (
                int(getattr(candidate.Position, "X", 0)) == details["x"]
                and int(getattr(candidate.Size, "Width", 0)) == details["width"]
            ):
                candidates.append(candidate)
        if not candidates:
            raise AssertionError(f'Chart shape "{name}" not found')
        shape = candidates[-1]

        title = None
        try:
            embedded = shape.EmbeddedObject.Component
            if getattr(embedded, "HasMainTitle", False):
                title = str(embedded.Title.String)
        except Exception:
            title = None

        details["title"] = "" if title is None else title
        try:
            data = embedded.getData()
            details["column_descriptions"] = [
                str(value) for value in data.getColumnDescriptions()
            ]
            details["row_descriptions"] = [
                str(value) for value in data.getRowDescriptions()
            ]
            details["data"] = [list(row) for row in data.getData()]
        except Exception:
            details["column_descriptions"] = []
            details["row_descriptions"] = []
            details["data"] = []
        details.update(
            {
                "x": int(shape.Position.X),
                "y": int(shape.Position.Y),
                "width": int(shape.Size.Width),
                "height": int(shape.Size.Height),
            }
        )
        return details


def get_media_url(
    doc_path: Path | str,
    slide_index: int,
    *,
    name: str,
) -> str:
    """Return the stored media or image URL for one named shape."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        matches = []
        normalized_expected = _normalize_name(name)
        for shape_index in range(slide.Count):
            shape = slide.getByIndex(shape_index)
            normalized_actual = _normalize_name(_shape_name(shape))
            if normalized_actual == normalized_expected or _matches_uno_duplicate_name(
                normalized_actual, normalized_expected
            ):
                matches.append(shape)
        if not matches:
            raise AssertionError(f'Media shape "{name}" not found')
        best_url = ""
        best_score = -1
        for shape in matches:
            candidate_url = ""
            description = getattr(shape, "Description", "")
            if description:
                candidate_url = str(description)
            for attribute in ("MediaURL", "PluginURL", "GraphicStreamURL"):
                value = getattr(shape, attribute, "")
                if value:
                    candidate_url = str(value)
            graphic = getattr(shape, "Graphic", None)
            if graphic is not None:
                for attribute in ("OriginURL", "URL"):
                    value = getattr(graphic, attribute, "")
                    if value:
                        candidate_url = str(value)
            if not candidate_url:
                continue
            score = 0
            if "replacement" in candidate_url or "_v2" in candidate_url:
                score += 2
            if candidate_url.endswith(".wav") or candidate_url.endswith(".mp3"):
                score += 2
            if candidate_url.endswith(".png") or candidate_url.endswith(".jpg"):
                score += 1
            if score > best_score:
                best_score = score
                best_url = candidate_url
        if best_url:
            return best_url
        raise AssertionError(f'No media URL found on shape "{name}"')


def get_notes_text(doc_path: Path | str, slide_index: int) -> str:
    """Return the speaker notes text for one slide."""
    return get_shape_text(doc_path, slide_index, notes=True)


def get_slide_master_name(doc_path: Path | str, slide_index: int) -> str:
    """Return the master page name applied to one slide."""
    with open_impress_doc(doc_path) as doc:
        slide = doc.DrawPages.getByIndex(slide_index)
        return str(slide.MasterPage.Name)


def get_master_background(doc_path: Path | str, master_name: str) -> int:
    """Return the fill color for one master page background."""
    with open_impress_doc(doc_path) as doc:
        masters = doc.MasterPages
        for index in range(masters.Count):
            master = masters.getByIndex(index)
            if str(master.Name) == master_name:
                return int(master.Background.FillColor)
    raise AssertionError(f'Master page "{master_name}" not found')


def resolve_shape(
    slide: Any,
    *,
    name: str | None = None,
    index: int | None = None,
    placeholder: str | None = None,
) -> Any | None:
    """Resolve one slide shape by name, index, or placeholder."""
    if name is not None:
        normalized_expected = _normalize_name(name)
        duplicate_match = None
        for shape_index in range(slide.Count):
            shape = slide.getByIndex(shape_index)
            normalized_actual = _normalize_name(_shape_name(shape))
            if normalized_actual == normalized_expected:
                return shape
            if _matches_uno_duplicate_name(normalized_actual, normalized_expected):
                duplicate_match = shape
        return duplicate_match

    if index is not None:
        if index < 0 or index >= slide.Count:
            return None
        return slide.getByIndex(index)

    if placeholder is not None:
        placeholder_key = placeholder.strip().lower()
        for shape_index in range(slide.Count):
            shape = slide.getByIndex(shape_index)
            if placeholder_key == "title" and _is_title_shape(shape):
                return shape
            if placeholder_key == "subtitle" and _is_subtitle_shape(shape):
                return shape
            if placeholder_key == "body" and _is_body_placeholder(shape):
                return shape
        return None

    return None


def find_notes_text_shape(notes_page: Any) -> Any | None:
    """Return the text-bearing notes shape for one notes page."""
    for index in range(notes_page.Count):
        shape = notes_page.getByIndex(index)
        if (
            str(getattr(shape, "ShapeType", ""))
            == "com.sun.star.presentation.NotesShape"
        ):
            return shape
    return None


def _shape_summary(shape: Any, index: int) -> dict[str, Any]:
    return {
        "index": index,
        "name": _shape_name(shape),
        "type": str(getattr(shape, "ShapeType", "")),
        "text": _shape_text(shape),
        "is_presentation": bool(getattr(shape, "IsPresentationObject", False)),
        "x_cm": int(getattr(shape.Position, "X", 0)) / 1000.0,
        "y_cm": int(getattr(shape.Position, "Y", 0)) / 1000.0,
        "width_cm": int(getattr(shape.Size, "Width", 0)) / 1000.0,
        "height_cm": int(getattr(shape.Size, "Height", 0)) / 1000.0,
    }


def _shape_name(shape: Any) -> str:
    if hasattr(shape, "getName"):
        try:
            return str(shape.getName())
        except Exception:
            pass
    return str(getattr(shape, "Name", ""))


def _shape_text(shape: Any) -> str:
    try:
        if _shape_is_text(shape):
            return str(shape.getString())
    except Exception:
        pass
    for attribute in ("String",):
        value = getattr(shape, attribute, "")
        if value:
            return str(value)
    return ""


def _shape_is_text(shape: Any) -> bool:
    try:
        return bool(shape.supportsService("com.sun.star.drawing.Text"))
    except Exception:
        return False


def _is_title_shape(shape: Any) -> bool:
    try:
        return bool(shape.supportsService("com.sun.star.presentation.TitleTextShape"))
    except Exception:
        return False


def _is_subtitle_shape(shape: Any) -> bool:
    try:
        return bool(
            shape.supportsService("com.sun.star.presentation.SubtitleTextShape")
        )
    except Exception:
        return False


def _is_body_placeholder(shape: Any) -> bool:
    if not _shape_is_text(shape):
        return False
    if not bool(getattr(shape, "IsPresentationObject", False)):
        return False
    return not _is_title_shape(shape) and not _is_subtitle_shape(shape)


def _assign_shape_name(shape: Any, name: str) -> None:
    for candidate in (name, "_".join(name.split())):
        if hasattr(shape, "setName"):
            try:
                shape.setName(candidate)
                return
            except Exception:
                pass
        try:
            shape.Name = candidate
            return
        except Exception:
            pass


def _get_numbering_type(paragraph: Any, level: int) -> int | None:
    rules = _paragraph_property(paragraph, "NumberingRules", None)
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


def _normalize_name(value: str) -> str:
    return "_".join(value.split()).lower()


def _normalize_list_item_text(text: str) -> str:
    return text.strip()


def _matches_uno_duplicate_name(actual: str, expected: str) -> bool:
    if not actual.startswith(expected):
        return False
    suffix = actual[len(expected) :]
    if not suffix:
        return False
    return suffix.startswith("_") and suffix[1:].isdigit()


def _paragraph_cursor(paragraph: Any) -> Any:
    text_obj = paragraph.getText()
    cursor = text_obj.createTextCursorByRange(paragraph.getStart())
    cursor.gotoRange(paragraph.getEnd(), True)
    return cursor


def _paragraph_property(paragraph: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(paragraph, name)
    except Exception:
        try:
            return getattr(_paragraph_cursor(paragraph), name)
        except Exception:
            return default


def _paragraph_numbering_style(paragraph: Any) -> str:
    if not _paragraph_is_list_item(paragraph):
        return ""

    value = _paragraph_property(paragraph, "NumberingStyleName", "")
    if value:
        return str(value)
    numbering_type = _get_numbering_type(
        paragraph, _paragraph_numbering_level(paragraph)
    )
    if numbering_type == ARABIC_NUMBERING_TYPE:
        return ORDERED_NUMBERING_STYLE
    if numbering_type is not None:
        return UNORDERED_NUMBERING_STYLE
    return UNORDERED_NUMBERING_STYLE


def _paragraph_numbering_level(paragraph: Any) -> int:
    value = _paragraph_property(paragraph, "NumberingLevel", 0)
    try:
        return int(value)
    except Exception:
        return 0


def _normalize_posture(value: Any) -> int | str:
    normalized = getattr(value, "value", value)
    if normalized == "ITALIC":
        return 2
    if normalized == "NONE":
        return 0
    return normalized


def _paragraph_is_list_item(paragraph: Any) -> bool:
    value = _paragraph_property(paragraph, "NumberingLevel", None)
    if value is None:
        return False
    try:
        return int(value) >= 0
    except Exception:
        return False
