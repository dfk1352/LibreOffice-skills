"""Session-based editing API for Writer documents."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from session import BaseSession
from uno_bridge import uno_context
from writer.core import EXPORT_FILTERS
from writer.exceptions import (
    DocumentNotFoundError,
    ImageNotFoundError,
    InvalidPayloadError,
    InvalidSelectorError,
    InvalidTableError,
    WriterSessionError,
    WriterSkillError,
)
from writer.patch import PatchApplyMode
from writer.selectors import (
    Selector,
    parse_selector,
    resolve_image_selector,
    resolve_table_selector,
    resolve_text_selector,
)


class WriterSession(BaseSession):
    """Long-lived Writer editing session bound to one document."""

    def __init__(self, path: str) -> None:
        super().__init__(closed_error_type=WriterSessionError)
        self._path = Path(path)
        if not self._path.exists():
            raise DocumentNotFoundError(f"Document not found: {path}")

        self._uno_manager: Any = None
        self._desktop: Any = None
        self._doc: Any = None
        self._open_document()

    @property
    def doc(self) -> Any:
        self._require_open()
        return self._doc

    def close(self, save: bool = True) -> None:
        self._require_open()
        try:
            if save:
                self._doc.store()
            self._doc.close(save)
        finally:
            try:
                self._uno_manager.__exit__(None, None, None)
            finally:
                self._closed = True
                self._doc = None
                self._desktop = None
                self._uno_manager = None

    def read_text(self, selector: str | None = None) -> str:
        self._require_open()
        if selector is None:
            return self.doc.Text.getString()
        text_selector = _require_text_selector(selector)
        return resolve_text_selector(text_selector, self.doc).getString()

    def insert_text(self, text: str, selector: str | None = None) -> None:
        self._require_open()
        if selector is not None:
            parsed_selector = _require_text_selector(selector)
            if _should_insert_as_adjacent_paragraph(parsed_selector, text):
                match = resolve_text_selector(
                    Selector("text", "content", parsed_selector.value, "contains"),
                    self.doc,
                )
                _insert_adjacent_paragraph(
                    self.doc.Text,
                    match,
                    text,
                    parsed_selector.operator,
                )
                return
        cursor = self._resolve_insert_cursor(selector)
        _insert_string(self.doc.Text, cursor, text)

    def replace_text(self, selector: str, new_text: str) -> None:
        self._require_open()
        match = resolve_text_selector(_require_text_selector(selector), self.doc)
        match.setString(new_text)

    def delete_text(self, selector: str) -> None:
        self._require_open()
        match = resolve_text_selector(_require_text_selector(selector), self.doc)
        match.setString("")

    def insert_table(
        self,
        rows: int,
        cols: int,
        data: list[list[Any]] | None = None,
        name: str | None = None,
        selector: str | None = None,
    ) -> None:
        self._require_open()
        _validate_table_data(rows, cols, data)

        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(rows, cols)

        cursor = self._resolve_insert_cursor(selector)
        self.doc.Text.insertTextContent(cursor, table, False)
        if name is not None:
            _assign_content_name(table, name)

        if data is not None:
            for row_index, row_data in enumerate(data):
                for col_index, cell_value in enumerate(row_data):
                    table.getCellByName(_get_cell_name(row_index, col_index)).setString(
                        str(cell_value)
                    )

    def update_table(self, selector: str, data: list[list[Any]]) -> None:
        self._require_open()
        table = resolve_table_selector(parse_selector(selector), self.doc)
        rows = table.Rows.Count
        cols = table.Columns.Count
        _validate_table_data(rows, cols, data)

        for row_index, row_data in enumerate(data):
            for col_index, cell_value in enumerate(row_data):
                table.getCellByName(_get_cell_name(row_index, col_index)).setString(
                    str(cell_value)
                )

    def delete_table(self, selector: str) -> None:
        self._require_open()
        table = resolve_table_selector(parse_selector(selector), self.doc)
        try:
            self.doc.Text.removeTextContent(table)
        except Exception:
            table.dispose()

    def insert_image(
        self,
        image_path: str,
        width: int | None = None,
        height: int | None = None,
        name: str | None = None,
        selector: str | None = None,
    ) -> None:
        self._require_open()
        image_file = Path(image_path)
        if not image_file.exists():
            raise ImageNotFoundError(f"Image not found: {image_path}")

        graphic = self.doc.createInstance("com.sun.star.text.GraphicObject")
        graphic.GraphicURL = image_file.resolve().as_uri()
        if width is not None or height is not None:
            _set_graphic_size(graphic, width, height)

        cursor = self._resolve_insert_cursor(selector)
        self.doc.Text.insertTextContent(cursor, graphic, False)
        if name is not None:
            _assign_content_name(graphic, name)

    def update_image(
        self,
        selector: str,
        image_path: str | None = None,
        width: int | None = None,
        height: int | None = None,
    ) -> None:
        self._require_open()
        if image_path is None and width is None and height is None:
            raise InvalidPayloadError("update_image requires image_path and/or size")

        graphic = resolve_image_selector(parse_selector(selector), self.doc)
        if image_path is not None:
            image_file = Path(image_path)
            if not image_file.exists():
                raise ImageNotFoundError(f"Image not found: {image_path}")
            graphic.GraphicURL = image_file.resolve().as_uri()
        if width is not None or height is not None:
            _set_graphic_size(graphic, width, height)

    def delete_image(self, selector: str) -> None:
        self._require_open()
        graphic = resolve_image_selector(parse_selector(selector), self.doc)
        try:
            self.doc.Text.removeTextContent(graphic)
        except Exception:
            graphic.dispose()

    def patch(self, patch_text: str, mode: PatchApplyMode = "atomic"):
        self._require_open()
        from writer.patch import apply_operations

        return apply_operations(self, patch_text, mode)

    def export(self, output_path: str, format: str) -> None:
        """Export the current document state to another Writer-supported format."""
        self._require_open()
        if format not in EXPORT_FILTERS:
            raise WriterSkillError(f"Unsupported export format: {format}")

        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)

        import uno  # type: ignore[import-not-found]

        filter_prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        filter_prop.Name = "FilterName"
        filter_prop.Value = EXPORT_FILTERS[format]
        self.doc.storeToURL(output.resolve().as_uri(), (filter_prop,))

    def reset(self) -> None:
        """Discard in-memory changes and reopen the backing document."""
        self._require_open()
        self._doc.close(False)
        self._uno_manager.__exit__(None, None, None)
        self._open_document()

    def _resolve_insert_cursor(self, selector: str | None) -> Any:
        text_obj = self.doc.Text
        if selector is None:
            cursor = text_obj.createTextCursor()
            cursor.gotoEnd(False)
            return cursor
        return resolve_text_selector(_require_text_selector(selector), self.doc)

    def _open_document(self) -> None:
        self._uno_manager = uno_context()
        self._desktop = self._uno_manager.__enter__()
        try:
            self._doc = self._desktop.loadComponentFromURL(
                self._path.resolve().as_uri(),
                "_blank",
                0,
                (),
            )
        except Exception as exc:
            self._uno_manager.__exit__(type(exc), exc, exc.__traceback__)
            self._uno_manager = None
            self._desktop = None
            raise WriterSkillError(
                f"Failed to open Writer document: {self._path}"
            ) from exc


def open_writer_session(path: str) -> WriterSession:
    """Open a Writer editing session for an existing document."""
    return WriterSession(path)


def _require_text_selector(raw: str) -> Selector:
    selector = parse_selector(raw)
    if selector.kind != "text":
        raise InvalidSelectorError(f"Selector does not target text: {raw}")
    return selector


def _insert_string(text_obj: Any, cursor: Any, text: str) -> None:
    parts = text.split("\n")
    for index, part in enumerate(parts):
        if index > 0:
            text_obj.insertControlCharacter(cursor, 0, False)
        if part:
            text_obj.insertString(cursor, part, False)


def _insert_adjacent_paragraph(
    text_obj: Any,
    match: Any,
    text: str,
    operator: str,
) -> None:
    if operator == "after":
        cursor = text_obj.createTextCursorByRange(match.End)
        text_obj.insertControlCharacter(cursor, 0, False)
        _insert_string(text_obj, cursor, text)
        text_obj.insertControlCharacter(cursor, 0, False)
        return

    cursor = text_obj.createTextCursorByRange(match.Start)
    _insert_string(text_obj, cursor, text)
    text_obj.insertControlCharacter(cursor, 0, False)


def _validate_table_data(
    rows: int,
    cols: int,
    data: list[list[Any]] | None,
) -> None:
    if rows < 1 or cols < 1:
        raise InvalidTableError("Rows and cols must be >= 1")
    if data is None:
        return
    if len(data) != rows:
        raise InvalidTableError(f"Data has {len(data)} rows but table needs {rows}")
    for row_index, row in enumerate(data):
        if len(row) != cols:
            raise InvalidTableError(
                f"Data row {row_index} has {len(row)} cols but table needs {cols}"
            )


def _should_insert_as_adjacent_paragraph(selector: Selector, text: str) -> bool:
    return (
        selector.operator in {"after", "before"}
        and not text.startswith("\n")
        and not text.endswith("\n")
    )


def _get_cell_name(row: int, col: int) -> str:
    col_name = ""
    col_num = col + 1
    while col_num > 0:
        col_num -= 1
        col_name = chr(65 + (col_num % 26)) + col_name
        col_num //= 26
    return f"{col_name}{row + 1}"


def _set_graphic_size(graphic: Any, width: int | None, height: int | None) -> None:
    import uno  # type: ignore[import-not-found]

    current = graphic.Size
    size = uno.createUnoStruct("com.sun.star.awt.Size")
    size.Width = current.Width if width is None else width
    size.Height = current.Height if height is None else height
    graphic.setSize(size)


def _assign_content_name(content: Any, name: str) -> None:
    candidates = [name]
    normalized = "_".join(name.split())
    if normalized != name:
        candidates.append(normalized)

    for candidate in candidates:
        if hasattr(content, "setName"):
            try:
                content.setName(candidate)
                return
            except Exception:
                pass
        try:
            content.Name = candidate
            return
        except Exception:
            pass
