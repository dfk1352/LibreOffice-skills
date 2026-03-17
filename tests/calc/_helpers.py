"""Shared helpers for Calc tests."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from uno_bridge import uno_context


@contextmanager
def open_calc_doc(doc_path: Path | str) -> Iterator[Any]:
    """Open a Calc document through UNO for test assertions."""
    path = Path(doc_path)
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            yield doc
        finally:
            doc.close(True)


def get_cell_number_format(
    doc_path: Path | str,
    sheet: str | int,
    row: int,
    col: int,
) -> int:
    """Return the number-format key for one Calc cell."""
    with open_calc_doc(doc_path) as doc:
        cell = _get_sheet(doc, sheet).getCellByPosition(col, row)
        return int(cell.NumberFormat)


def get_named_range_formula(doc_path: Path | str, name: str) -> str:
    """Return the formula backing one named range."""
    with open_calc_doc(doc_path) as doc:
        named_range = doc.NamedRanges.getByName(name)
        return str(named_range.Content)


def named_range_exists(doc_path: Path | str, name: str) -> bool:
    """Return whether a named range exists."""
    with open_calc_doc(doc_path) as doc:
        return bool(doc.NamedRanges.hasByName(name))


def get_validation_properties(
    doc_path: Path | str,
    sheet: str | int,
    row: int,
    col: int,
    end_row: int | None = None,
    end_col: int | None = None,
) -> dict[str, Any]:
    """Return validation properties for a cell or rectangular range."""
    if end_row is None:
        end_row = row
    if end_col is None:
        end_col = col

    with open_calc_doc(doc_path) as doc:
        cell_range = _get_sheet(doc, sheet).getCellRangeByPosition(
            col,
            row,
            end_col,
            end_row,
        )
        validation = cell_range.Validation
        return {
            "type": str(getattr(validation.Type, "value", validation.Type)),
            "operator": str(getattr(validation.Operator, "value", validation.Operator)),
            "formula1": str(validation.Formula1),
            "formula2": str(validation.Formula2),
            "show_error": bool(validation.ShowErrorMessage),
            "error_message": str(validation.ErrorMessage),
            "show_input": bool(validation.ShowInputMessage),
            "input_title": str(validation.InputTitle),
            "input_message": str(validation.InputMessage),
            "ignore_blank": bool(validation.IgnoreBlankCells),
            "error_style": str(
                getattr(validation.ErrorAlertStyle, "value", validation.ErrorAlertStyle)
            ),
        }


def list_chart_names(doc_path: Path | str, sheet: str | int) -> list[str]:
    """Return chart names for one sheet in document order."""
    with open_calc_doc(doc_path) as doc:
        charts = _get_sheet(doc, sheet).Charts
        return [charts.getByIndex(index).Name for index in range(charts.getCount())]


def get_chart_details(
    doc_path: Path | str,
    sheet: str | int,
    *,
    name: str | None = None,
    index: int | None = None,
) -> dict[str, Any]:
    """Return persisted details for one chart."""
    with open_calc_doc(doc_path) as doc:
        sheet_obj = _get_sheet(doc, sheet)
        charts = sheet_obj.Charts
        if name is not None:
            table_chart = charts.getByName(name)
        else:
            assert index is not None
            table_chart = charts.getByIndex(index)

        embedded = table_chart.EmbeddedObject
        shape = _get_chart_shape(sheet_obj, str(table_chart.Name))
        ranges = table_chart.getRanges()
        range_details = None
        if ranges:
            first_range = ranges[0]
            range_details = {
                "start_row": int(first_range.StartRow),
                "start_col": int(first_range.StartColumn),
                "end_row": int(first_range.EndRow),
                "end_col": int(first_range.EndColumn),
            }

        title = None
        if getattr(embedded, "HasMainTitle", False):
            title = str(embedded.Title.String)

        return {
            "name": str(table_chart.Name),
            "title": title,
            "range": range_details,
            "x": int(shape.BoundRect.X),
            "y": int(shape.BoundRect.Y),
            "width": int(shape.BoundRect.Width),
            "height": int(shape.BoundRect.Height),
        }


def _get_sheet(doc: Any, sheet: str | int) -> Any:
    if isinstance(sheet, int):
        return doc.Sheets.getByIndex(sheet)
    return doc.Sheets.getByName(sheet)


def _get_chart_shape(sheet: Any, chart_name: str) -> Any:
    draw_page = sheet.DrawPage
    for index in range(draw_page.getCount()):
        shape = draw_page.getByIndex(index)
        if str(getattr(shape, "PersistName", "")) == chart_name:
            return shape
    raise AssertionError(f'Chart shape not found for "{chart_name}"')
