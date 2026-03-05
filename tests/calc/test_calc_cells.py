"""Tests for Calc cell operations."""


def test_set_and_get_cell_value(tmp_path) -> None:
    from calc.cells import get_cell, set_cell
    from calc.core import create_spreadsheet

    path = tmp_path / "cells.ods"
    create_spreadsheet(str(path))
    set_cell(str(path), 0, 0, 0, 42, type="number")

    cell = get_cell(str(path), 0, 0, 0)
    assert cell["value"] == 42
    assert cell["type"] == "number"
    assert cell["error"] is None
