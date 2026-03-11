"""Tests for Calc cell operations."""

import pytest


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


def test_set_cell_raises_on_missing_file(tmp_path) -> None:
    from calc.cells import set_cell
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        set_cell(str(tmp_path / "missing.ods"), "Sheet1", 0, 0, "x")


def test_get_cell_raises_on_missing_file(tmp_path) -> None:
    from calc.cells import get_cell
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        get_cell(str(tmp_path / "missing.ods"), "Sheet1", 0, 0)
