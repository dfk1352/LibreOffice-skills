"""Tests for Calc sheet helpers."""

import pytest


def test_add_and_list_sheets(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.sheets import add_sheet, list_sheets

    path = tmp_path / "sheets.ods"
    create_spreadsheet(str(path))
    add_sheet(str(path), "Data")

    sheets = list_sheets(str(path))
    names = [sheet["name"] for sheet in sheets]
    assert "Data" in names


def test_add_rename_remove_sheet(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.sheets import (
        add_sheet,
        list_sheets,
        remove_sheet,
        rename_sheet,
    )

    path = tmp_path / "sheets_manage.ods"
    create_spreadsheet(str(path))
    add_sheet(str(path), "Draft")
    rename_sheet(str(path), "Draft", "Final")

    sheets = list_sheets(str(path))
    names = [sheet["name"] for sheet in sheets]
    assert "Final" in names
    assert "Draft" not in names

    remove_sheet(str(path), "Final")
    names = [sheet["name"] for sheet in list_sheets(str(path))]
    assert "Final" not in names


def test_add_sheet_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.sheets import add_sheet

    with pytest.raises(DocumentNotFoundError):
        add_sheet(str(tmp_path / "missing.ods"), "Data")


def test_list_sheets_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.sheets import list_sheets

    with pytest.raises(DocumentNotFoundError):
        list_sheets(str(tmp_path / "missing.ods"))


def test_rename_sheet_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.sheets import rename_sheet

    with pytest.raises(DocumentNotFoundError):
        rename_sheet(str(tmp_path / "missing.ods"), "Old", "New")


def test_remove_sheet_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.sheets import remove_sheet

    with pytest.raises(DocumentNotFoundError):
        remove_sheet(str(tmp_path / "missing.ods"), "Sheet1")
