"""Tests for Calc sheet helpers."""


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
