"""Tests for Calc find & replace operations."""


def test_find_replace_in_calc(tmp_path):
    from libreoffice_skills.calc.cells import get_cell, set_cell
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.find_replace import find_replace

    path = tmp_path / "find_replace.ods"
    create_spreadsheet(str(path))
    set_cell(str(path), 0, 0, 0, "Hello World", type="text")

    count = find_replace(str(path), "Hello", "Goodbye")

    assert count >= 1

    cell = get_cell(str(path), 0, 0, 0)
    assert "Goodbye" in cell["value"]


def test_find_replace_in_calc_scoped_to_sheet(tmp_path):
    from libreoffice_skills.calc.cells import get_cell, set_cell
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.find_replace import find_replace
    from libreoffice_skills.calc.sheets import add_sheet

    path = tmp_path / "scoped.ods"
    create_spreadsheet(str(path))
    add_sheet(str(path), "Sheet2")
    set_cell(str(path), 0, 0, 0, "Replace me", type="text")
    set_cell(str(path), "Sheet2", 0, 0, "Replace me", type="text")

    # Only replace on Sheet2
    count = find_replace(str(path), "Replace", "Changed", sheet="Sheet2")

    assert count >= 1

    # Sheet1 should be unchanged
    cell_sheet1 = get_cell(str(path), 0, 0, 0)
    assert "Replace" in cell_sheet1["value"]

    # Sheet2 should have the replacement
    cell_sheet2 = get_cell(str(path), "Sheet2", 0, 0)
    assert "Changed" in cell_sheet2["value"]
