"""Tests for Calc formula recalculation."""


def test_formula_recalculation(tmp_path) -> None:
    from libreoffice_skills.calc.cells import get_cell, set_cell
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.recalc import recalculate

    path = tmp_path / "formula.ods"
    create_spreadsheet(str(path))
    set_cell(str(path), 0, 0, 0, 10, type="number")
    set_cell(str(path), 0, 0, 1, "=A1*2", type="formula")

    recalculate(str(path))

    cell = get_cell(str(path), 0, 0, 1)
    assert cell["value"] == 20
    assert cell["error"] is None
