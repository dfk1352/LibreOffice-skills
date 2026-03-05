"""Test basic Calc package imports."""


def test_imports_calc_package() -> None:
    import calc

    assert hasattr(calc, "create_spreadsheet")
    assert callable(calc.create_spreadsheet)
    assert hasattr(calc, "get_cell")
    assert callable(calc.get_cell)
    assert hasattr(calc, "set_cell")
    assert callable(calc.set_cell)
