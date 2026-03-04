"""Test basic Calc package imports."""


def test_imports_calc_package() -> None:
    import libreoffice_skills.calc

    assert hasattr(libreoffice_skills.calc, "create_spreadsheet")
    assert callable(libreoffice_skills.calc.create_spreadsheet)
    assert hasattr(libreoffice_skills.calc, "get_cell")
    assert callable(libreoffice_skills.calc.get_cell)
    assert hasattr(libreoffice_skills.calc, "set_cell")
    assert callable(libreoffice_skills.calc.set_cell)
