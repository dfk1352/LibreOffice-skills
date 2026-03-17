"""Test basic Calc package imports."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false


def test_imports_calc_package() -> None:
    import calc

    assert hasattr(calc, "create_spreadsheet")
    assert callable(calc.create_spreadsheet)
    assert hasattr(calc, "export_spreadsheet")
    assert callable(calc.export_spreadsheet)
    assert hasattr(calc, "snapshot_area")
    assert callable(calc.snapshot_area)
    assert hasattr(calc, "open_calc_session")
    assert callable(calc.open_calc_session)
    assert hasattr(calc, "CalcSession")
    assert hasattr(calc, "CalcTarget")
    assert hasattr(calc, "CellFormatting")
    assert hasattr(calc, "ValidationRule")
    assert hasattr(calc, "ChartSpec")
    assert hasattr(calc, "patch")
    assert callable(calc.patch)
    assert hasattr(calc, "PatchApplyResult")
    assert hasattr(calc, "PatchOperationResult")
