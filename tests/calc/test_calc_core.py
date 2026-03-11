"""Tests for Calc core operations."""

import pytest


def test_create_spreadsheet_creates_file(tmp_path) -> None:
    from calc.core import create_spreadsheet

    output_path = tmp_path / "sample.ods"
    create_spreadsheet(str(output_path))
    assert output_path.exists()


def test_calc_package_exports_create_spreadsheet() -> None:
    from calc import create_spreadsheet

    assert callable(create_spreadsheet)
    # Verify it's the actual function, not an arbitrary callable
    assert create_spreadsheet.__module__ == "calc.core"


def test_export_spreadsheet_pdf(tmp_path) -> None:
    from calc.core import create_spreadsheet, export_spreadsheet

    path = tmp_path / "export.ods"
    output = tmp_path / "export.pdf"
    create_spreadsheet(str(path))
    export_spreadsheet(str(path), str(output), "pdf")
    assert output.exists()
    assert output.stat().st_size > 0
    # Verify it is a valid PDF by checking magic bytes
    with open(output, "rb") as f:
        assert f.read(5) == b"%PDF-"


def test_calc_exceptions_exports_document_not_found_error() -> None:
    from calc.exceptions import CalcSkillError, DocumentNotFoundError

    assert issubclass(DocumentNotFoundError, CalcSkillError)


def test_export_spreadsheet_raises_on_missing_file(tmp_path) -> None:
    from calc.core import export_spreadsheet
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        export_spreadsheet(
            str(tmp_path / "missing.ods"),
            str(tmp_path / "export.pdf"),
            "pdf",
        )
