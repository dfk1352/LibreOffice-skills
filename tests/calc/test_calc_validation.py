"""Tests for Calc data validation."""

import pytest


def test_add_validation_rule(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.validation import add_validation
    from uno_bridge import uno_context

    path = tmp_path / "validate.ods"
    create_spreadsheet(str(path))

    add_validation(
        str(path),
        0,
        0,
        0,
        1,
        0,
        {
            "type": "whole",
            "condition": "between",
            "value1": 1,
            "value2": 10,
            "show_error": True,
            "error_message": "Value must be 1-10",
        },
    )

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            sheet = doc.Sheets.getByIndex(0)
            cell_range = sheet.getCellRangeByPosition(0, 0, 0, 1)
            validation = cell_range.Validation
            assert validation.Type != 0
            assert validation.Operator != 0
            assert validation.Formula1 == "1"
            assert validation.Formula2 == "10"
            assert validation.ShowErrorMessage
        finally:
            doc.close(True)


def test_add_validation_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.validation import add_validation

    with pytest.raises(DocumentNotFoundError):
        add_validation(
            str(tmp_path / "missing.ods"),
            "Sheet1",
            0,
            0,
            0,
            0,
            {"type": "whole", "condition": "equal", "value1": 1},
        )
