"""Tests for Calc named ranges."""

import pytest


def test_define_named_range(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.named_ranges import (
        define_named_range,
        get_named_range,
    )

    path = tmp_path / "named.ods"
    create_spreadsheet(str(path))

    define_named_range(str(path), "Total", 0, 0, 0, 0, 1)
    named = get_named_range(str(path), "Total")
    assert named["name"] == "Total"
    # The formula should reference cells A1:B1 on the first sheet
    formula = named["formula"]
    assert "A1" in formula or "$A$1" in formula
    assert "B1" in formula or "$B$1" in formula


def test_define_named_range_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.named_ranges import define_named_range

    with pytest.raises(DocumentNotFoundError):
        define_named_range(str(tmp_path / "missing.ods"), "Total", "Sheet1", 0, 0)


def test_get_named_range_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.named_ranges import get_named_range

    with pytest.raises(DocumentNotFoundError):
        get_named_range(str(tmp_path / "missing.ods"), "Total")
