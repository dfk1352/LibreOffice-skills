"""Tests for Calc range operations."""

import pytest


def test_set_and_get_range(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.ranges import get_range, set_range

    path = tmp_path / "ranges.ods"
    create_spreadsheet(str(path))

    data = [[1, 2], [3, 4]]
    set_range(str(path), 0, 0, 0, data)

    results = get_range(str(path), 0, 0, 0, 1, 1)
    assert results[0][0]["value"] == 1
    assert results[1][1]["value"] == 4


def test_set_range_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.ranges import set_range

    with pytest.raises(DocumentNotFoundError):
        set_range(str(tmp_path / "missing.ods"), "Sheet1", 0, 0, [[1]])


def test_get_range_raises_on_missing_file(tmp_path) -> None:
    from calc.exceptions import DocumentNotFoundError
    from calc.ranges import get_range

    with pytest.raises(DocumentNotFoundError):
        get_range(str(tmp_path / "missing.ods"), "Sheet1", 0, 0, 0, 0)
