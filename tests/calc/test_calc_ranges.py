"""Tests for Calc range operations."""


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
