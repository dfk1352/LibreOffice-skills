# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.calc._helpers import open_calc_doc


def _create_target_fixture(doc_path):
    from calc.core import create_spreadsheet

    create_spreadsheet(str(doc_path))

    with open_calc_doc(doc_path) as doc:
        import uno

        overview = doc.Sheets.getByIndex(0)
        overview.Name = "Overview"
        doc.Sheets.insertNewByName("Data", 1)
        data_sheet = doc.Sheets.getByName("Data")

        data_sheet.getCellByPosition(0, 0).String = "Label"
        data_sheet.getCellByPosition(1, 0).String = "Value"
        data_sheet.getCellByPosition(0, 1).String = "Revenue"
        data_sheet.getCellByPosition(1, 1).String = "marker"
        data_sheet.getCellByPosition(0, 2).String = "Cost"
        data_sheet.getCellByPosition(1, 2).Value = 80

        named_range_cells = data_sheet.getCellRangeByPosition(1, 1, 1, 2)
        base_address = data_sheet.getCellByPosition(1, 1).getCellAddress()
        doc.NamedRanges.addNewByName(
            "RevenueValues",
            named_range_cells.AbsoluteName,
            base_address,
            0,
        )

        anchor_cell = data_sheet.getCellByPosition(3, 0)
        rectangle = uno.createUnoStruct("com.sun.star.awt.Rectangle")
        rectangle.X = anchor_cell.Position.X
        rectangle.Y = anchor_cell.Position.Y
        rectangle.Width = 5000
        rectangle.Height = 3000

        charts = data_sheet.Charts
        chart_range = data_sheet.getCellRangeByPosition(0, 0, 1, 2).getRangeAddress()
        charts.addNewByName("RevenueChart", rectangle, (chart_range,), True, True)
        embedded = charts.getByName("RevenueChart").EmbeddedObject
        embedded.HasMainTitle = True
        embedded.Title.String = "Revenue Trend"
        doc.store()


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (
            {"kind": "cell", "sheet": "Data", "row": 1, "col": 1},
            {"kind": "cell", "sheet": "Data", "row": 1, "col": 1},
        ),
        (
            {
                "kind": "range",
                "sheet_index": 1,
                "row": 1,
                "col": 0,
                "end_row": 2,
                "end_col": 1,
            },
            {
                "kind": "range",
                "sheet_index": 1,
                "row": 1,
                "col": 0,
                "end_row": 2,
                "end_col": 1,
            },
        ),
        (
            {"kind": "sheet", "sheet": "Data"},
            {"kind": "sheet", "sheet": "Data"},
        ),
        (
            {"kind": "named_range", "name": "RevenueValues"},
            {"kind": "named_range", "name": "RevenueValues"},
        ),
        (
            {"kind": "chart", "sheet": "Data", "name": "RevenueChart"},
            {"kind": "chart", "sheet": "Data", "name": "RevenueChart"},
        ),
    ],
)
def test_parse_target_accepts_valid_calc_target_forms(raw, expected):
    from calc.targets import parse_target

    target = parse_target(raw)

    for key, value in expected.items():
        assert getattr(target, key) == value


def test_parse_target_rejects_sheet_and_sheet_index_together():
    from calc.exceptions import InvalidTargetError
    from calc.targets import parse_target

    with pytest.raises(InvalidTargetError):
        parse_target(
            {"kind": "cell", "sheet": "Data", "sheet_index": 1, "row": 0, "col": 0}
        )


def test_parse_target_rejects_negative_cell_coordinates():
    from calc.exceptions import InvalidTargetError
    from calc.targets import parse_target

    with pytest.raises(InvalidTargetError):
        parse_target({"kind": "cell", "sheet": "Data", "row": -1, "col": 0})


def test_resolve_sheet_target_supports_zero_based_sheet_index(tmp_path):
    from calc import CalcTarget
    from calc.targets import resolve_sheet_target

    doc_path = tmp_path / "resolve_sheet.ods"
    _create_target_fixture(doc_path)

    with open_calc_doc(doc_path) as doc:
        sheet = resolve_sheet_target(CalcTarget(kind="sheet", sheet_index=1), doc)
        sheet_name = sheet.Name

    assert sheet_name == "Data"


def test_resolve_cell_target_uses_sheet_and_exact_coordinates(tmp_path):
    from calc import CalcTarget
    from calc.targets import resolve_cell_target

    doc_path = tmp_path / "resolve_cell.ods"
    _create_target_fixture(doc_path)

    with open_calc_doc(doc_path) as doc:
        cell = resolve_cell_target(
            CalcTarget(kind="cell", sheet="Data", row=1, col=1),
            doc,
        )
        cell_value = cell.String

    assert cell_value == "marker"


def test_resolve_range_target_preserves_rectangular_bounds_exactly(tmp_path):
    from calc import CalcTarget
    from calc.targets import resolve_range_target

    doc_path = tmp_path / "resolve_range.ods"
    _create_target_fixture(doc_path)

    with open_calc_doc(doc_path) as doc:
        cell_range = resolve_range_target(
            CalcTarget(kind="range", sheet="Data", row=1, col=0, end_row=2, end_col=1),
            doc,
        )
        address = cell_range.getRangeAddress()

    assert int(address.StartRow) == 1
    assert int(address.StartColumn) == 0
    assert int(address.EndRow) == 2
    assert int(address.EndColumn) == 1


def test_resolve_named_range_target_resolves_name_and_missing_name_fails_clearly(
    tmp_path,
):
    from calc import CalcTarget
    from calc.exceptions import NamedRangeNotFoundError
    from calc.targets import resolve_named_range_target

    doc_path = tmp_path / "resolve_named_range.ods"
    _create_target_fixture(doc_path)

    with open_calc_doc(doc_path) as doc:
        named_range = resolve_named_range_target(
            CalcTarget(kind="named_range", name="RevenueValues"),
            doc,
        )
        assert "$B$2" in named_range.Content
        assert "$B$3" in named_range.Content

        with pytest.raises(NamedRangeNotFoundError):
            resolve_named_range_target(
                CalcTarget(kind="named_range", name="MissingRange"),
                doc,
            )


def test_resolve_chart_target_supports_name_and_zero_based_index(tmp_path):
    from calc import CalcTarget
    from calc.targets import resolve_chart_target

    doc_path = tmp_path / "resolve_chart.ods"
    _create_target_fixture(doc_path)

    with open_calc_doc(doc_path) as doc:
        named_chart = resolve_chart_target(
            CalcTarget(kind="chart", sheet="Data", name="RevenueChart"),
            doc,
        )
        indexed_chart = resolve_chart_target(
            CalcTarget(kind="chart", sheet_index=1, index=0),
            doc,
        )
        named_chart_name = named_chart.Name
        indexed_chart_name = indexed_chart.Name

    assert named_chart_name == "RevenueChart"
    assert indexed_chart_name == "RevenueChart"
