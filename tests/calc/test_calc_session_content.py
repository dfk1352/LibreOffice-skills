"""Tests for Calc content CRUD through sessions."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.calc._helpers import (
    get_cell_number_format,
    get_chart_details,
    get_validation_properties,
    list_chart_names,
)


def test_session_write_and_read_cell_preserves_text_number_and_formula_behaviour(
    tmp_path,
):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "cells.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Label",
            value_type="text",
        )
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=1),
            21,
            value_type="number",
        )
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=2),
            "=B1*2",
            value_type="formula",
        )
        session.recalculate()

        text_cell = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0)
        )
        number_cell = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=1)
        )
        formula_cell = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=2)
        )

    assert text_cell["value"] == "Label"
    assert text_cell["type"] == "text"
    assert number_cell["value"] == 21
    assert number_cell["type"] == "number"
    assert formula_cell["formula"] == "=B1*2"
    assert formula_cell["value"] == 42
    assert formula_cell["type"] == "formula"


def test_session_write_and_read_range_preserves_rectangular_shape_and_values(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "ranges.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_range(
            CalcTarget(
                kind="range", sheet="Sheet1", row=0, col=0, end_row=1, end_col=1
            ),
            [["Label", 10], ["Cost", 5]],
        )
        results = session.read_range(
            CalcTarget(kind="range", sheet="Sheet1", row=0, col=0, end_row=1, end_col=1)
        )

    assert len(results) == 2
    assert len(results[0]) == 2
    assert results[0][0]["value"] == "Label"
    assert results[0][1]["value"] == 10
    assert results[1][0]["value"] == "Cost"
    assert results[1][1]["value"] == 5


def test_session_format_range_applies_number_format_to_multi_cell_range(tmp_path):
    from calc import CalcTarget, CellFormatting, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "format_range.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_range(
            CalcTarget(
                kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1
            ),
            [[100], [80]],
        )
        session.format_range(
            CalcTarget(
                kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1
            ),
            CellFormatting(number_format="currency", bold=True),
        )

    assert get_cell_number_format(doc_path, "Sheet1", 1, 1) == get_cell_number_format(
        doc_path,
        "Sheet1",
        2,
        1,
    )


def test_session_sheet_operations_preserve_zero_based_order(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "sheets.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.add_sheet("Summary")
        session.add_sheet("Data", index=1)
        session.rename_sheet(CalcTarget(kind="sheet", sheet_index=1), "Inputs")
        session.delete_sheet(CalcTarget(kind="sheet", sheet="Summary"))
        sheets = session.list_sheets()

    assert sheets == [
        {"name": "Sheet1", "index": 0, "visible": True},
        {"name": "Inputs", "index": 1, "visible": True},
    ]


def test_session_named_range_lifecycle_is_predictable(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet
    from calc.exceptions import NamedRangeNotFoundError

    doc_path = tmp_path / "named_ranges.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_range(
            CalcTarget(
                kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1
            ),
            [[100], [80]],
        )
        session.define_named_range(
            "RevenueValues",
            CalcTarget(
                kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1
            ),
        )
        named_range = session.get_named_range(
            CalcTarget(kind="named_range", name="RevenueValues")
        )
        session.delete_named_range(CalcTarget(kind="named_range", name="RevenueValues"))
        with pytest.raises(NamedRangeNotFoundError):
            session.get_named_range(
                CalcTarget(kind="named_range", name="RevenueValues")
            )

    assert named_range["name"] == "RevenueValues"
    assert "$B$2" in named_range["formula"]
    assert "$B$3" in named_range["formula"]


def test_session_set_validation_attaches_rule_to_target_range(tmp_path):
    from calc import CalcTarget, ValidationRule, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "validation_set.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.set_validation(
            CalcTarget(
                kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1
            ),
            ValidationRule(
                type="whole",
                condition="between",
                value1=1,
                value2=1000,
                show_error=True,
                error_message="Value must stay within 1-1000",
            ),
        )

    validation = get_validation_properties(doc_path, "Sheet1", 1, 1, 2, 1)
    assert validation["type"] == "WHOLE"
    assert validation["operator"] == "BETWEEN"
    assert validation["formula1"] == "1"
    assert validation["formula2"] == "1000"
    assert validation["show_error"] is True
    assert validation["error_message"] == "Value must stay within 1-1000"


def test_session_clear_validation_removes_existing_rule(tmp_path):
    from calc import CalcTarget, ValidationRule, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "validation_clear.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        range_target = CalcTarget(
            kind="range",
            sheet="Sheet1",
            row=1,
            col=1,
            end_row=2,
            end_col=1,
        )
        session.set_validation(
            range_target,
            ValidationRule(type="whole", condition="between", value1=1, value2=10),
        )

    with open_calc_session(str(doc_path)) as session:
        session.clear_validation(
            CalcTarget(kind="range", sheet="Sheet1", row=1, col=1, end_row=2, end_col=1)
        )

    validation = get_validation_properties(doc_path, "Sheet1", 1, 1, 2, 1)
    assert validation["type"] == "ANY"
    assert validation["formula1"] == ""
    assert validation["formula2"] == ""


def test_session_chart_lifecycle_visibly_mutates_chart_state(tmp_path):
    from calc import CalcTarget, ChartSpec, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "charts.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.add_sheet("Data")
        session.write_range(
            CalcTarget(kind="range", sheet="Data", row=0, col=0, end_row=2, end_col=1),
            [["Label", "Value"], ["Revenue", 100], ["Cost", 80]],
        )
        session.create_chart(
            CalcTarget(kind="sheet", sheet="Data"),
            ChartSpec(
                chart_type="line",
                data_range=CalcTarget(
                    kind="range",
                    sheet="Data",
                    row=0,
                    col=0,
                    end_row=2,
                    end_col=1,
                ),
                anchor_row=5,
                anchor_col=0,
                width=5000,
                height=3000,
                title="Revenue Trend",
            ),
        )

    created_chart = get_chart_details(doc_path, "Data", index=0)
    assert created_chart["title"] == "Revenue Trend"
    assert created_chart["range"] == {
        "start_row": 0,
        "start_col": 0,
        "end_row": 2,
        "end_col": 1,
    }
    assert created_chart["width"] > 0
    assert created_chart["height"] > 0

    with open_calc_session(str(doc_path)) as session:
        session.write_range(
            CalcTarget(kind="range", sheet="Data", row=0, col=0, end_row=3, end_col=1),
            [["Label", "Value"], ["Revenue", 100], ["Cost", 80], ["Profit", 20]],
        )
        session.update_chart(
            CalcTarget(kind="chart", sheet="Data", index=0),
            ChartSpec(
                chart_type="bar",
                data_range=CalcTarget(
                    kind="range",
                    sheet="Data",
                    row=0,
                    col=0,
                    end_row=3,
                    end_col=1,
                ),
                anchor_row=7,
                anchor_col=2,
                width=6000,
                height=3500,
                title="Revenue Trend Updated",
            ),
        )

    updated_chart = get_chart_details(doc_path, "Data", index=0)
    assert updated_chart["title"] == "Revenue Trend Updated"
    assert updated_chart["range"] == {
        "start_row": 0,
        "start_col": 0,
        "end_row": 3,
        "end_col": 1,
    }
    assert updated_chart["title"] != created_chart["title"]
    assert updated_chart["width"] > created_chart["width"]
    assert updated_chart["height"] > created_chart["height"]
    assert updated_chart["x"] != created_chart["x"]
    assert updated_chart["y"] != created_chart["y"]

    with open_calc_session(str(doc_path)) as session:
        session.delete_chart(CalcTarget(kind="chart", sheet="Data", index=0))

    assert list_chart_names(doc_path, "Data") == []


def test_session_recalculate_updates_formula_results_after_earlier_writes(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "recalculate.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            10,
            value_type="number",
        )
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=1),
            "=A1*2",
            value_type="formula",
        )
        session.recalculate()
        first_result = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=1)
        )

        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            12,
            value_type="number",
        )
        session.recalculate()
        second_result = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=1)
        )

    assert first_result["value"] == 20
    assert second_result["value"] == 24


def test_session_export_writes_alternate_format_from_current_session_state(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "export_session.ods"
    output_path = tmp_path / "export_session.pdf"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Export me",
            value_type="text",
        )
        session.export(str(output_path), "pdf")

    assert output_path.exists()
    assert output_path.stat().st_size > 0
    with open(output_path, "rb") as handle:
        assert handle.read(5) == b"%PDF-"


def test_session_patch_applies_operations_against_open_document(tmp_path):
    from calc import CalcTarget, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "session_patch.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        result = session.patch(
            "[operation]\n"
            "type = add_sheet\n"
            "name = Data\n"
            "[operation]\n"
            "type = write_cell\n"
            "target.kind = cell\n"
            "target.sheet = Data\n"
            "target.row = 0\n"
            "target.col = 0\n"
            "value = Patched\n"
            "value_type = text\n",
            mode="atomic",
        )
        patched_cell = session.read_cell(
            CalcTarget(kind="cell", sheet="Data", row=0, col=0)
        )

    assert result.overall_status == "ok"
    assert [operation.status for operation in result.operations] == ["ok", "ok"]
    assert patched_cell["value"] == "Patched"
