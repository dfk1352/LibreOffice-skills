# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import json
from pathlib import Path

from tests.calc._helpers import (
    get_cell_number_format,
    get_chart_details,
    get_validation_properties,
    list_chart_names,
    named_range_exists,
)


def run_calc_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build Calc workflow artifacts that exercise every public Calc tool."""
    from calc import (
        CalcSession,
        CalcTarget,
        CellFormatting,
        ChartSpec,
        ValidationRule,
        patch,
        snapshot_area,
    )
    from calc.core import create_spreadsheet

    output_dir.mkdir(parents=True, exist_ok=True)

    session_doc = output_dir / "session_workflow.ods"
    create_spreadsheet(str(session_doc))

    session = CalcSession(str(session_doc))
    session.write_cell(
        CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
        "Workflow Seed",
        value_type="text",
    )
    assert (
        session.read_cell(CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0))[
            "value"
        ]
        == "Workflow Seed"
    )

    session.rename_sheet(CalcTarget(kind="sheet", sheet="Sheet1"), "Revenue Data")
    session.add_sheet("Summary")
    session.add_sheet("Scratch")
    assert [sheet["name"] for sheet in session.list_sheets()] == [
        "Revenue Data",
        "Summary",
        "Scratch",
    ]

    summary_range = CalcTarget(
        kind="range",
        sheet="Summary",
        row=0,
        col=0,
        end_row=2,
        end_col=1,
    )
    session.write_range(
        summary_range,
        [["Metric", "Value"], ["Status", "Draft"], ["Scope", "Session workflow"]],
    )

    data_range = CalcTarget(
        kind="range",
        sheet="Revenue Data",
        row=0,
        col=0,
        end_row=2,
        end_col=1,
    )
    session.write_range(
        data_range,
        [["Label", "Value"], ["Revenue", 100], ["Cost", 80]],
    )
    assert session.read_range(data_range)[1][1]["value"] == 100

    currency_target = CalcTarget(
        kind="range",
        sheet="Revenue Data",
        row=1,
        col=1,
        end_row=2,
        end_col=1,
    )
    session.format_range(
        currency_target, CellFormatting(number_format="currency", bold=True)
    )

    retained_range = CalcTarget(
        kind="range",
        sheet="Revenue Data",
        row=1,
        col=1,
        end_row=2,
        end_col=1,
    )
    session.define_named_range("RevenueValues", retained_range)
    assert (
        session.get_named_range(CalcTarget(kind="named_range", name="RevenueValues"))[
            "name"
        ]
        == "RevenueValues"
    )

    temporary_range = CalcTarget(
        kind="range",
        sheet="Revenue Data",
        row=1,
        col=0,
        end_row=2,
        end_col=0,
    )
    session.define_named_range("TemporaryLabels", temporary_range)
    session.delete_named_range(CalcTarget(kind="named_range", name="TemporaryLabels"))

    session.set_validation(
        currency_target,
        ValidationRule(
            type="whole",
            condition="between",
            value1=1,
            value2=1000,
            show_error=True,
            error_message="Revenue must stay within 1-1000",
        ),
    )
    session.set_validation(
        temporary_range,
        ValidationRule(type="text_length", condition="greater_than", value1=3),
    )
    session.clear_validation(temporary_range)

    total_target = CalcTarget(kind="cell", sheet="Revenue Data", row=4, col=1)
    session.write_cell(total_target, "=SUM(RevenueValues)", value_type="formula")
    session.recalculate()
    assert session.read_cell(total_target)["value"] == 180

    session.create_chart(
        CalcTarget(kind="sheet", sheet="Revenue Data"),
        ChartSpec(
            chart_type="line",
            data_range=data_range,
            anchor_row=5,
            anchor_col=0,
            width=5000,
            height=3000,
            title="Revenue Trend",
        ),
    )
    session.create_chart(
        CalcTarget(kind="sheet", sheet="Summary"),
        ChartSpec(
            chart_type="line",
            data_range=summary_range,
            anchor_row=4,
            anchor_col=0,
            width=3500,
            height=2000,
            title="Temporary Summary Chart",
        ),
    )
    session.close()

    snapshot_before = output_dir / "calc_snapshot_before.png"
    snapshot_area(
        str(session_doc),
        str(snapshot_before),
        sheet="Revenue Data",
        row=0,
        col=0,
        width=900,
        height=500,
    )

    session = CalcSession(str(session_doc))
    session.write_range(
        CalcTarget(
            kind="range",
            sheet="Revenue Data",
            row=0,
            col=0,
            end_row=3,
            end_col=1,
        ),
        [["Label", "Value"], ["Revenue", 100], ["Cost", 80], ["Profit", 20]],
    )
    session.update_chart(
        CalcTarget(kind="chart", sheet="Revenue Data", index=0),
        ChartSpec(
            chart_type="bar",
            data_range=CalcTarget(
                kind="range",
                sheet="Revenue Data",
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
    session.write_cell(
        CalcTarget(kind="cell", sheet="Summary", row=1, col=1),
        "Ready for review",
        value_type="text",
    )
    session.write_cell(
        CalcTarget(kind="cell", sheet="Summary", row=3, col=0),
        "Updated chart retained on Revenue Data",
        value_type="text",
    )
    session.close()

    with CalcSession(str(session_doc)) as session:
        session.format_range(
            CalcTarget(
                kind="range", sheet="Revenue Data", row=0, col=0, end_row=0, end_col=1
            ),
            CellFormatting(bold=True, font_size=14),
        )
        session.format_range(
            CalcTarget(
                kind="range", sheet="Revenue Data", row=1, col=0, end_row=3, end_col=0
            ),
            CellFormatting(italic=True, font_size=13),
        )
        patch_result = session.patch(
            "[operation]\n"
            "type = write_cell\n"
            "target.kind = cell\n"
            "target.sheet = Revenue Data\n"
            "target.row = 0\n"
            "target.col = 0\n"
            "value = Category\n"
            "value_type = text\n"
            "[operation]\n"
            "type = write_cell\n"
            "target.kind = cell\n"
            "target.sheet = Revenue Data\n"
            "target.row = 5\n"
            "target.col = 0\n"
            "value = Patched note\n"
            "value_type = text\n"
            "[operation]\n"
            "type = format_range\n"
            "target.kind = range\n"
            "target.sheet = Revenue Data\n"
            "target.row = 0\n"
            "target.col = 0\n"
            "target.end_row = 0\n"
            "target.end_col = 1\n"
            "format.bold = true\n"
            "format.font_size = 14\n",
            mode="atomic",
        )
        assert patch_result.overall_status == "ok"
        session.delete_chart(CalcTarget(kind="chart", sheet="Summary", index=0))
        session.delete_sheet(CalcTarget(kind="sheet", sheet="Scratch"))
        workflow_pdf = output_dir / "workflow.pdf"
        session.export(str(workflow_pdf), "pdf")

    snapshot_after = output_dir / "calc_snapshot_after.png"
    snapshot_area(
        str(session_doc),
        str(snapshot_after),
        sheet="Revenue Data",
        row=0,
        col=0,
        width=900,
        height=500,
    )

    atomic_doc = output_dir / "patch_atomic.ods"
    create_spreadsheet(str(atomic_doc))
    with CalcSession(str(atomic_doc)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Atomic baseline",
            value_type="text",
        )

    patch(
        str(atomic_doc),
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 0\n"
        "target.col = 0\n"
        "value = rolled back\n"
        "value_type = text\n"
        "[operation]\n"
        "type = delete_chart\n"
        "target.kind = chart\n"
        "target.sheet = Sheet1\n"
        "target.name = MissingChart\n"
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 1\n"
        "target.col = 0\n"
        "value = skipped later\n"
        "value_type = text\n",
        mode="atomic",
    )

    best_effort_doc = output_dir / "patch_best_effort.ods"
    create_spreadsheet(str(best_effort_doc))
    with CalcSession(str(best_effort_doc)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Best effort baseline",
            value_type="text",
        )

    patch(
        str(best_effort_doc),
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 1\n"
        "target.col = 0\n"
        "value = preserved first\n"
        "value_type = text\n"
        "[operation]\n"
        "type = delete_named_range\n"
        "target.kind = named_range\n"
        "target.name = MissingRange\n"
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 2\n"
        "target.col = 0\n"
        "value = preserved later\n"
        "value_type = text\n",
        mode="best_effort",
    )

    json_source = output_dir / "import_data.json"
    json_source.write_text(
        json.dumps(
            [
                {"Product": "Widget", "Units": 150, "Revenue": 4500},
                {"Product": "Gadget", "Units": 80, "Revenue": 3200},
            ]
        )
    )
    json_imported = output_dir / "json_imported.ods"
    create_spreadsheet(str(json_imported), source=str(json_source))

    return {
        "session_workflow": session_doc,
        "patch_atomic": atomic_doc,
        "patch_best_effort": best_effort_doc,
        "workflow_pdf": workflow_pdf,
        "snapshot_before": snapshot_before,
        "snapshot_after": snapshot_after,
        "json_imported": json_imported,
    }


def test_session_workflow_document_state(tmp_path):
    """Session workflow leaves visible session-first Calc state behind."""
    from calc import CalcTarget, CalcSession

    outputs = run_calc_end_to_end_workflow(tmp_path)

    with CalcSession(str(outputs["session_workflow"])) as session:
        sheets = session.list_sheets()
        total = session.read_cell(
            CalcTarget(kind="cell", sheet="Revenue Data", row=4, col=1)
        )
        patched_note = session.read_cell(
            CalcTarget(kind="cell", sheet="Revenue Data", row=5, col=0)
        )
        patched_header = session.read_cell(
            CalcTarget(kind="cell", sheet="Revenue Data", row=0, col=0)
        )
        summary_status = session.read_cell(
            CalcTarget(kind="cell", sheet="Summary", row=1, col=1)
        )

    assert [sheet["name"] for sheet in sheets] == ["Revenue Data", "Summary"]
    assert total["value"] == 180
    assert patched_note["value"] == "Patched note"
    assert patched_header["value"] == "Category"
    assert summary_status["value"] == "Ready for review"
    assert named_range_exists(outputs["session_workflow"], "RevenueValues")
    assert not named_range_exists(outputs["session_workflow"], "TemporaryLabels")
    assert get_cell_number_format(
        outputs["session_workflow"], "Revenue Data", 1, 1
    ) == get_cell_number_format(
        outputs["session_workflow"],
        "Revenue Data",
        2,
        1,
    )
    validation = get_validation_properties(
        outputs["session_workflow"],
        "Revenue Data",
        1,
        1,
        2,
        1,
    )
    assert validation["type"] == "WHOLE"
    cleared_validation = get_validation_properties(
        outputs["session_workflow"],
        "Revenue Data",
        1,
        0,
        2,
        0,
    )
    assert cleared_validation["type"] == "ANY"
    assert list_chart_names(outputs["session_workflow"], "Revenue Data") == [
        "Revenue Trend"
    ]
    assert list_chart_names(outputs["session_workflow"], "Summary") == []
    chart = get_chart_details(outputs["session_workflow"], "Revenue Data", index=0)
    assert chart["title"] == "Revenue Trend Updated"
    assert chart["range"] == {
        "start_row": 0,
        "start_col": 0,
        "end_row": 3,
        "end_col": 1,
    }
    assert chart["width"] >= 6000
    assert chart["height"] >= 3500


def test_patch_workflow_documents_capture_atomic_and_best_effort_results(tmp_path):
    """Standalone patch workflow preserves atomic and best-effort semantics."""
    from calc import CalcTarget, CalcSession

    outputs = run_calc_end_to_end_workflow(tmp_path)

    with CalcSession(str(outputs["patch_atomic"])) as session:
        atomic_baseline = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0)
        )
        atomic_skipped = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=1, col=0)
        )

    with CalcSession(str(outputs["patch_best_effort"])) as session:
        best_effort_first = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=1, col=0)
        )
        best_effort_later = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=2, col=0)
        )

    assert atomic_baseline["value"] == "Atomic baseline"
    assert atomic_skipped["value"] == 0.0
    assert best_effort_first["value"] == "preserved first"
    assert best_effort_later["value"] == "preserved later"


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable Calc workflow files in test-output/calc/."""
    output_dir = Path("test-output/calc")

    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_calc_end_to_end_workflow(output_dir)

    for key in (
        "session_workflow",
        "patch_atomic",
        "patch_best_effort",
        "workflow_pdf",
        "snapshot_before",
        "snapshot_after",
        "json_imported",
    ):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    with open(outputs["workflow_pdf"], "rb") as handle:
        assert handle.read(5) == b"%PDF-"

    before_bytes = outputs["snapshot_before"].read_bytes()
    after_bytes = outputs["snapshot_after"].read_bytes()
    assert before_bytes != after_bytes

    for key in ("snapshot_before", "snapshot_after"):
        with open(outputs[key], "rb") as handle:
            assert handle.read(8) == b"\x89PNG\r\n\x1a\n"


def test_json_import_workflow_data(tmp_path):
    """JSON import produces a spreadsheet with the expected tabular data."""
    from tests.calc._helpers import open_calc_doc

    outputs = run_calc_end_to_end_workflow(tmp_path)

    with open_calc_doc(outputs["json_imported"]) as doc:
        sheet = doc.getSheets().getByIndex(0)
        assert sheet.getCellByPosition(0, 0).getString() == "Product"
        assert sheet.getCellByPosition(1, 0).getString() == "Units"
        assert sheet.getCellByPosition(2, 0).getString() == "Revenue"
        assert sheet.getCellByPosition(0, 1).getString() == "Widget"
        assert sheet.getCellByPosition(1, 1).getString() == "150"
        assert sheet.getCellByPosition(0, 2).getString() == "Gadget"
