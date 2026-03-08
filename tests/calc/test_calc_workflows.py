"""Integration tests for Calc workflows."""

# pyright: reportMissingImports=false

from pathlib import Path


def run_calc_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build a Calc spreadsheet exercising every tool, then snapshot before/after.

    This is the single function that produces all inspectable output files.
    It exercises: create_spreadsheet, add/rename/remove/list_sheets,
    set/get_cell, set/get_range, define_named_range, recalculate,
    apply_format, add_validation, create_chart, export_spreadsheet,
    and snapshot_area.

    Args:
        output_dir: Directory where all output files are written.

    Returns:
        Dict mapping logical names to output file paths:
            "spreadsheet"     -> workflow.ods
            "pdf"             -> workflow.pdf
            "snapshot_before" -> calc_snapshot_before.png
            "snapshot_after"  -> calc_snapshot_after.png
    """
    from calc.charts import create_chart
    from calc.cells import get_cell, set_cell
    from calc.core import create_spreadsheet, export_spreadsheet
    from calc.formatting import apply_format
    from calc.named_ranges import define_named_range
    from calc.ranges import get_range, set_range
    from calc.recalc import recalculate
    from calc.sheets import (
        add_sheet,
        list_sheets,
        remove_sheet,
        rename_sheet,
    )
    from calc.snapshot import snapshot_area
    from calc.validation import add_validation

    output_dir.mkdir(parents=True, exist_ok=True)

    # --- Build the spreadsheet ---
    path = output_dir / "workflow.ods"
    create_spreadsheet(str(path))
    add_sheet(str(path), "Data")
    rename_sheet(str(path), "Data", "DataFinal")
    sheets = [sheet["name"] for sheet in list_sheets(str(path))]
    assert "DataFinal" in sheets

    set_range(
        str(path),
        "DataFinal",
        0,
        0,
        [["Label", "Value"], ["Revenue", 100], ["Revenue", 200]],
    )
    apply_format(str(path), "DataFinal", 1, 1, {"number_format": "currency"})
    apply_format(str(path), "DataFinal", 2, 1, {"number_format": "currency"})

    define_named_range(str(path), "RevenueValues", "DataFinal", 1, 1, 2, 1)
    set_cell(str(path), "DataFinal", 3, 1, "=SUM(RevenueValues)", type="formula")
    recalculate(str(path))
    total = get_cell(str(path), "DataFinal", 3, 1)
    assert total["value"] == 300

    add_validation(
        str(path),
        "DataFinal",
        1,
        1,
        2,
        1,
        {
            "type": "whole",
            "condition": "between",
            "value1": 1,
            "value2": 1000,
            "show_error": True,
            "error_message": "Value must be 1-1000",
        },
    )

    create_chart(
        str(path),
        "DataFinal",
        (0, 0, 2, 1),
        "line",
        anchor=(5, 0),
        size=(5000, 3000),
        title="Revenue Trend",
    )

    range_results = get_range(str(path), "DataFinal", 1, 1, 2, 1)
    assert range_results[0][0]["value"] == 100
    assert range_results[1][0]["value"] == 200

    add_sheet(str(path), "Scratch")
    remove_sheet(str(path), "Scratch")

    export_spreadsheet(str(path), str(output_dir / "workflow.pdf"), "pdf")

    # --- Snapshot BEFORE formatting changes ---
    before_path = output_dir / "calc_snapshot_before.png"
    snapshot_area(str(path), str(before_path), sheet="DataFinal")

    # --- Apply visible formatting changes ---
    apply_format(str(path), "DataFinal", 0, 0, {"bold": True, "font_size": 16})
    apply_format(str(path), "DataFinal", 0, 1, {"bold": True, "font_size": 16})
    apply_format(str(path), "DataFinal", 1, 0, {"italic": True, "font_size": 14})
    apply_format(str(path), "DataFinal", 2, 0, {"italic": True, "font_size": 14})
    apply_format(str(path), "DataFinal", 1, 1, {"font_size": 14})
    apply_format(str(path), "DataFinal", 2, 1, {"font_size": 14})
    set_range(str(path), "DataFinal", 4, 0, [["TOTAL (formatted)"]])
    apply_format(str(path), "DataFinal", 4, 0, {"bold": True, "font_size": 14})

    # --- Snapshot AFTER formatting changes ---
    after_path = output_dir / "calc_snapshot_after.png"
    snapshot_area(str(path), str(after_path), sheet="DataFinal")

    return {
        "spreadsheet": path,
        "pdf": output_dir / "workflow.pdf",
        "snapshot_before": before_path,
        "snapshot_after": after_path,
    }


# ---------------------------------------------------------------------------
# Deterministic assertion tests
# ---------------------------------------------------------------------------


def test_calc_snapshot_in_workflow(tmp_path):
    """Assert snapshot_area produces a valid PNG for a Calc document."""
    from calc.snapshot import snapshot_area

    outputs = run_calc_end_to_end_workflow(tmp_path)
    # The workflow already produces snapshots; verify the before snapshot
    snapshot_path = outputs["snapshot_before"]

    assert snapshot_path.exists()
    assert snapshot_path.stat().st_size > 0

    # Verify PNG magic bytes
    with open(snapshot_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    # Also verify we can take an independent snapshot
    independent_path = tmp_path / "independent_snapshot.png"
    result = snapshot_area(
        str(outputs["spreadsheet"]), str(independent_path), sheet="DataFinal"
    )
    assert independent_path.exists()
    assert result.width > 0
    assert result.height > 0
    assert result.dpi == 150


def test_calc_workflow_outputs_to_test_output_dir():
    """Produce inspectable output files in test-output/calc/.

    Calls run_calc_end_to_end_workflow which builds a Calc spreadsheet
    with data, formulas, chart, and validation, snapshots it before and
    after formatting changes. Assertions verify that all output files
    exist and the snapshots differ.

    Output files:
        test-output/calc/workflow.ods              - the spreadsheet
        test-output/calc/workflow.pdf              - PDF export
        test-output/calc/calc_snapshot_before.png  - before formatting changes
        test-output/calc/calc_snapshot_after.png   - after formatting changes
    """
    output_dir = Path("test-output/calc")

    # Clean up previous runs
    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_calc_end_to_end_workflow(output_dir)

    # Assert all output files exist and are non-empty
    for key in ("spreadsheet", "pdf", "snapshot_before", "snapshot_after"):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    # Assert before and after snapshots differ (formatting changes are visible)
    before_bytes = outputs["snapshot_before"].read_bytes()
    after_bytes = outputs["snapshot_after"].read_bytes()
    assert before_bytes != after_bytes, "Snapshots should differ after formatting"

    # Assert PNG magic bytes on both snapshots
    for key in ("snapshot_before", "snapshot_after"):
        with open(outputs[key], "rb") as f:
            assert f.read(8) == b"\x89PNG\r\n\x1a\n"
