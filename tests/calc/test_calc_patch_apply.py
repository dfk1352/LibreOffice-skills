# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from tests.calc._helpers import get_cell_number_format, named_range_exists


def test_patch_atomic_mode_success_saves_document(tmp_path):
    from calc import CalcTarget, CalcSession, patch
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "atomic_ok.ods"
    create_spreadsheet(str(doc_path))

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = add_sheet\n"
        "name = Data\n"
        "[operation]\n"
        "type = write_range\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 0\n"
        "target.col = 0\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "data <<JSON\n"
        '[["Label", "Value"], ["Revenue", 100], ["Cost", 80]]\n'
        "JSON\n"
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Data\n"
        "target.row = 3\n"
        "target.col = 1\n"
        "value = =SUM(B2:B3)\n"
        "value_type = formula\n"
        "[operation]\n"
        "type = format_range\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 1\n"
        "target.col = 1\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "format.number_format = currency\n"
        "[operation]\n"
        "type = define_named_range\n"
        "name = RevenueValues\n"
        "target.kind = range\n"
        "target.sheet = Data\n"
        "target.row = 1\n"
        "target.col = 1\n"
        "target.end_row = 2\n"
        "target.end_col = 1\n"
        "[operation]\n"
        "type = recalculate\n",
        mode="atomic",
    )

    assert result.mode == "atomic"
    assert result.overall_status == "ok"
    assert result.document_persisted is True
    assert [operation.status for operation in result.operations] == [
        "ok",
        "ok",
        "ok",
        "ok",
        "ok",
        "ok",
    ]

    with CalcSession(str(doc_path)) as session:
        total = session.read_cell(CalcTarget(kind="cell", sheet="Data", row=3, col=1))

    assert total["formula"] == "=SUM(B2:B3)"
    assert total["value"] == 180
    assert named_range_exists(doc_path, "RevenueValues")
    assert get_cell_number_format(doc_path, "Data", 1, 1) == get_cell_number_format(
        doc_path,
        "Data",
        2,
        1,
    )


def test_patch_atomic_mode_failure_rolls_back_document(tmp_path):
    from calc import CalcTarget, CalcSession, patch
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "atomic_fail.ods"
    create_spreadsheet(str(doc_path))

    with CalcSession(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "baseline",
            value_type="text",
        )

    result = patch(
        str(doc_path),
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

    assert result.mode == "atomic"
    assert result.overall_status == "failed"
    assert result.document_persisted is False
    assert [operation.status for operation in result.operations] == [
        "ok",
        "failed",
        "skipped",
    ]

    with CalcSession(str(doc_path)) as session:
        preserved = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0)
        )
        untouched = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=1, col=0)
        )

    assert preserved["value"] == "baseline"
    assert untouched["value"] == 0.0


def test_patch_best_effort_mode_records_partial_success_and_persists_mutations(
    tmp_path,
):
    from calc import CalcTarget, CalcSession, patch
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "best_effort.ods"
    create_spreadsheet(str(doc_path))

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 0\n"
        "target.col = 0\n"
        "value = kept\n"
        "value_type = text\n"
        "[operation]\n"
        "type = delete_named_range\n"
        "target.kind = named_range\n"
        "target.name = MissingRange\n"
        "[operation]\n"
        "type = write_cell\n"
        "target.kind = cell\n"
        "target.sheet = Sheet1\n"
        "target.row = 1\n"
        "target.col = 0\n"
        "value = 99\n"
        "value_type = number\n",
        mode="best_effort",
    )

    assert result.mode == "best_effort"
    assert result.overall_status == "partial"
    assert result.document_persisted is True
    assert [operation.status for operation in result.operations] == [
        "ok",
        "failed",
        "ok",
    ]

    with CalcSession(str(doc_path)) as session:
        first = session.read_cell(CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0))
        second = session.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=1, col=0)
        )

    assert first["value"] == "kept"
    assert second["value"] == 99


def test_patch_document_persisted_true_only_when_saved_changes_exist(tmp_path):
    from calc import patch
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "persisted_flag.ods"
    create_spreadsheet(str(doc_path))

    empty_result = patch(str(doc_path), "", mode="atomic")
    assert empty_result.overall_status == "ok"
    assert empty_result.operations == []
    assert empty_result.document_persisted is False

    failed_result = patch(
        str(doc_path),
        "[operation]\n"
        "type = delete_named_range\n"
        "target.kind = named_range\n"
        "target.name = MissingRange\n",
        mode="best_effort",
    )
    assert failed_result.overall_status == "partial"
    assert failed_result.document_persisted is False


def test_patch_results_preserve_original_operation_order(tmp_path):
    from calc import patch
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "operation_order.ods"
    create_spreadsheet(str(doc_path))

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = add_sheet\n"
        "name = Data\n"
        "[operation]\n"
        "type = delete_named_range\n"
        "target.kind = named_range\n"
        "target.name = MissingRange\n"
        "[operation]\n"
        "type = recalculate\n",
        mode="best_effort",
    )

    assert [operation.operation_type for operation in result.operations] == [
        "add_sheet",
        "delete_named_range",
        "recalculate",
    ]
