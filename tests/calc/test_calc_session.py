"""Tests for Calc session lifecycle behaviour."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_open_calc_session_returns_session(tmp_path):
    from calc import CalcSession, open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "session.ods"
    create_spreadsheet(str(doc_path))

    session = open_calc_session(str(doc_path))
    try:
        assert isinstance(session, CalcSession)
    finally:
        session.close()


def test_open_calc_session_missing_path_raises(tmp_path):
    from calc import open_calc_session
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        open_calc_session(str(tmp_path / "missing.ods"))


def test_session_close_save_true_persists_changes(tmp_path):
    from calc import open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "persist.ods"
    create_spreadsheet(str(doc_path))

    session = open_calc_session(str(doc_path))
    session.write_cell(_cell_target(), "Saved once", value_type="text")
    session.close(save=True)

    with open_calc_session(str(doc_path)) as reopened:
        cell = reopened.read_cell(_cell_target())

    assert cell["value"] == "Saved once"
    assert cell["type"] == "text"


def test_session_close_save_false_discards_changes(tmp_path):
    from calc import open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "discard.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as initial_session:
        initial_session.write_cell(_cell_target(), "Persisted", value_type="text")

    session = open_calc_session(str(doc_path))
    session.write_cell(_cell_target(), "Discard me", value_type="text")
    session.close(save=False)

    with open_calc_session(str(doc_path)) as reopened:
        cell = reopened.read_cell(_cell_target())

    assert cell["value"] == "Persisted"


def test_session_reset_discards_in_memory_changes_and_reloads_saved_state(tmp_path):
    from calc import open_calc_session
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "reset.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as initial_session:
        initial_session.write_cell(_cell_target(), "Persisted", value_type="text")

    with open_calc_session(str(doc_path)) as session:
        session.write_cell(_cell_target(), "Draft", value_type="text")
        session.reset()
        cell = session.read_cell(_cell_target())

    assert cell["value"] == "Persisted"


@pytest.mark.parametrize(
    ("label", "call"),
    [
        ("read_cell", lambda session, tmp_path: session.read_cell(_cell_target())),
        (
            "write_cell",
            lambda session, tmp_path: session.write_cell(
                _cell_target(),
                "after close",
                value_type="text",
            ),
        ),
        (
            "read_range",
            lambda session, tmp_path: session.read_range(_range_target()),
        ),
        (
            "write_range",
            lambda session, tmp_path: session.write_range(_range_target(), [[1, 2]]),
        ),
        (
            "format_range",
            lambda session, tmp_path: session.format_range(
                _range_target(),
                _formatting(),
            ),
        ),
        ("list_sheets", lambda session, tmp_path: session.list_sheets()),
        (
            "add_sheet",
            lambda session, tmp_path: session.add_sheet("Data"),
        ),
        (
            "rename_sheet",
            lambda session, tmp_path: session.rename_sheet(
                _sheet_target(),
                "Renamed",
            ),
        ),
        (
            "delete_sheet",
            lambda session, tmp_path: session.delete_sheet(_sheet_target()),
        ),
        (
            "define_named_range",
            lambda session, tmp_path: session.define_named_range(
                "Values",
                _range_target(),
            ),
        ),
        (
            "get_named_range",
            lambda session, tmp_path: session.get_named_range(_named_range_target()),
        ),
        (
            "delete_named_range",
            lambda session, tmp_path: session.delete_named_range(_named_range_target()),
        ),
        (
            "set_validation",
            lambda session, tmp_path: session.set_validation(
                _range_target(),
                _validation_rule(),
            ),
        ),
        (
            "clear_validation",
            lambda session, tmp_path: session.clear_validation(_range_target()),
        ),
        (
            "create_chart",
            lambda session, tmp_path: session.create_chart(
                _sheet_target(),
                _chart_spec(),
            ),
        ),
        (
            "update_chart",
            lambda session, tmp_path: session.update_chart(
                _chart_target(),
                _chart_spec(),
            ),
        ),
        (
            "delete_chart",
            lambda session, tmp_path: session.delete_chart(_chart_target()),
        ),
        ("recalculate", lambda session, tmp_path: session.recalculate()),
        (
            "export",
            lambda session, tmp_path: session.export(
                str(tmp_path / "closed.pdf"),
                "pdf",
            ),
        ),
        (
            "patch",
            lambda session, tmp_path: session.patch(
                "[operation]\ntype = recalculate\n"
            ),
        ),
        ("reset", lambda session, tmp_path: session.reset()),
    ],
)
def test_closed_session_methods_raise_calc_session_error(tmp_path, label, call):
    from calc import open_calc_session
    from calc.core import create_spreadsheet
    from calc.exceptions import CalcSessionError

    doc_path = tmp_path / f"closed_{label}.ods"
    create_spreadsheet(str(doc_path))

    session = open_calc_session(str(doc_path))
    session.close()

    with pytest.raises(CalcSessionError):
        call(session, tmp_path)


def test_session_close_twice_raises_calc_session_error(tmp_path):
    from calc import open_calc_session
    from calc.core import create_spreadsheet
    from calc.exceptions import CalcSessionError

    doc_path = tmp_path / "double_close.ods"
    create_spreadsheet(str(doc_path))

    session = open_calc_session(str(doc_path))
    session.close()

    with pytest.raises(CalcSessionError):
        session.close()


def test_open_calc_session_context_manager_closes_after_block(tmp_path):
    from calc import open_calc_session
    from calc.core import create_spreadsheet
    from calc.exceptions import CalcSessionError

    doc_path = tmp_path / "context.ods"
    create_spreadsheet(str(doc_path))

    with open_calc_session(str(doc_path)) as session:
        session.write_cell(_cell_target(), "context managed", value_type="text")

    with pytest.raises(CalcSessionError):
        session.list_sheets()

    with open_calc_session(str(doc_path)) as reopened:
        cell = reopened.read_cell(_cell_target())

    assert cell["value"] == "context managed"


def _cell_target():
    from calc import CalcTarget

    return CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0)


def _range_target():
    from calc import CalcTarget

    return CalcTarget(kind="range", sheet="Sheet1", row=0, col=0, end_row=0, end_col=1)


def _sheet_target():
    from calc import CalcTarget

    return CalcTarget(kind="sheet", sheet="Sheet1")


def _named_range_target():
    from calc import CalcTarget

    return CalcTarget(kind="named_range", name="Values")


def _chart_target():
    from calc import CalcTarget

    return CalcTarget(kind="chart", sheet="Sheet1", index=0)


def _formatting():
    from calc import CellFormatting

    return CellFormatting(bold=True)


def _validation_rule():
    from calc import ValidationRule

    return ValidationRule(type="whole", condition="between", value1=1, value2=10)


def _chart_spec():
    from calc import CalcTarget, ChartSpec

    return ChartSpec(
        chart_type="line",
        data_range=CalcTarget(
            kind="range",
            sheet="Sheet1",
            row=0,
            col=0,
            end_row=1,
            end_col=1,
        ),
        anchor_row=3,
        anchor_col=0,
        width=4000,
        height=2500,
        title="Example",
    )
