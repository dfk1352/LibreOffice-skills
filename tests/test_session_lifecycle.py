# pyright: reportMissingImports=false, reportAttributeAccessIssue=false
"""Tests for session lifecycle: __exit__ save behaviour (#1) and close() idempotency (#9)."""

from __future__ import annotations

import pytest


def test_exit_on_exception_does_not_save(tmp_path):
    """When an exception is raised inside a with-block, changes must not persist (#1)."""
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "no_save.odt"
    create_document(str(doc_path))

    try:
        with WriterSession(str(doc_path)) as session:
            session.insert_text("should not persist")
            raise RuntimeError("deliberate error")
    except RuntimeError:
        pass

    with WriterSession(str(doc_path)) as reopened:
        assert reopened.read_text() == ""


def test_exit_normal_saves(tmp_path):
    """Normal exit from a with-block persists changes."""
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "yes_save.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("should persist")

    with WriterSession(str(doc_path)) as reopened:
        assert reopened.read_text() == "should persist"


def test_close_idempotent_writer(tmp_path):
    """Calling close() twice on WriterSession must not raise (#9)."""
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "idempotent.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.close()
    session.close()  # second call must not raise


def test_close_idempotent_calc(tmp_path):
    """Calling close() twice on CalcSession must not raise (#9)."""
    from calc import CalcSession
    from calc.core import create_spreadsheet

    doc_path = tmp_path / "idempotent.ods"
    create_spreadsheet(str(doc_path))

    session = CalcSession(str(doc_path))
    session.close()
    session.close()


def test_close_idempotent_impress(tmp_path):
    """Calling close() twice on ImpressSession must not raise (#9)."""
    from impress import ImpressSession
    from impress.core import create_presentation

    doc_path = tmp_path / "idempotent.odp"
    create_presentation(str(doc_path))

    session = ImpressSession(str(doc_path))
    session.close()
    session.close()


def test_with_block_manual_close_then_exit_writer(tmp_path):
    """Manual close() inside a with-block followed by __exit__ must not raise (#9)."""
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "manual_close.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("hello")
        session.close(save=True)
    # __exit__ calls close() again — should be a no-op


def test_exit_on_exception_calc(tmp_path):
    """Calc: exception inside with-block discards changes (#1)."""
    from calc import CalcSession
    from calc.core import create_spreadsheet
    from calc.targets import CalcTarget

    doc_path = tmp_path / "no_save.ods"
    create_spreadsheet(str(doc_path))

    try:
        with CalcSession(str(doc_path)) as session:
            session.write_cell(
                CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
                "should not persist",
                value_type="text",
            )
            raise RuntimeError("deliberate error")
    except RuntimeError:
        pass

    with CalcSession(str(doc_path)) as reopened:
        result = reopened.read_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0)
        )
        # Cell should be empty (not persisted)
        assert (
            result["value"] is None or result["value"] == "" or result["value"] == 0.0
        )


def test_exit_on_exception_impress(tmp_path):
    """Impress: exception inside with-block discards changes (#1)."""
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation

    doc_path = tmp_path / "no_save.odp"
    create_presentation(str(doc_path))

    original_bytes = doc_path.read_bytes()

    try:
        with ImpressSession(str(doc_path)) as session:
            session.insert_text_box(
                ImpressTarget(kind="slide", slide_index=0),
                "should not persist",
                ShapePlacement(x_cm=1, y_cm=1, width_cm=10, height_cm=3),
            )
            raise RuntimeError("deliberate error")
    except RuntimeError:
        pass

    # File should be unchanged
    assert doc_path.read_bytes() == original_bytes


# --- Atomic patch rollback tests (#8) ---


def test_atomic_rollback_preserves_unsaved_edits(tmp_path):
    """Atomic rollback must snapshot the live in-memory state, not stale disk (#8).

    If unsaved edits exist when patch() is called with mode=atomic and the
    patch fails, rollback should restore the pre-patch in-memory state
    (including the unsaved edits), not the on-disk state from before session.
    """
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "atomic_rollback.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        # Make an edit but do NOT save
        session.insert_text("pre-patch edit")

        # Apply a patch that will fail (invalid operation target)
        result = session.patch(
            "[operation]\n"
            "type = delete_text\n"
            "target.kind = text\n"
            "target.text = nonexistent text that does not exist anywhere\n",
            mode="atomic",
        )

        assert result.overall_status == "failed"
        # After rollback, the pre-patch edit should still be present
        text = session.read_text()
        assert "pre-patch edit" in text
