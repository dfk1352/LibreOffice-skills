# pyright: reportMissingImports=false, reportAttributeAccessIssue=false
"""Tests for session lifecycle: __exit__ save behaviour (#1) and close() idempotency (#9)."""

from __future__ import annotations

from pathlib import Path


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


class _DummyManager:
    def __init__(self) -> None:
        self.exit_calls: list[tuple[object, object, object]] = []

    def __exit__(self, exc_type, exc, exc_tb) -> None:
        self.exit_calls.append((exc_type, exc, exc_tb))


class _DummyDoc:
    def __init__(self) -> None:
        self.store_calls = 0
        self.close_calls: list[bool] = []

    def store(self) -> None:
        self.store_calls += 1

    def close(self, deliver_ownership: bool) -> None:
        self.close_calls.append(deliver_ownership)


def test_base_session_defines_close_reset(tmp_path):
    from session import BaseSession

    class DummySession(BaseSession):
        def __init__(self, path: Path) -> None:
            super().__init__()
            self._path = path
            self.open_count = 0
            self._open_document()

        def _open_document(self) -> None:
            self.open_count += 1
            self._uno_manager = _DummyManager()
            self._desktop = object()
            self._doc = _DummyDoc()

    snapshot_path = tmp_path / "snapshot.bin"
    snapshot_path.write_bytes(b"before")
    session = DummySession(snapshot_path)
    original_doc = session._doc
    original_manager = session._uno_manager

    assert "close" not in DummySession.__dict__
    assert "reset" not in DummySession.__dict__
    assert "restore_snapshot" not in DummySession.__dict__

    session.reset()

    assert original_doc.close_calls == [True]
    assert len(original_manager.exit_calls) == 1
    assert session.open_count == 2

    replacement_doc = session._doc
    replacement_manager = session._uno_manager
    session.restore_snapshot(b"after")

    assert replacement_doc.close_calls == [True]
    assert len(replacement_manager.exit_calls) == 1
    assert snapshot_path.read_bytes() == b"after"
    assert session.open_count == 3

    final_doc = session._doc
    session.close(save=False)

    assert final_doc.store_calls == 0
    assert final_doc.close_calls == [True]
    assert session._closed is True


def test_writer_close_always_delivers_ownership() -> None:
    from writer.session import WriterSession

    session = WriterSession.__new__(WriterSession)
    session._closed = False
    session._doc = _DummyDoc()
    session._desktop = object()
    session._uno_manager = _DummyManager()

    doc = session._doc
    session.close(save=False)

    assert doc.close_calls == [True]


def test_calc_close_always_delivers_ownership() -> None:
    from calc.session import CalcSession

    session = CalcSession.__new__(CalcSession)
    session._closed = False
    session._doc = _DummyDoc()
    session._desktop = object()
    session._uno_manager = _DummyManager()

    doc = session._doc
    session.close(save=False)

    assert doc.close_calls == [True]


def test_impress_close_always_delivers_ownership() -> None:
    from impress.session import ImpressSession

    session = ImpressSession.__new__(ImpressSession)
    session._closed = False
    session._doc = _DummyDoc()
    session._desktop = object()
    session._uno_manager = _DummyManager()

    doc = session._doc
    session.close(save=False)

    assert doc.close_calls == [True]
