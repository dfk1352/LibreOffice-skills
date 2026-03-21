# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import create_test_image


@pytest.fixture(scope="module")
def closed_writer_session(tmp_path_factory):
    """A WriterSession that has been opened and immediately closed.

    Module-scoped so a single LibreOffice round-trip is shared across all
    parametrized ``test_closed_session_methods_raise_writer_session_error``
    variants. Each variant only checks that calling a method on the
    already-closed session raises ``WriterSessionError`` -- no UNO access
    is needed for that assertion.

    Returns a ``(session, tmp_path)`` tuple so parametrized lambdas that
    reference file paths still work.
    """
    from writer import WriterSession
    from writer.core import create_document

    base = tmp_path_factory.mktemp("closed_writer")
    doc_path = base / "closed.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.close()
    return session, base


def test_writer_session_returns_session(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "session.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    try:
        assert isinstance(session, WriterSession)
    finally:
        session.close()


def test_writer_session_missing_path_raises(tmp_path):
    from writer import WriterSession
    from writer.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        WriterSession(str(tmp_path / "missing.odt"))


def test_session_close_save_true_persists_changes(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "persist.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.insert_text("Saved once")
    session.close(save=True)

    with WriterSession(str(doc_path)) as reopened:
        assert reopened.read_text() == "Saved once"


def test_session_close_save_false_discards_changes(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "discard.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.insert_text("Discard me")
    session.close(save=False)

    with WriterSession(str(doc_path)) as reopened:
        assert reopened.read_text() == ""


def test_session_restore_snapshot_reverts_to_original_content(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "restore.odt"
    create_document(str(doc_path))

    original_bytes = doc_path.read_bytes()

    session = WriterSession(str(doc_path))
    session.insert_text("Temporary content that should be reverted")
    session.restore_snapshot(original_bytes)
    assert session.read_text() == ""
    session.close(save=False)


def test_session_restore_snapshot_marks_closed_on_reopen_failure(tmp_path):
    from unittest.mock import patch as mock_patch

    from writer import WriterSession
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "restore_fail.odt"
    create_document(str(doc_path))

    original_bytes = doc_path.read_bytes()
    session = WriterSession(str(doc_path))

    with mock_patch.object(
        WriterSession, "_open_document", side_effect=RuntimeError("simulated failure")
    ):
        with pytest.raises(RuntimeError, match="simulated failure"):
            session.restore_snapshot(original_bytes)

    with pytest.raises(WriterSessionError):
        session.read_text()


@pytest.mark.parametrize(
    ("label", "call"),
    [
        ("read_text", lambda session, image_path: session.read_text()),
        (
            "insert_text",
            lambda session, image_path: session.insert_text("after close"),
        ),
        (
            "replace_text",
            lambda session, image_path: session.replace_text(
                _text_target(),
                "updated",
            ),
        ),
        (
            "delete_text",
            lambda session, image_path: session.delete_text(_text_target()),
        ),
        (
            "format_text",
            lambda session, image_path: session.format_text(
                _text_target(),
                _formatting(),
            ),
        ),
        (
            "insert_table",
            lambda session, image_path: session.insert_table(1, 1, [["x"]], "T1"),
        ),
        (
            "update_table",
            lambda session, image_path: session.update_table(_table_target(), [["x"]]),
        ),
        (
            "delete_table",
            lambda session, image_path: session.delete_table(_table_target()),
        ),
        (
            "insert_image",
            lambda session, image_path: session.insert_image(
                str(image_path), name="Logo"
            ),
        ),
        (
            "update_image",
            lambda session, image_path: session.update_image(
                _image_target(),
                image_path=str(image_path),
            ),
        ),
        (
            "delete_image",
            lambda session, image_path: session.delete_image(_image_target()),
        ),
        (
            "insert_list",
            lambda session, image_path: session.insert_list(
                _list_items(), ordered=False
            ),
        ),
        (
            "replace_list",
            lambda session, image_path: session.replace_list(
                _list_target(),
                _list_items(),
                ordered=True,
            ),
        ),
        (
            "delete_list",
            lambda session, image_path: session.delete_list(_list_target()),
        ),
        (
            "patch",
            lambda session, image_path: session.patch(
                "[operation]\ntype = insert_text\ntext = ignored after close\n"
            ),
        ),
        (
            "export",
            lambda session, image_path: session.export(
                str(image_path.with_suffix(".pdf")),
                "pdf",
            ),
        ),
    ],
)
def test_closed_session_methods_raise_writer_session_error(
    closed_writer_session,
    label,
    call,
):
    from writer.exceptions import WriterSessionError

    session, base_tmp = closed_writer_session
    image_path = base_tmp / f"{label}.png"
    if not image_path.exists():
        create_test_image(image_path)

    with pytest.raises(WriterSessionError):
        call(session, image_path)


def test_session_close_twice_raises_writer_session_error(tmp_path):
    from writer import WriterSession
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "double_close.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.close()

    with pytest.raises(WriterSessionError):
        session.close()


def test_writer_session_context_manager_closes_after_block(tmp_path):
    from writer import WriterSession
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "context.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("context managed")

    with pytest.raises(WriterSessionError):
        session.read_text()


def test_writer_session_context_manager_closes_on_exception(tmp_path):
    from writer import WriterSession
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "context_exception.odt"
    create_document(str(doc_path))

    session = None

    with pytest.raises(RuntimeError):
        with WriterSession(str(doc_path)) as managed_session:
            session = managed_session
            managed_session.insert_text("before boom")
            raise RuntimeError("boom")

    assert session is not None
    with pytest.raises(WriterSessionError):
        session.read_text()


def _text_target():
    from writer import WriterTarget

    return WriterTarget(kind="text", text="seed")


def _table_target():
    from writer import WriterTarget

    return WriterTarget(kind="table", name="T1")


def _image_target():
    from writer import WriterTarget

    return WriterTarget(kind="image", name="Logo")


def _list_target():
    from writer import WriterTarget

    return WriterTarget(kind="list", text="seed")


def _formatting():
    from writer import TextFormatting

    return TextFormatting(bold=True)


def _list_items():
    from writer import ListItem

    return [ListItem(text="seed", level=0)]
