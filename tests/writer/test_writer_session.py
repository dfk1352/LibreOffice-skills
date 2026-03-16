"""Tests for Writer session lifecycle behaviour."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import create_test_image


def test_open_writer_session_returns_session(tmp_path):
    from writer import WriterSession, open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "session.odt"
    create_document(str(doc_path))

    session = open_writer_session(str(doc_path))
    try:
        assert isinstance(session, WriterSession)
    finally:
        session.close()


def test_open_writer_session_missing_path_raises(tmp_path):
    from writer import open_writer_session
    from writer.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        open_writer_session(str(tmp_path / "missing.odt"))


def test_session_close_save_true_persists_changes(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "persist.odt"
    create_document(str(doc_path))

    session = open_writer_session(str(doc_path))
    session.insert_text("Saved once")
    session.close(save=True)

    with open_writer_session(str(doc_path)) as reopened:
        assert reopened.read_text() == "Saved once"


def test_session_close_save_false_discards_changes(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "discard.odt"
    create_document(str(doc_path))

    session = open_writer_session(str(doc_path))
    session.insert_text("Discard me")
    session.close(save=False)

    with open_writer_session(str(doc_path)) as reopened:
        assert reopened.read_text() == ""


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
    tmp_path,
    label,
    call,
):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / f"closed_{label}.odt"
    image_path = create_test_image(tmp_path / f"{label}.png")
    create_document(str(doc_path))

    session = open_writer_session(str(doc_path))
    session.close()

    with pytest.raises(WriterSessionError):
        call(session, image_path)


def test_session_close_twice_raises_writer_session_error(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "double_close.odt"
    create_document(str(doc_path))

    session = open_writer_session(str(doc_path))
    session.close()

    with pytest.raises(WriterSessionError):
        session.close()


def test_open_writer_session_context_manager_closes_after_block(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "context.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_text("context managed")

    with pytest.raises(WriterSessionError):
        session.read_text()


def test_open_writer_session_context_manager_closes_on_exception(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import WriterSessionError

    doc_path = tmp_path / "context_exception.odt"
    create_document(str(doc_path))

    session = None

    with pytest.raises(RuntimeError):
        with open_writer_session(str(doc_path)) as managed_session:
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
