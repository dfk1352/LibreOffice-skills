"""Tests for applying Writer patches."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations


def _create_seed_document(doc_path):
    from writer import open_writer_session
    from writer.core import create_document

    create_document(str(doc_path))
    with open_writer_session(str(doc_path)) as session:
        session.insert_text("Introduction\n\nOld sentence.\n\nTail section.")


def test_patch_atomic_mode_success_saves_document(tmp_path):
    from writer import open_writer_session, patch

    doc_path = tmp_path / "atomic_ok.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        'selector = after:"Introduction"\n'
        "text = Inserted paragraph.\n"
        "[operation]\n"
        "type = replace_text\n"
        'selector = contains:"Old sentence."\n'
        "new_text = Updated sentence.\n",
        mode="atomic",
    )

    assert result.mode == "atomic"
    assert result.overall_status == "ok"
    assert result.document_persisted is True
    assert [operation.status for operation in result.operations] == ["ok", "ok"]

    with open_writer_session(str(doc_path)) as session:
        text = session.read_text()

    assert "Introduction\nInserted paragraph." in text
    assert "Inserted paragraph." in text
    assert "Updated sentence." in text
    assert "Old sentence." not in text


def test_patch_atomic_mode_failure_rolls_back_document(tmp_path):
    from writer import open_writer_session, patch

    doc_path = tmp_path / "atomic_fail.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        'selector = after:"Introduction"\n'
        "text = Should not persist.\n"
        "[operation]\n"
        "type = delete_text\n"
        'selector = contains:"does not exist"\n',
        mode="atomic",
    )

    assert result.mode == "atomic"
    assert result.overall_status == "failed"
    assert result.document_persisted is False
    assert result.operations[0].status == "ok"
    assert result.operations[1].status == "failed"

    with open_writer_session(str(doc_path)) as session:
        text = session.read_text()

    assert "Should not persist." not in text
    assert text == "Introduction\n\nOld sentence.\n\nTail section."


def test_patch_best_effort_mode_records_partial_success(tmp_path):
    from writer import open_writer_session, patch

    doc_path = tmp_path / "best_effort.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        'selector = after:"Introduction"\n'
        "text = First success.\n"
        "[operation]\n"
        "type = delete_text\n"
        'selector = contains:"missing body"\n'
        "[operation]\n"
        "type = replace_text\n"
        'selector = contains:"Tail section."\n'
        "new_text = Final section.\n",
        mode="best_effort",
    )

    assert result.mode == "best_effort"
    assert result.overall_status == "partial"
    assert [operation.status for operation in result.operations] == [
        "ok",
        "failed",
        "ok",
    ]
    assert result.document_persisted is True

    with open_writer_session(str(doc_path)) as session:
        text = session.read_text()

    assert "Introduction\nFirst success." in text
    assert "First success." in text
    assert "Final section." in text
    assert "Tail section." not in text


def test_patch_result_document_persisted_false_when_nothing_mutates(tmp_path):
    from writer import patch
    from writer.core import create_document

    doc_path = tmp_path / "empty_patch.odt"
    create_document(str(doc_path))

    result = patch(str(doc_path), "", mode="atomic")

    assert result.overall_status == "ok"
    assert result.operations == []
    assert result.document_persisted is False


def test_patch_result_preserves_original_operation_order(tmp_path):
    from writer import patch

    doc_path = tmp_path / "operation_order.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        "text = First\n"
        "[operation]\n"
        "type = replace_text\n"
        'selector = contains:"Old sentence."\n'
        "new_text = Second\n"
        "[operation]\n"
        "type = delete_text\n"
        'selector = contains:"Tail section."\n',
        mode="best_effort",
    )

    assert [operation.operation_type for operation in result.operations] == [
        "insert_text",
        "replace_text",
        "delete_text",
    ]
