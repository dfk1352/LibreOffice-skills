# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from tests.writer._helpers import (
    ARABIC_NUMBERING_TYPE,
    BULLET_NUMBERING_TYPE,
    assert_list_items,
    assert_text_formatting,
)


def _create_seed_document(doc_path):
    from writer import ListItem, WriterTarget, WriterSession
    from writer.core import create_document

    create_document(str(doc_path))
    with WriterSession(str(doc_path)) as session:
        session.insert_text(
            "Financial Summary\n\n"
            "Quarterly revenue grew 18%.\n\n"
            "Action Items\n\n"
            "Risks\n\n"
            "Tail section."
        )
        session.insert_list(
            [ListItem(text="Legacy item", level=0)],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Action Items"),
        )


def test_patch_atomic_mode_success_saves_document(tmp_path):
    from writer import WriterSession, patch

    doc_path = tmp_path / "atomic_ok.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        "target.kind = insertion\n"
        "target.after = Financial Summary\n"
        "text = Inserted paragraph.\n"
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Quarterly revenue grew 18%.\n"
        "format.bold = true\n"
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Action Items\n"
        "list.ordered = true\n"
        'items = [{"text": "Confirm scope", "level": 0}, {"text": "Review output", "level": 0}]\n',
        mode="atomic",
    )

    assert result.mode == "atomic"
    assert result.overall_status == "ok"
    assert result.document_persisted is True
    assert [operation.status for operation in result.operations] == ["ok", "ok", "ok"]

    with WriterSession(str(doc_path)) as session:
        text = session.read_text()

    assert "Inserted paragraph." in text
    assert "Quarterly revenue grew 18%." in text
    assert_text_formatting(doc_path, "Quarterly revenue grew 18%.", char_weight=150.0)
    assert_list_items(
        doc_path,
        ["Legacy item", "Confirm scope", "Review output"],
        expected_levels=[0, 0, 0],
        expected_numbering_type=ARABIC_NUMBERING_TYPE,
    )


def test_patch_atomic_mode_failure_rolls_back_document(tmp_path):
    from writer import WriterSession, patch

    doc_path = tmp_path / "atomic_fail.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_text\n"
        "target.kind = insertion\n"
        "target.after = Financial Summary\n"
        "text = Should not persist.\n"
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = missing text\n"
        "format.italic = true\n"
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Risks\n"
        "list.ordered = false\n"
        'items = [{"text": "Skipped later", "level": 0}]\n',
        mode="atomic",
    )

    assert result.mode == "atomic"
    assert result.overall_status == "failed"
    assert result.document_persisted is False
    assert result.operations[0].status == "ok"
    assert result.operations[1].status == "failed"
    assert result.operations[2].status == "skipped"

    with WriterSession(str(doc_path)) as session:
        text = session.read_text()

    assert "Should not persist." not in text
    assert "Quarterly revenue grew 18%." in text
    assert_list_items(
        doc_path,
        ["Legacy item"],
        expected_levels=[0],
        expected_numbering_type=BULLET_NUMBERING_TYPE,
    )


def test_patch_best_effort_mode_records_partial_success(tmp_path):
    from writer import WriterSession, patch

    doc_path = tmp_path / "best_effort.odt"
    _create_seed_document(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Risks\n"
        "list.ordered = false\n"
        'items = [{"text": "Best effort item", "level": 0}]\n'
        "[operation]\n"
        "type = replace_list\n"
        "target.kind = list\n"
        "target.text = missing item\n"
        'items = [{"text": "never applied", "level": 0}]\n'
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Tail section.\n"
        "format.underline = true\n",
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

    with WriterSession(str(doc_path)) as session:
        text = session.read_text()

    assert "Best effort item" in text
    assert "Tail section." in text
    assert_list_items(
        doc_path,
        ["Legacy item", "Best effort item"],
        expected_levels=[0, 0],
    )
    assert_text_formatting(doc_path, "Tail section.", char_underline=1)


def test_patch_result_document_persisted_true_only_when_changes_saved(tmp_path):
    from writer import WriterTarget, WriterSession, patch
    from writer.core import create_document

    doc_path = tmp_path / "empty_patch.odt"
    create_document(str(doc_path))

    result = patch(str(doc_path), "", mode="atomic")

    assert result.overall_status == "ok"
    assert result.operations == []
    assert result.document_persisted is False

    failed_result = patch(
        str(doc_path),
        "[operation]\n"
        "type = delete_list\n"
        "target.kind = list\n"
        "target.text = missing item\n",
        mode="best_effort",
    )

    assert failed_result.overall_status == "partial"
    assert failed_result.document_persisted is False

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Financial Summary\n\nQuarterly revenue grew 18%.")
        session_result = session.patch(
            "[operation]\n"
            "type = format_text\n"
            "target.kind = text\n"
            "target.text = Quarterly revenue grew 18%.\n"
            "format.bold = true\n",
            mode="atomic",
        )
        assert session_result.overall_status == "ok"
        assert session_result.document_persisted is False
        assert session.read_text(
            WriterTarget(kind="text", text="Quarterly revenue grew 18%.")
        ) == ("Quarterly revenue grew 18%.")


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
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Quarterly revenue grew 18%.\n"
        "format.bold = true\n"
        "[operation]\n"
        "type = delete_list\n"
        "target.kind = list\n"
        "target.text = Legacy item\n",
        mode="best_effort",
    )

    assert [operation.operation_type for operation in result.operations] == [
        "insert_text",
        "format_text",
        "delete_list",
    ]
