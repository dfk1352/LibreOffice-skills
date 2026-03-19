"""Tests for applying Impress patches."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from tests.impress._helpers import (
    TITLE_AND_CONTENT_LAYOUT,
    add_text_shape,
    append_slide,
    create_test_audio,
    create_test_video,
    get_list_paragraphs,
    get_notes_text,
    get_shape_text,
    open_impress_doc,
    resolve_shape,
)


def _create_seed_presentation(doc_path):
    from impress.core import create_presentation

    create_presentation(str(doc_path))

    with open_impress_doc(doc_path) as doc:
        slide = append_slide(doc, TITLE_AND_CONTENT_LAYOUT)
        title_shape = resolve_shape(slide, placeholder="title")
        body_shape = resolve_shape(slide, placeholder="body")
        assert title_shape is not None
        assert body_shape is not None
        title_shape.setString("Q1 Summary")
        body_shape.setString("Quarterly revenue rose 18%.")
        add_text_shape(
            doc,
            slide,
            "Action Items",
            name="Agenda Box",
            x_cm=1.5,
            y_cm=8.0,
            width_cm=10.0,
            height_cm=4.0,
        )
        doc.store()


def test_patch_atomic_mode_success_saves_document(tmp_path):
    from impress import patch

    doc_path = tmp_path / "atomic_ok.odp"
    _create_seed_presentation(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.slide_index = 1\n"
        "target.placeholder = body\n"
        "new_text = Quarterly revenue rose 21%.\n"
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.slide_index = 1\n"
        "target.placeholder = body\n"
        "format.bold = true\n"
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.slide_index = 1\n"
        "target.shape_name = Agenda Box\n"
        "target.after = Action Items\n"
        "list.ordered = true\n"
        'items = [{"text": "Confirm scope", "level": 0}, {"text": "Review outputs", "level": 0}]\n'
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Ready to present\n",
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
    ]

    assert get_shape_text(doc_path, 1, placeholder="body") == (
        "Quarterly revenue rose 21%."
    )
    assert [
        item["text"] for item in get_list_paragraphs(doc_path, 1, name="Agenda Box")
    ] == [
        "Confirm scope",
        "Review outputs",
    ]
    assert get_notes_text(doc_path, 1) == "Ready to present"


def test_patch_atomic_mode_failure_rolls_back_document(tmp_path):
    from impress import patch

    doc_path = tmp_path / "atomic_fail.odp"
    _create_seed_presentation(doc_path)

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.slide_index = 1\n"
        "target.placeholder = body\n"
        "new_text = Should roll back\n"
        "[operation]\n"
        "type = delete_item\n"
        "target.kind = chart\n"
        "target.slide_index = 1\n"
        "target.shape_name = Missing Chart\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Skipped later\n",
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

    assert get_shape_text(doc_path, 1, placeholder="body") == (
        "Quarterly revenue rose 18%."
    )
    assert get_notes_text(doc_path, 1) == ""


def test_patch_best_effort_mode_records_partial_success_and_persists_mutations(
    tmp_path,
):
    from impress import patch

    doc_path = tmp_path / "best_effort.odp"
    _create_seed_presentation(doc_path)
    replacement_video = create_test_video(tmp_path / "replacement.mp4")

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.slide_index = 1\n"
        "target.shape_name = Agenda Box\n"
        "target.after = Action Items\n"
        "list.ordered = false\n"
        'items = [{"text": "Best effort item", "level": 0}]\n'
        "[operation]\n"
        "type = replace_media\n"
        "target.kind = media\n"
        "target.slide_index = 1\n"
        "target.shape_name = Missing Media\n"
        f"media_path = {replacement_video}\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Best effort notes\n",
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

    assert [
        item["text"] for item in get_list_paragraphs(doc_path, 1, name="Agenda Box")
    ] == ["Best effort item"]
    assert get_notes_text(doc_path, 1) == "Best effort notes"


def test_patch_document_persisted_true_only_when_saved_changes_exist(tmp_path):
    from impress import patch

    doc_path = tmp_path / "persisted_flag.odp"
    _create_seed_presentation(doc_path)

    empty_result = patch(str(doc_path), "", mode="atomic")
    assert empty_result.overall_status == "ok"
    assert empty_result.operations == []
    assert empty_result.document_persisted is False

    failed_result = patch(
        str(doc_path),
        "[operation]\n"
        "type = delete_item\n"
        "target.kind = list\n"
        "target.slide_index = 1\n"
        "target.shape_name = Agenda Box\n"
        "target.text = Missing item\n",
        mode="best_effort",
    )
    assert failed_result.overall_status == "partial"
    assert failed_result.document_persisted is False


def test_patch_results_preserve_original_operation_order(tmp_path):
    from impress import patch

    doc_path = tmp_path / "operation_order.odp"
    _create_seed_presentation(doc_path)
    replacement_audio = create_test_audio(tmp_path / "replacement.wav")

    result = patch(
        str(doc_path),
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.slide_index = 1\n"
        "target.shape_name = Agenda Box\n"
        "list.ordered = false\n"
        'items = [{"text": "First", "level": 0}]\n'
        "[operation]\n"
        "type = replace_media\n"
        "target.kind = media\n"
        "target.slide_index = 1\n"
        "target.shape_name = Missing Media\n"
        f"media_path = {replacement_audio}\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Ordered results\n",
        mode="best_effort",
    )

    assert [operation.operation_type for operation in result.operations] == [
        "insert_list",
        "replace_media",
        "set_notes",
    ]
