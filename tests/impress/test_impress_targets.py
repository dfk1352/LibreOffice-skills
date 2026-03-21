# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.impress._helpers import (
    TITLE_AND_CONTENT_LAYOUT,
    TITLE_SLIDE_LAYOUT,
    add_text_shape,
    append_slide,
    apply_list_to_paragraphs,
    open_impress_doc,
    resolve_shape,
    set_notes_text,
)


def _create_target_fixture(doc_path):
    from impress.core import create_presentation

    create_presentation(str(doc_path))

    with open_impress_doc(doc_path) as doc:
        title_slide = append_slide(doc, TITLE_SLIDE_LAYOUT)
        body_slide = append_slide(doc, TITLE_AND_CONTENT_LAYOUT)
        fixture_slide = append_slide(doc)

        title_shape = resolve_shape(title_slide, placeholder="title")
        assert title_shape is not None
        title_shape.setString("Quarterly Planning")

        body_title = resolve_shape(body_slide, placeholder="title")
        body_placeholder = resolve_shape(body_slide, placeholder="body")
        assert body_title is not None
        assert body_placeholder is not None
        body_title.setString("Q1 Summary")
        body_placeholder.setString("Placeholder body text")

        add_text_shape(
            doc,
            fixture_slide,
            "Lead in\nAnchor start\nShared phrase\nTarget sentence\n"
            "Shared phrase\nAnchor end\nTail",
            name="Copy Box",
            x_cm=1.5,
            y_cm=1.5,
            width_cm=12.0,
            height_cm=6.0,
        )
        agenda_box = add_text_shape(
            doc,
            fixture_slide,
            "Action Items\nConfirm scope\nReview outputs\nRisks",
            name="Agenda Box",
            x_cm=1.5,
            y_cm=8.0,
            width_cm=12.0,
            height_cm=5.0,
        )
        apply_list_to_paragraphs(
            agenda_box,
            ["Confirm scope", "Review outputs"],
            ordered=False,
        )
        set_notes_text(fixture_slide, "Speaker notes mention review outputs")
        doc.store()


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (
            {"kind": "slide", "slide_index": 2},
            {"kind": "slide", "slide_index": 2},
        ),
        (
            {
                "kind": "shape",
                "slide_index": 1,
                "placeholder": "title",
            },
            {"kind": "shape", "slide_index": 1, "placeholder": "title"},
        ),
        (
            {
                "kind": "text",
                "slide_index": 3,
                "shape_name": "Copy Box",
                "text": "Target sentence",
                "after": "Anchor start",
                "before": "Anchor end",
                "occurrence": 0,
            },
            {
                "kind": "text",
                "slide_index": 3,
                "shape_name": "Copy Box",
                "text": "Target sentence",
                "after": "Anchor start",
                "before": "Anchor end",
                "occurrence": 0,
            },
        ),
        (
            {"kind": "notes", "slide_index": 3},
            {"kind": "notes", "slide_index": 3},
        ),
        (
            {
                "kind": "list",
                "slide_index": 3,
                "shape_name": "Agenda Box",
                "text": "Review outputs",
            },
            {
                "kind": "list",
                "slide_index": 3,
                "shape_name": "Agenda Box",
                "text": "Review outputs",
            },
        ),
        (
            {"kind": "master_page", "master_name": "Default"},
            {"kind": "master_page", "master_name": "Default"},
        ),
    ],
)
def test_parse_target_accepts_valid_impress_target_forms(raw, expected):
    from impress.targets import parse_target

    target = parse_target(raw)

    for key, value in expected.items():
        assert getattr(target, key) == value


def test_parse_target_rejects_shape_name_and_shape_index_together():
    from impress.exceptions import InvalidTargetError
    from impress.targets import parse_target

    with pytest.raises(InvalidTargetError):
        parse_target(
            {
                "kind": "shape",
                "slide_index": 1,
                "shape_name": "Agenda Box",
                "shape_index": 0,
            }
        )


def test_resolve_shape_target_resolves_title_placeholder_without_guessing_body(
    tmp_path,
):
    from impress import ImpressTarget
    from impress.targets import resolve_shape_target

    doc_path = tmp_path / "resolve_shape.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        shape = resolve_shape_target(
            ImpressTarget(kind="shape", slide_index=1, placeholder="title"),
            doc,
        )
        resolved_text = shape.getString()

    assert resolved_text == "Quarterly Planning"


def test_resolve_text_range_supports_after_and_before_matching_inside_text_box(
    tmp_path,
):
    from impress import ImpressTarget
    from impress.targets import resolve_text_range

    doc_path = tmp_path / "resolve_text.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        text_range = resolve_text_range(
            ImpressTarget(
                kind="text",
                slide_index=3,
                shape_name="Copy Box",
                text="Target sentence",
                after="Anchor start",
                before="Anchor end",
            ),
            doc,
        )
        resolved_text = text_range.getString()

    assert resolved_text == "Target sentence"


def test_resolve_text_range_without_occurrence_raises_for_repeated_text(tmp_path):
    from impress import ImpressTarget
    from impress.exceptions import TargetAmbiguousError
    from impress.targets import resolve_text_range

    doc_path = tmp_path / "ambiguous_text.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        with pytest.raises(TargetAmbiguousError):
            resolve_text_range(
                ImpressTarget(
                    kind="text",
                    slide_index=3,
                    shape_name="Copy Box",
                    text="Shared phrase",
                ),
                doc,
            )


def test_notes_targets_resolve_the_notes_text_shape_for_requested_slide(tmp_path):
    from impress import ImpressTarget
    from impress.targets import resolve_text_range

    doc_path = tmp_path / "notes_target.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        notes_range = resolve_text_range(
            ImpressTarget(kind="notes", slide_index=3),
            doc,
        )
        notes_text = notes_range.getString()

    assert "review outputs" in notes_text.lower()


def test_list_targets_identify_structural_list_instead_of_bullet_text(tmp_path):
    from impress import ImpressTarget
    from impress.targets import resolve_list_target

    doc_path = tmp_path / "list_target.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        paragraphs = resolve_list_target(
            ImpressTarget(
                kind="list",
                slide_index=3,
                shape_name="Agenda Box",
                text="Review outputs",
            ),
            doc,
        )
        paragraph_texts = [paragraph.getString() for paragraph in paragraphs]

    assert paragraph_texts == ["Confirm scope", "Review outputs"]


def test_master_page_targets_resolve_by_name_and_fail_clearly_when_missing(tmp_path):
    from impress import ImpressTarget
    from impress.exceptions import TargetNoMatchError
    from impress.targets import resolve_master_page_target

    doc_path = tmp_path / "master_target.odp"
    _create_target_fixture(doc_path)

    with open_impress_doc(doc_path) as doc:
        master = resolve_master_page_target(
            ImpressTarget(kind="master_page", master_name="Default"),
            doc,
        )
        assert str(master.Name)

        with pytest.raises(TargetNoMatchError):
            resolve_master_page_target(
                ImpressTarget(kind="master_page", master_name="Missing Master"),
                doc,
            )
