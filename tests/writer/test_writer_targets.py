# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import create_test_image, get_list_paragraphs, open_uno_doc


def _create_target_fixture_document(doc_path, image_path):
    from writer import ListItem, WriterTarget, WriterSession
    from writer.core import create_document

    create_document(str(doc_path))
    with WriterSession(str(doc_path)) as session:
        session.insert_text(
            "Overview\n\n"
            "Financial Summary\n\n"
            "Quarterly revenue grew 18%.\n\n"
            "Shared phrase\n\n"
            "Shared phrase\n\n"
            "Action Items\n\n"
            "Risks\n\n"
            "Appendix"
        )
        session.insert_list(
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review output", level=0),
            ],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Action Items"),
        )
        session.insert_table(1, 1, [["budget"]], name="Budget Table")
        session.insert_image(str(image_path), name="Logo")


def test_parse_target_builds_bounded_text_target():
    from writer.targets import parse_target

    target = parse_target(
        {
            "kind": "text",
            "text": "quarterly revenue",
            "after": "Financial Summary",
            "before": "Risks",
            "occurrence": 0,
        }
    )

    assert target.kind == "text"
    assert target.text == "quarterly revenue"
    assert target.after == "Financial Summary"
    assert target.before == "Risks"
    assert target.occurrence == 0


def test_parse_target_rejects_name_and_index_together():
    from writer.exceptions import InvalidTargetError
    from writer.targets import parse_target

    with pytest.raises(InvalidTargetError):
        parse_target({"kind": "table", "name": "Budget", "index": 0})


def test_parse_target_rejects_negative_occurrence():
    from writer.exceptions import InvalidTargetError
    from writer.targets import parse_target

    with pytest.raises(InvalidTargetError):
        parse_target({"kind": "text", "text": "Shared phrase", "occurrence": -1})


def test_resolve_text_range_respects_after_and_before_bounds(tmp_path):
    from writer import WriterTarget
    from writer.targets import resolve_text_range

    doc_path = tmp_path / "bounded_text.odt"
    image_path = create_test_image(tmp_path / "bounded_text.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        text_range = resolve_text_range(
            WriterTarget(kind="text", after="Financial Summary", before="Risks"),
            doc,
        )
        bounded_text = text_range.getString()
    assert "Quarterly revenue grew 18%." in bounded_text
    assert "Action Items" in bounded_text
    assert "Risks" not in bounded_text
    assert "Overview" not in bounded_text


def test_resolve_text_range_uses_zero_based_occurrence_with_repeated_text(tmp_path):
    from writer import WriterTarget
    from writer.targets import resolve_text_range

    doc_path = tmp_path / "occurrence.odt"
    image_path = create_test_image(tmp_path / "occurrence.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        text_range = resolve_text_range(
            WriterTarget(kind="text", text="Shared phrase", occurrence=1),
            doc,
        )
        resolved_text = text_range.getString()

    assert resolved_text == "Shared phrase"


def test_resolve_text_range_without_occurrence_raises_for_repeated_text(tmp_path):
    from writer import WriterTarget
    from writer.exceptions import TargetAmbiguousError
    from writer.targets import resolve_text_range

    doc_path = tmp_path / "ambiguous_text.odt"
    image_path = create_test_image(tmp_path / "ambiguous_text.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        with pytest.raises(TargetAmbiguousError):
            resolve_text_range(WriterTarget(kind="text", text="Shared phrase"), doc)


def test_resolve_text_range_rejects_after_anchor_after_before_anchor(tmp_path):
    from writer import WriterTarget
    from writer.exceptions import InvalidTargetError
    from writer.targets import resolve_text_range

    doc_path = tmp_path / "bad_window.odt"
    image_path = create_test_image(tmp_path / "bad_window.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        with pytest.raises(InvalidTargetError):
            resolve_text_range(
                WriterTarget(kind="text", after="Risks", before="Financial Summary"),
                doc,
            )


def test_resolve_list_target_identifies_list_by_item_text_within_bounds(tmp_path):
    from writer import WriterTarget
    from writer.targets import resolve_list_target

    doc_path = tmp_path / "list_target.odt"
    image_path = create_test_image(tmp_path / "list_target.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        resolved_list = resolve_list_target(
            WriterTarget(
                kind="list",
                text="Review output",
                after="Action Items",
                before="Risks",
            ),
            doc,
        )
        assert [paragraph.getString() for paragraph in resolved_list] == [
            "Confirm scope",
            "Review output",
        ]


def test_resolve_table_target_resolves_named_table(tmp_path):
    from writer import WriterTarget
    from writer.targets import resolve_table_target

    doc_path = tmp_path / "table_target.odt"
    image_path = create_test_image(tmp_path / "table_target.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        table = resolve_table_target(
            WriterTarget(kind="table", name="Budget Table"), doc
        )
        table_name = table.Name

    assert table_name == "Budget_Table"


def test_resolve_image_target_resolves_image_by_index(tmp_path):
    from writer import WriterTarget
    from writer.targets import resolve_image_target

    doc_path = tmp_path / "image_target.odt"
    image_path = create_test_image(tmp_path / "image_target.png")
    _create_target_fixture_document(doc_path, image_path)

    with open_uno_doc(doc_path) as doc:
        graphic = resolve_image_target(WriterTarget(kind="image", index=0), doc)
        graphic_name = graphic.Name

    assert graphic_name == "Logo"
    assert [item["text"] for item in get_list_paragraphs(doc_path)] == [
        "Confirm scope",
        "Review output",
    ]
