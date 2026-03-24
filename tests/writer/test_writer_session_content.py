# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import (
    ARABIC_NUMBERING_TYPE,
    BULLET_NUMBERING_TYPE,
    assert_list_items,
    assert_text_formatting,
    create_test_image,
    get_graphic_names,
    get_graphic_size,
    get_list_paragraphs,
    get_table_names,
    get_text_properties,
)


def _create_document_with_text(doc_path, text):
    from writer import WriterSession
    from writer.core import create_document

    create_document(str(doc_path))
    with WriterSession(str(doc_path)) as session:
        session.insert_text(text)


def test_session_insert_text_appends_to_end(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "append_text.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Hello")
        assert session.read_text() == "Hello"


def test_session_read_text_without_target_returns_full_document(tmp_path):
    from writer import WriterSession

    doc_path = tmp_path / "read_full.odt"
    _create_document_with_text(doc_path, "Section 1\n\nSection 2\n\nTail")

    with WriterSession(str(doc_path)) as session:
        assert session.read_text() == "Section 1\n\nSection 2\n\nTail"


def test_session_read_text_with_bounded_target_returns_only_window(tmp_path):
    from writer import WriterTarget, WriterSession

    doc_path = tmp_path / "read_partial.odt"
    _create_document_with_text(
        doc_path,
        "Overview\n\nFinancial Summary\n\nSection 2\n\nRisks\n\nTail",
    )

    with WriterSession(str(doc_path)) as session:
        selected_text = session.read_text(
            target=WriterTarget(kind="text", after="Financial Summary", before="Risks")
        )

    assert "Section 2" in selected_text
    assert "Risks" not in selected_text
    assert "Overview" not in selected_text


def test_session_insert_text_at_insertion_target_after_anchor(tmp_path):
    from writer import WriterTarget, WriterSession

    doc_path = tmp_path / "insert_after.odt"
    _create_document_with_text(doc_path, "Introduction\n\nBody")

    with WriterSession(str(doc_path)) as session:
        session.insert_text(
            "Inserted",
            target=WriterTarget(kind="insertion", after="Introduction"),
        )
        text = session.read_text()

    assert text == "Introduction\nInserted\n\nBody"


def test_session_replace_text_uses_bounded_target_for_repeated_phrase(tmp_path):
    from writer import WriterTarget, WriterSession

    doc_path = tmp_path / "replace_text.odt"
    _create_document_with_text(
        doc_path,
        "Section A\n\nShared phrase\n\nSection B\n\nShared phrase\n\nTail",
    )

    with WriterSession(str(doc_path)) as session:
        session.replace_text(
            WriterTarget(
                kind="text", text="Shared phrase", after="Section B", before="Tail"
            ),
            "Updated phrase",
        )
        text = session.read_text()

    assert text.count("Shared phrase") == 1
    assert "Updated phrase" in text


def test_session_format_text_applies_character_formatting_to_targeted_phrase(tmp_path):
    from writer import TextFormatting, WriterTarget, WriterSession

    doc_path = tmp_path / "format_text.odt"
    _create_document_with_text(
        doc_path, "Financial Summary\n\nQuarterly revenue grew 18%."
    )

    with WriterSession(str(doc_path)) as session:
        session.format_text(
            WriterTarget(kind="text", text="Quarterly revenue grew 18%."),
            TextFormatting(bold=True, italic=True, underline=True),
        )

    assert_text_formatting(
        doc_path,
        "Quarterly revenue grew 18%.",
        char_weight=150.0,
        char_underline=1,
    )
    assert (
        "ITALIC"
        in get_text_properties(doc_path, "Quarterly revenue grew 18%.")["char_posture"]
    )


def test_session_format_text_applies_paragraph_alignment_to_targeted_paragraph(
    tmp_path,
):
    from writer import TextFormatting, WriterTarget, WriterSession

    doc_path = tmp_path / "format_alignment.odt"
    _create_document_with_text(doc_path, "Heading\n\nCentered paragraph\n\nTail")

    with WriterSession(str(doc_path)) as session:
        session.format_text(
            WriterTarget(kind="text", text="Centered paragraph"),
            TextFormatting(align="center"),
        )

    assert_text_formatting(doc_path, "Centered paragraph", align=3)


def test_session_format_text_start_and_end_alignment(tmp_path):
    """align='start' maps to ParagraphAdjust.START (5), align='end' to END (6)."""
    from writer import TextFormatting, WriterTarget, WriterSession

    doc_path = tmp_path / "start_end_align.odt"
    _create_document_with_text(doc_path, "Start paragraph\n\nEnd paragraph")

    with WriterSession(str(doc_path)) as session:
        session.format_text(
            WriterTarget(kind="text", text="Start paragraph"),
            TextFormatting(align="start"),
        )
        session.format_text(
            WriterTarget(kind="text", text="End paragraph"),
            TextFormatting(align="end"),
        )

    assert_text_formatting(doc_path, "Start paragraph", align=5)
    assert_text_formatting(doc_path, "End paragraph", align=6)


def test_session_format_text_invalid_alignment_raises(tmp_path):
    from writer import TextFormatting, WriterTarget, WriterSession
    from writer.exceptions import InvalidFormattingError

    doc_path = tmp_path / "bad_align.odt"
    _create_document_with_text(doc_path, "Some text")

    with WriterSession(str(doc_path)) as session:
        with pytest.raises(InvalidFormattingError, match="Unknown align value"):
            session.format_text(
                WriterTarget(kind="text", text="Some text"),
                TextFormatting(align="middle"),
            )


def test_session_format_text_empty_payload_raises_invalid_formatting_error(tmp_path):
    from writer import TextFormatting, WriterTarget, WriterSession
    from writer.exceptions import InvalidFormattingError

    doc_path = tmp_path / "empty_formatting.odt"
    _create_document_with_text(doc_path, "Only existing text")

    with WriterSession(str(doc_path)) as session:
        with pytest.raises(InvalidFormattingError):
            session.format_text(
                WriterTarget(kind="text", text="Only existing text"),
                TextFormatting(),
            )


def test_session_table_crud_continues_to_work_with_writer_target(tmp_path):
    from writer import WriterTarget, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "table_crud.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_table(2, 3, [["A", "B", "C"], ["1", "2", "3"]], "T1")
        session.update_table(
            WriterTarget(kind="table", name="T1"),
            [["X", "Y", "Z"], ["3", "4", "5"]],
        )

    assert get_table_names(doc_path) == ["T1"]

    with WriterSession(str(doc_path)) as session:
        session.delete_table(WriterTarget(kind="table", index=0))

    assert get_table_names(doc_path) == []


def test_session_image_crud_continues_to_work_with_writer_target(tmp_path):
    from writer import WriterTarget, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "image_crud.odt"
    first_image = create_test_image(tmp_path / "first.png", color="blue")
    second_image = create_test_image(tmp_path / "second.png", color="green")
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_image(str(first_image), width=2000, height=2000, name="Logo")
        session.update_image(
            WriterTarget(kind="image", name="Logo"),
            image_path=str(second_image),
            width=4000,
            height=3000,
        )

    width, height = get_graphic_size(doc_path, "Logo")
    assert abs(width - 4000) <= 1
    assert abs(height - 3000) <= 1
    assert get_graphic_names(doc_path) == ["Logo"]

    with WriterSession(str(doc_path)) as session:
        session.delete_image(WriterTarget(kind="image", index=0))

    assert get_graphic_names(doc_path) == []


def test_session_insert_list_creates_unordered_list_in_order(tmp_path):
    from writer import ListItem, WriterTarget, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "unordered_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Action Items")
        session.insert_list(
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review output", level=0),
            ],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Action Items"),
        )

    assert_list_items(
        doc_path,
        ["Confirm scope", "Review output"],
        expected_levels=[0, 0],
        expected_numbering_type=BULLET_NUMBERING_TYPE,
    )


def test_session_insert_list_creates_ordered_list_when_requested(tmp_path):
    from writer import ListItem, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "ordered_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_list(
            [ListItem(text="First", level=0), ListItem(text="Second", level=0)],
            ordered=True,
        )

    assert_list_items(
        doc_path,
        ["First", "Second"],
        expected_levels=[0, 0],
        expected_numbering_type=ARABIC_NUMBERING_TYPE,
    )


def test_session_insert_list_supports_nested_items(tmp_path):
    from writer import ListItem, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "nested_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_list(
            [
                ListItem(text="Parent", level=0),
                ListItem(text="Child", level=1),
                ListItem(text="Sibling", level=0),
            ],
            ordered=False,
        )

    assert_list_items(
        doc_path,
        ["Parent", "Child", "Sibling"],
        expected_levels=[0, 1, 0],
        expected_numbering_type=BULLET_NUMBERING_TYPE,
    )


def test_session_replace_list_replaces_existing_list_and_updates_ordering(tmp_path):
    from writer import ListItem, WriterTarget, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "replace_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Action Items\n\nRisks")
        session.insert_list(
            [ListItem(text="Old item", level=0)],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Action Items"),
        )
        session.replace_list(
            WriterTarget(
                kind="list", text="Old item", after="Action Items", before="Risks"
            ),
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Update packaging", level=1),
            ],
            ordered=True,
        )

    assert_list_items(
        doc_path,
        ["Confirm scope", "Update packaging"],
        expected_levels=[0, 1],
        expected_numbering_type=ARABIC_NUMBERING_TYPE,
    )


def test_session_delete_list_removes_targeted_list(tmp_path):
    from writer import ListItem, WriterTarget, WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "delete_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_list([ListItem(text="Remove me", level=0)], ordered=False)
        session.delete_list(WriterTarget(kind="list", text="Remove me"))

    assert get_list_paragraphs(doc_path) == []


def test_session_invalid_list_payloads_raise_invalid_list_error(tmp_path):
    from writer import ListItem, WriterSession
    from writer.core import create_document
    from writer.exceptions import InvalidListError

    doc_path = tmp_path / "invalid_list.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        with pytest.raises(InvalidListError):
            session.insert_list([], ordered=False)
        with pytest.raises(InvalidListError):
            session.insert_list([ListItem(text="Bad", level=-1)], ordered=False)
