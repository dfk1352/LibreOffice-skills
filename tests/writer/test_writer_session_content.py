"""Tests for Writer content CRUD through sessions."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest

from tests.writer._helpers import (
    create_test_image,
    get_graphic_names,
    get_graphic_size,
    get_table_cell_value,
    get_table_dimensions,
    get_table_names,
)


def _create_document_with_text(doc_path, text):
    from writer import open_writer_session
    from writer.core import create_document

    create_document(str(doc_path))
    with open_writer_session(str(doc_path)) as session:
        session.insert_text(text)


def test_session_insert_text_appends_to_end(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "append_text.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_text("Hello")
        assert session.read_text() == "Hello"


def test_session_read_text_without_selector_returns_full_document(tmp_path):
    from writer import open_writer_session

    doc_path = tmp_path / "read_full.odt"
    _create_document_with_text(doc_path, "Section 1\n\nSection 2\n\nTail")

    with open_writer_session(str(doc_path)) as session:
        assert session.read_text() == "Section 1\n\nSection 2\n\nTail"


def test_session_read_text_with_selector_returns_only_matched_span(tmp_path):
    from writer import open_writer_session

    doc_path = tmp_path / "read_partial.odt"
    _create_document_with_text(doc_path, "Section 1\n\nSection 2\n\nTail")

    with open_writer_session(str(doc_path)) as session:
        selected_text = session.read_text(selector='contains:"Section 2"')

    assert "Section 2" in selected_text
    assert selected_text != "Section 1\n\nSection 2\n\nTail"


def test_session_read_text_with_missing_selector_raises(tmp_path):
    from writer import open_writer_session
    from writer.exceptions import SelectorNoMatchError

    doc_path = tmp_path / "read_missing.odt"
    _create_document_with_text(doc_path, "Section 1\n\nSection 2")

    with open_writer_session(str(doc_path)) as session:
        with pytest.raises(SelectorNoMatchError):
            session.read_text(selector='contains:"Missing"')


def test_session_insert_text_after_anchor(tmp_path):
    from writer import open_writer_session

    doc_path = tmp_path / "insert_after.odt"
    _create_document_with_text(doc_path, "Introduction\n\nBody")

    with open_writer_session(str(doc_path)) as session:
        session.insert_text("\n\nInserted", selector='after:"Introduction"')
        text = session.read_text()

    assert text.index("Introduction") < text.index("Inserted") < text.index("Body")


def test_session_replace_text_replaces_matched_span(tmp_path):
    from writer import open_writer_session

    doc_path = tmp_path / "replace_text.odt"
    _create_document_with_text(doc_path, "Keep this. Replace old text. Keep tail.")

    with open_writer_session(str(doc_path)) as session:
        session.replace_text('contains:"old text"', "new text")
        text = session.read_text()

    assert "new text" in text
    assert "old text" not in text


def test_session_delete_text_removes_matched_span(tmp_path):
    from writer import open_writer_session

    doc_path = tmp_path / "delete_text.odt"
    _create_document_with_text(doc_path, "Keep this. remove me. Keep tail.")

    with open_writer_session(str(doc_path)) as session:
        session.delete_text('contains:"remove me"')
        text = session.read_text()

    assert "remove me" not in text
    assert "Keep this." in text


def test_session_replace_text_missing_selector_raises(tmp_path):
    from writer import open_writer_session
    from writer.exceptions import SelectorNoMatchError

    doc_path = tmp_path / "replace_missing.odt"
    _create_document_with_text(doc_path, "Only existing text")

    with open_writer_session(str(doc_path)) as session:
        with pytest.raises(SelectorNoMatchError):
            session.replace_text('contains:"missing"', "new")


def test_session_insert_table_creates_named_table(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "insert_table.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_table(2, 3, [["A", "B", "C"], ["1", "2", "3"]], "T1")

    assert "T1" in get_table_names(doc_path)
    assert get_table_dimensions(doc_path, "T1") == (2, 3)
    assert get_table_cell_value(doc_path, "T1", "B2") == "2"


def test_session_update_table_updates_cell_contents(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "update_table.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_table(2, 2, [["A", "B"], ["1", "2"]], "T1")
        session.update_table('name:"T1"', [["X", "Y"], ["3", "4"]])

    assert get_table_cell_value(doc_path, "T1", "A1") == "X"
    assert get_table_cell_value(doc_path, "T1", "B2") == "4"


def test_session_delete_table_removes_named_table(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "delete_table.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_table(1, 1, [["gone"]], "T1")
        session.delete_table('name:"T1"')

    assert "T1" not in get_table_names(doc_path)


def test_session_delete_table_unknown_name_raises(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import SelectorNoMatchError

    doc_path = tmp_path / "delete_missing_table.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        with pytest.raises(SelectorNoMatchError):
            session.delete_table('name:"Missing Table"')


def test_session_insert_image_creates_named_graphic(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "insert_image.odt"
    image_path = create_test_image(tmp_path / "insert_image.png", color="blue")
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_image(str(image_path), name="Logo")

    assert "Logo" in get_graphic_names(doc_path)


def test_session_update_image_replaces_properties_of_named_graphic(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "update_image.odt"
    first_image = create_test_image(tmp_path / "first.png", color="blue")
    second_image = create_test_image(tmp_path / "second.png", color="green")
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_image(str(first_image), width=2000, height=2000, name="Logo")
        session.update_image(
            'name:"Logo"',
            image_path=str(second_image),
            width=4000,
            height=3000,
        )

    width, height = get_graphic_size(doc_path, "Logo")
    assert abs(width - 4000) <= 1
    assert abs(height - 3000) <= 1


def test_session_delete_image_removes_named_graphic(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "delete_image.odt"
    image_path = create_test_image(tmp_path / "delete.png", color="black")
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_image(str(image_path), name="Logo")
        session.delete_image('name:"Logo"')

    assert get_graphic_names(doc_path) == []


def test_session_insert_image_missing_file_raises(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document
    from writer.exceptions import ImageNotFoundError

    doc_path = tmp_path / "missing_image.odt"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        with pytest.raises(ImageNotFoundError):
            session.insert_image(str(tmp_path / "missing.png"), name="Logo")
