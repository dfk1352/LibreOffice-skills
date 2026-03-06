"""Test Writer text operations."""

import pytest


def test_insert_text_requires_existing_document(tmp_path):
    from writer.text import insert_text
    from writer.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        insert_text(str(tmp_path / "missing.odt"), "Hello")


def test_insert_text_adds_content(tmp_path):
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import insert_text

    doc_path = tmp_path / "test_insert.odt"
    create_document(str(doc_path))

    # Insert text
    insert_text(str(doc_path), "Hello, World!")

    # Verify text was inserted
    content = read_document_text(str(doc_path))
    assert content == "Hello, World!"


def test_append_text_adds_to_end(tmp_path):
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import insert_text

    doc_path = tmp_path / "test_append.odt"
    create_document(str(doc_path))

    # Insert initial text
    insert_text(str(doc_path), "First line")

    # Append more text
    insert_text(str(doc_path), "\nSecond line", position=None)

    # Verify order is preserved
    content = read_document_text(str(doc_path))
    assert content == "First line\nSecond line"


def test_replace_text_modifies_content(tmp_path):
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import insert_text, replace_text

    doc_path = tmp_path / "test_replace.odt"
    create_document(str(doc_path))

    # Insert initial text
    insert_text(str(doc_path), "Hello, World!")

    # Replace text
    replace_text(str(doc_path), "World", "LibreOffice")

    # Verify replacement
    content = read_document_text(str(doc_path))
    assert "Hello, LibreOffice!" in content
    assert "World" not in content


def test_insert_text_at_index(tmp_path):
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import insert_text

    doc_path = tmp_path / "test_insert_index.odt"
    create_document(str(doc_path))

    insert_text(str(doc_path), "World", position=0)
    insert_text(str(doc_path), "Hello ", position=0)
    insert_text(str(doc_path), "brave ", position=6)

    content = read_document_text(str(doc_path))
    assert content == "Hello brave World"
