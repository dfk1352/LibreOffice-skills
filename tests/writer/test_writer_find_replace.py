"""Tests for Writer find & replace operations."""


def test_find_replace_in_writer(tmp_path):
    from libreoffice_skills.writer.core import (
        create_document,
        read_document_text,
    )
    from libreoffice_skills.writer.find_replace import find_replace
    from libreoffice_skills.writer.text import insert_text

    path = tmp_path / "find_replace.odt"
    create_document(str(path))
    insert_text(str(path), "Hello World")

    count = find_replace(str(path), "Hello", "Goodbye")

    assert count >= 1

    text = read_document_text(str(path))
    assert "Goodbye" in text
