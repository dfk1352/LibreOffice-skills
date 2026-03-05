"""Test Writer core document operations."""

import zipfile


def test_create_document_requires_output_path(tmp_path):
    from writer.core import create_document

    output_path = tmp_path / "sample.odt"
    create_document(str(output_path))

    # Verify file exists
    assert output_path.exists()

    # Verify it's a valid ODT (ZIP file with specific structure)
    assert zipfile.is_zipfile(output_path)

    with zipfile.ZipFile(output_path) as zf:
        # ODT files must contain these files
        assert "content.xml" in zf.namelist()
        assert "META-INF/manifest.xml" in zf.namelist()


def test_read_document_text(tmp_path):
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import insert_text

    output_path = tmp_path / "test_read.odt"
    create_document(str(output_path))

    # Read empty document
    text = read_document_text(str(output_path))
    assert text == ""  # New document should be empty

    # Insert content and read again
    insert_text(str(output_path), "Read me")
    text = read_document_text(str(output_path))
    assert text == "Read me"


def test_read_nonexistent_document(tmp_path):
    from writer.core import read_document_text
    from writer.exceptions import WriterSkillError
    import pytest

    with pytest.raises(WriterSkillError):
        read_document_text(str(tmp_path / "missing.odt"))


def test_export_writer_pdf(tmp_path):
    from writer.core import create_document, export_document

    path = tmp_path / "writer.odt"
    output = tmp_path / "writer.pdf"
    create_document(str(path))
    export_document(str(path), str(output), "pdf")
    assert output.exists()
    assert output.stat().st_size > 0
    # Verify it is a valid PDF by checking magic bytes
    with open(output, "rb") as f:
        assert f.read(5) == b"%PDF-"
