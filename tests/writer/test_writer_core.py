"""Test Writer core document operations."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

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


def test_export_document_writes_pdf(tmp_path):
    from writer.core import create_document, export_document

    doc_path = tmp_path / "writer.odt"
    output_path = tmp_path / "writer.pdf"
    create_document(str(doc_path))

    export_document(str(doc_path), str(output_path), "pdf")

    assert output_path.exists()
    assert output_path.stat().st_size > 0
    with open(output_path, "rb") as handle:
        assert handle.read(5) == b"%PDF-"


def test_session_export_uses_current_unsaved_document_state(tmp_path):
    from writer import open_writer_session
    from writer.core import create_document

    doc_path = tmp_path / "session_export.odt"
    output_path = tmp_path / "session_export.docx"
    create_document(str(doc_path))

    with open_writer_session(str(doc_path)) as session:
        session.insert_text("Unsaved export text")
        session.export(str(output_path), "docx")
        assert session.read_text() == "Unsaved export text"

    assert output_path.exists()
    assert output_path.stat().st_size > 0
    assert zipfile.is_zipfile(output_path)
