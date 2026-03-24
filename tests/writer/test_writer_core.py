# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

import zipfile

import pytest


def test_create_document_requires_output_path(tmp_path):
    from writer.core import create_document

    output_path = tmp_path / "sample.odt"
    create_document(str(output_path))

    assert output_path.exists()

    assert zipfile.is_zipfile(output_path)

    with zipfile.ZipFile(output_path) as zf:
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
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "session_export.odt"
    output_path = tmp_path / "session_export.docx"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Unsaved export text")
        session.export(str(output_path), "docx")
        assert session.read_text() == "Unsaved export text"

    assert output_path.exists()
    assert output_path.stat().st_size > 0
    assert zipfile.is_zipfile(output_path)


def test_session_reset_discards_unsaved_changes(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "reset_test.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        session.insert_text("Temporary content")
        assert session.read_text() == "Temporary content"
        session.reset()
        assert session.read_text() == ""


def test_session_export_unknown_format_raises(tmp_path):
    from writer import WriterSession
    from writer.core import create_document
    from writer.exceptions import WriterSkillError

    doc_path = tmp_path / "bad_export.odt"
    create_document(str(doc_path))

    with WriterSession(str(doc_path)) as session:
        with pytest.raises(WriterSkillError, match="Unsupported export format"):
            session.export(str(tmp_path / "out.bmp"), "bmp")


# --- Markdown import/export tests ---


def test_create_document_from_markdown(tmp_path):
    from writer.core import create_document

    md_path = tmp_path / "input.md"
    md_path.write_text("# Hello\n\nA paragraph.\n")
    odt_path = tmp_path / "output.odt"

    create_document(str(odt_path), source=str(md_path))

    assert odt_path.exists()
    assert zipfile.is_zipfile(odt_path)
    with zipfile.ZipFile(odt_path) as zf:
        assert "content.xml" in zf.namelist()


def test_create_document_from_markdown_preserves_text(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    md_path = tmp_path / "input.md"
    md_path.write_text("# Heading\n\nSome body text.\n")
    odt_path = tmp_path / "imported.odt"

    create_document(str(odt_path), source=str(md_path))

    with WriterSession(str(odt_path)) as session:
        text = session.read_text()
        assert "Heading" in text
        assert "Some body text." in text


def test_create_document_from_markdown_source_not_found(tmp_path):
    from writer.core import create_document
    from writer.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        create_document(str(tmp_path / "out.odt"), source=str(tmp_path / "missing.md"))


def test_create_document_from_markdown_creates_parent_dirs(tmp_path):
    from writer.core import create_document

    md_path = tmp_path / "input.md"
    md_path.write_text("# Test\n")
    odt_path = tmp_path / "sub" / "deep" / "imported.odt"

    create_document(str(odt_path), source=str(md_path))

    assert odt_path.exists()


def test_create_document_unsupported_source_format_raises(tmp_path):
    from writer.core import create_document
    from writer.exceptions import WriterSkillError

    txt_path = tmp_path / "input.txt"
    txt_path.write_text("plain text")

    with pytest.raises(WriterSkillError, match="Unsupported import format"):
        create_document(str(tmp_path / "out.odt"), source=str(txt_path))


def test_export_document_writes_md(tmp_path):
    from writer import WriterSession
    from writer.core import create_document, export_document

    doc_path = tmp_path / "source.odt"
    create_document(str(doc_path))
    with WriterSession(str(doc_path)) as session:
        session.insert_text("Markdown export content")

    md_path = tmp_path / "exported.md"
    export_document(str(doc_path), str(md_path), "md")

    assert md_path.exists()
    content = md_path.read_text()
    assert "Markdown export content" in content


def test_session_export_md(tmp_path):
    from writer import WriterSession
    from writer.core import create_document

    doc_path = tmp_path / "session_md.odt"
    create_document(str(doc_path))

    md_path = tmp_path / "session_export.md"
    with WriterSession(str(doc_path)) as session:
        session.insert_text("Unsaved markdown text")
        session.export(str(md_path), "md")

    assert md_path.exists()
    content = md_path.read_text()
    assert "Unsaved markdown text" in content
