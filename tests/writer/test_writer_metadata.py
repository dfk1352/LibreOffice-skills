import pytest


def test_set_metadata_rejects_empty_key(tmp_path):
    from writer.core import create_document
    from writer.exceptions import InvalidMetadataError
    from writer.session import WriterSession

    doc_path = tmp_path / "sample.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    try:
        with pytest.raises(InvalidMetadataError):
            session.set_metadata({"": "value"})
    finally:
        session.close()


def test_set_and_get_metadata(tmp_path):
    from writer.core import create_document
    from writer.session import WriterSession

    doc_path = tmp_path / "test_metadata.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    try:
        metadata = {
            "title": "Test Document",
            "author": "Test Author",
            "subject": "Metadata Subject",
            "keywords": "alpha, beta",
            "description": "A description",
        }
        session.set_metadata(metadata)

        retrieved = session.get_metadata()

        assert retrieved["title"] == "Test Document"
        assert retrieved["author"] == "Test Author"
        assert retrieved["subject"] == "Metadata Subject"
        assert retrieved["keywords"] == "alpha, beta"
        assert retrieved["description"] == "A description"
    finally:
        session.close()


def test_metadata_persists_after_save(tmp_path):
    """Verify metadata survives a close/reopen cycle."""
    from writer.core import create_document
    from writer.session import WriterSession

    doc_path = tmp_path / "persist.odt"
    create_document(str(doc_path))

    session = WriterSession(str(doc_path))
    session.set_metadata({"title": "Persistent Title"})
    session.close(save=True)

    session2 = WriterSession(str(doc_path))
    try:
        assert session2.get_metadata()["title"] == "Persistent Title"
    finally:
        session2.close()
