"""Test basic package imports."""


def test_imports_writer_package():
    import libreoffice_skills.writer

    assert hasattr(libreoffice_skills.writer, "create_document")
    assert callable(libreoffice_skills.writer.create_document)
    assert hasattr(libreoffice_skills.writer, "snapshot_page")
    assert callable(libreoffice_skills.writer.snapshot_page)
