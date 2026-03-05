"""Test basic package imports."""


def test_imports_writer_package():
    import writer

    assert hasattr(writer, "create_document")
    assert callable(writer.create_document)
    assert hasattr(writer, "snapshot_page")
    assert callable(writer.snapshot_page)
