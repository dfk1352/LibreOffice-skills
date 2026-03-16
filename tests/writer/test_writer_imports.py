"""Test basic package imports."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false


def test_imports_writer_package():
    import writer

    assert hasattr(writer, "create_document")
    assert callable(writer.create_document)
    assert hasattr(writer, "export_document")
    assert callable(writer.export_document)
    assert hasattr(writer, "snapshot_page")
    assert callable(writer.snapshot_page)
    assert hasattr(writer, "open_writer_session")
    assert callable(writer.open_writer_session)
    assert hasattr(writer, "patch")
    assert callable(writer.patch)
    assert hasattr(writer, "WriterSession")
    assert hasattr(writer, "WriterTarget")
    assert hasattr(writer, "TextFormatting")
    assert hasattr(writer, "ListItem")
    assert hasattr(writer, "PatchApplyResult")
    assert hasattr(writer, "PatchOperationResult")
