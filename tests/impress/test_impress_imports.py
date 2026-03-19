"""Tests for Impress package exports."""

# pyright: reportAttributeAccessIssue=false, reportMissingImports=false


def test_imports_impress_session_first_package_exports():
    import impress

    assert hasattr(impress, "ImpressSession")
    assert hasattr(impress, "ImpressTarget")
    assert hasattr(impress, "ListItem")
    assert hasattr(impress, "ShapePlacement")
    assert hasattr(impress, "TextFormatting")
    assert hasattr(impress, "create_presentation")
    assert hasattr(impress, "export_presentation")
    assert hasattr(impress, "open_impress_session")
    assert hasattr(impress, "patch")
    assert hasattr(impress, "snapshot_slide")
