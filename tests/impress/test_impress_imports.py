"""Test basic package imports."""


def test_imports_impress_package():
    import impress

    assert hasattr(impress, "create_presentation")
    assert callable(impress.create_presentation)
