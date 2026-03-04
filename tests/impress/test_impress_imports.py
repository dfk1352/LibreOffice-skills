"""Test basic package imports."""


def test_imports_impress_package():
    import libreoffice_skills.impress

    assert hasattr(libreoffice_skills.impress, "create_presentation")
    assert callable(libreoffice_skills.impress.create_presentation)
