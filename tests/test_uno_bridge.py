"""Test UNO bridge functionality."""

import pytest


def test_bridge_validates_missing_lo_path():
    from libreoffice_skills.uno_bridge import validate_lo_path
    from libreoffice_skills.exceptions import UnoBridgeError

    with pytest.raises(UnoBridgeError):
        validate_lo_path("/missing/libreoffice")


def test_find_libreoffice():
    from libreoffice_skills.uno_bridge import find_libreoffice

    path = find_libreoffice()
    assert path is not None, "LibreOffice should be found on this system"
    assert "soffice" in path or "libreoffice" in path.lower()


def test_uno_context_connection(tmp_path):
    """Test that we can establish a UNO connection to LibreOffice."""
    from libreoffice_skills.uno_bridge import uno_context

    with uno_context() as desktop:
        assert desktop is not None
        # Verify by creating and closing a blank Writer document
        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        try:
            assert doc.supportsService("com.sun.star.text.TextDocument")
        finally:
            doc.close(True)
