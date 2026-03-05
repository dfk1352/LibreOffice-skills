"""Test Writer formatting operations."""

import pytest


def test_apply_formatting_rejects_unknown_keys(tmp_path):
    from writer.core import create_document
    from writer.formatting import apply_formatting
    from writer.exceptions import InvalidFormattingError

    doc_path = tmp_path / "sample.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidFormattingError):
        apply_formatting(str(doc_path), {"unknown": True})


def test_apply_character_formatting(tmp_path):
    from writer.core import create_document
    from writer.text import insert_text
    from writer.formatting import apply_formatting
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_char_format.odt"
    create_document(str(doc_path))
    insert_text(str(doc_path), "Test text for formatting")

    # Apply character formatting
    apply_formatting(str(doc_path), {"bold": True, "italic": True, "underline": True})

    # Verify formatting by reading cursor properties
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            cursor = doc.Text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.CharWeight == 150
            assert "ITALIC" in str(cursor.CharPosture)
            assert cursor.CharUnderline == 1
        finally:
            doc.close(True)


def test_apply_paragraph_formatting(tmp_path):
    from writer.core import create_document
    from writer.text import insert_text
    from writer.formatting import apply_formatting
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_para_format.odt"
    create_document(str(doc_path))
    insert_text(str(doc_path), "Test paragraph")

    # Apply paragraph formatting
    apply_formatting(str(doc_path), {"align": "CENTER"})

    # Verify paragraph alignment
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            cursor = doc.Text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.ParaAdjust == 2
        finally:
            doc.close(True)


def test_apply_paragraph_formatting_invalid_align(tmp_path):
    from writer.core import create_document
    from writer.formatting import apply_formatting
    from writer.exceptions import InvalidFormattingError

    doc_path = tmp_path / "test_para_bad_align.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidFormattingError):
        apply_formatting(str(doc_path), {"align": "sideways"})


def test_apply_font_properties(tmp_path):
    from writer.core import create_document
    from writer.text import insert_text
    from writer.formatting import apply_formatting
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_font.odt"
    create_document(str(doc_path))
    insert_text(str(doc_path), "Font test")

    # Apply font properties
    apply_formatting(str(doc_path), {"font_name": "Arial", "font_size": 14})

    # Verify font properties
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            cursor = doc.Text.createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.CharFontName == "Arial"
            assert cursor.CharHeight == 14
        finally:
            doc.close(True)
