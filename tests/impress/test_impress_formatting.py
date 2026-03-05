"""Tests for Impress formatting operations."""


def test_format_shape_text_bold(tmp_path):
    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.formatting import format_shape_text
    from uno_bridge import uno_context

    path = tmp_path / "format_bold.odp"
    create_presentation(str(path))

    shape_idx = add_text_box(str(path), 0, "Bold text", 1.0, 1.0, 10.0, 3.0)
    format_shape_text(
        str(path),
        0,
        shape_idx,
        bold=True,
        font_size=18,
    )

    # Verify via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            path.resolve().as_uri(),
            "_blank",
            0,
            (),
        )
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(shape_idx)
            cursor = shape.getText().createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.CharWeight == 150
            assert cursor.CharHeight == 18
        finally:
            doc.close(True)


def test_set_slide_background_color(tmp_path):
    from impress.core import create_presentation
    from impress.formatting import set_slide_background
    from uno_bridge import uno_context

    path = tmp_path / "background.odp"
    create_presentation(str(path))

    set_slide_background(str(path), 0, "red")

    # Verify via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            path.resolve().as_uri(),
            "_blank",
            0,
            (),
        )
        try:
            slide = doc.DrawPages.getByIndex(0)
            bg = slide.Background
            assert bg.FillColor == 0xFF0000
        finally:
            doc.close(True)


def test_format_shape_text_alignment(tmp_path):
    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.formatting import format_shape_text
    from uno_bridge import uno_context

    path = tmp_path / "format_align.odp"
    create_presentation(str(path))

    shape_idx = add_text_box(str(path), 0, "Center text", 1.0, 1.0, 10.0, 3.0)
    format_shape_text(str(path), 0, shape_idx, alignment="center")

    # Verify alignment via UNO roundtrip (CENTER = 3)
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(shape_idx)
            cursor = shape.getText().createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.ParaAdjust == 3  # CENTER
        finally:
            doc.close(True)


def test_format_shape_text_alignment_invalid_raises(tmp_path):
    import pytest

    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.formatting import format_shape_text

    path = tmp_path / "format_align_invalid.odp"
    create_presentation(str(path))

    shape_idx = add_text_box(str(path), 0, "Bad align", 1.0, 1.0, 10.0, 3.0)

    with pytest.raises(ValueError):
        format_shape_text(str(path), 0, shape_idx, alignment="sideways")


def test_format_shape_text_alignment_case_insensitive(tmp_path):
    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.formatting import format_shape_text
    from uno_bridge import uno_context

    path = tmp_path / "format_align_case.odp"
    create_presentation(str(path))

    shape_idx = add_text_box(str(path), 0, "Center text", 1.0, 1.0, 10.0, 3.0)
    format_shape_text(str(path), 0, shape_idx, alignment="CENTER")

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(shape_idx)
            cursor = shape.getText().createTextCursor()
            cursor.gotoStart(False)
            cursor.gotoEnd(True)
            assert cursor.ParaAdjust == 3  # CENTER
        finally:
            doc.close(True)
