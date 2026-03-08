"""Tests for Impress content placement operations."""

# pyright: reportMissingImports=false

import pytest


def test_set_title_missing_doc_raises(tmp_path):
    import impress.content as content
    from impress.exceptions import DocumentNotFoundError

    missing = str(tmp_path / "no_such.odp")
    with pytest.raises(DocumentNotFoundError):
        content.set_title(missing, 0, "Title")


def test_set_body_missing_doc_raises(tmp_path):
    import impress.content as content
    from impress.exceptions import DocumentNotFoundError

    missing = str(tmp_path / "no_such.odp")
    with pytest.raises(DocumentNotFoundError):
        content.set_body(missing, 0, "Body")


def test_add_text_box_missing_doc_raises(tmp_path):
    import impress.content as content
    from impress.exceptions import DocumentNotFoundError

    missing = str(tmp_path / "no_such.odp")
    with pytest.raises(DocumentNotFoundError):
        content.add_text_box(missing, 0, "Hello", 1.0, 1.0, 5.0, 2.0)


def test_add_image_missing_doc_raises(tmp_path):
    import impress.content as content
    from impress.exceptions import DocumentNotFoundError

    image_path = tmp_path / "image.png"
    image_path.write_bytes(b"png")

    missing = str(tmp_path / "no_such.odp")
    with pytest.raises(DocumentNotFoundError):
        content.add_image(missing, 0, str(image_path), 1.0, 1.0, 5.0, 5.0)


def test_add_shape_missing_doc_raises(tmp_path):
    import impress.content as content
    from impress.exceptions import DocumentNotFoundError

    missing = str(tmp_path / "no_such.odp")
    with pytest.raises(DocumentNotFoundError):
        content.add_shape(missing, 0, "rectangle", 1.0, 1.0, 5.0, 3.0)


def test_set_title_on_title_slide(tmp_path):
    from impress.content import set_title
    from impress.core import create_presentation
    from impress.slides import add_slide, get_slide_inventory

    path = tmp_path / "title.odp"
    create_presentation(str(path))
    add_slide(str(path), layout="TITLE_SLIDE")

    set_title(str(path), 1, "My Title")

    inventory = get_slide_inventory(str(path), 1)
    texts = [s["text"] for s in inventory["shapes"]]
    assert "My Title" in texts


def test_set_body_on_content_slide(tmp_path):
    from impress.content import set_body
    from impress.core import create_presentation
    from impress.slides import add_slide, get_slide_inventory

    path = tmp_path / "body.odp"
    create_presentation(str(path))
    add_slide(str(path), layout="TITLE_AND_CONTENT")

    set_body(str(path), 1, "Body text here")

    inventory = get_slide_inventory(str(path), 1)
    texts = [s["text"] for s in inventory["shapes"]]
    assert "Body text here" in texts


def test_add_text_box_returns_shape_index(tmp_path):
    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.slides import get_slide_inventory

    path = tmp_path / "textbox.odp"
    create_presentation(str(path))

    result = add_text_box(str(path), 0, "Hello", 2.0, 3.0, 10.0, 5.0)

    assert isinstance(result, int)
    # Verify the text box contains our text via inventory
    inventory = get_slide_inventory(str(path), 0)
    texts = [s["text"] for s in inventory["shapes"]]
    assert "Hello" in texts


def test_add_text_box_roundtrip(tmp_path):
    from impress.content import add_text_box
    from impress.core import create_presentation
    from impress.slides import get_slide_inventory

    path = tmp_path / "textbox_rt.odp"
    create_presentation(str(path))

    add_text_box(str(path), 0, "Roundtrip text", 1.0, 1.0, 8.0, 3.0)

    inventory = get_slide_inventory(str(path), 0)
    texts = [s["text"] for s in inventory["shapes"]]
    assert "Roundtrip text" in texts


def test_add_image_returns_shape_index(tmp_path):
    from PIL import Image

    from impress.content import add_image
    from impress.core import create_presentation
    from uno_bridge import uno_context

    path = tmp_path / "image.odp"
    create_presentation(str(path))

    img_path = tmp_path / "test.png"
    img = Image.new("RGB", (100, 100), color="red")
    img.save(img_path)

    result = add_image(str(path), 0, str(img_path), 2.0, 2.0, 5.0, 5.0)

    assert isinstance(result, int)

    # Verify image shape exists on the slide via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            # Graphic shapes support this service
            assert shape.supportsService(
                "com.sun.star.drawing.GraphicObjectShape"
            ) or shape.supportsService("com.sun.star.presentation.MediaShape")
        finally:
            doc.close(True)


def test_add_image_missing_file_raises(tmp_path):
    from impress.content import add_image
    from impress.core import create_presentation
    from impress.exceptions import MediaNotFoundError

    path = tmp_path / "image_missing.odp"
    create_presentation(str(path))

    with pytest.raises(MediaNotFoundError):
        add_image(str(path), 0, str(tmp_path / "missing.png"), 1.0, 1.0, 5.0, 5.0)


def test_add_shape_returns_index(tmp_path):
    from impress.content import add_shape
    from impress.core import create_presentation
    from uno_bridge import uno_context

    path = tmp_path / "shape.odp"
    create_presentation(str(path))

    result = add_shape(
        str(path),
        0,
        "rectangle",
        1.0,
        1.0,
        5.0,
        3.0,
        fill_color="red",
        line_color="black",
    )

    assert isinstance(result, int)

    # Verify shape exists on the slide and has a drawing shape type
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(0)
            shape = slide.getByIndex(result)
            # Rectangle shapes have this type
            assert "RectangleShape" in shape.ShapeType or (
                "CustomShape" in shape.ShapeType
            )
        finally:
            doc.close(True)


def test_add_shape_rejects_invalid_type(tmp_path):
    from impress.content import add_shape
    from impress.core import create_presentation
    from impress.exceptions import InvalidShapeError

    path = tmp_path / "bad_shape.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidShapeError):
        add_shape(str(path), 0, "hexagon", 1.0, 1.0, 5.0, 3.0)
