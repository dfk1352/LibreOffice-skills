"""Test Writer image operations."""

import pytest


def test_insert_image_requires_existing_file(tmp_path):
    from writer.core import create_document
    from writer.images import insert_image
    from writer.exceptions import ImageNotFoundError

    doc_path = tmp_path / "sample.odt"
    create_document(str(doc_path))

    with pytest.raises(ImageNotFoundError):
        insert_image(str(doc_path), str(tmp_path / "missing.png"))


def test_insert_image_adds_graphic(tmp_path):
    from PIL import Image
    from writer.core import create_document
    from writer.images import insert_image
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_image.odt"
    create_document(str(doc_path))

    # Create a simple test image
    img_path = tmp_path / "test.png"
    img = Image.new("RGB", (100, 100), color=0)
    img.save(img_path)

    # Insert image
    insert_image(str(doc_path), str(img_path))

    # Verify image exists in the document via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            graphics = doc.getGraphicObjects()
            assert graphics.getCount() >= 1
        finally:
            doc.close(True)


def test_insert_image_with_size(tmp_path):
    from PIL import Image
    from writer.core import create_document
    from writer.images import insert_image
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_image_sized.odt"
    create_document(str(doc_path))

    # Create a simple test image
    img_path = tmp_path / "test.png"
    img = Image.new("RGB", (100, 100), color=0)
    img.save(img_path)

    # Insert image with specific size (in 1/100mm units)
    insert_image(str(doc_path), str(img_path), width=5000, height=5000)

    # Verify image dimensions via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            graphics = doc.getGraphicObjects()
            assert graphics.getCount() >= 1
            names = graphics.getElementNames()
            graphic = graphics.getByName(names[0])
            size = graphic.Size
            assert abs(size.Width - 5000) <= 1
            assert abs(size.Height - 5000) <= 1
        finally:
            doc.close(True)


def test_insert_image_at_index(tmp_path):
    from PIL import Image
    from writer.core import create_document
    from writer.text import insert_text
    from writer.images import insert_image
    from uno_bridge import uno_context

    doc_path = tmp_path / "test_image_index.odt"
    create_document(str(doc_path))

    img_path = tmp_path / "test_index.png"
    img = Image.new("RGB", (10, 10), color=0)
    img.save(img_path)

    insert_text(str(doc_path), "After", position=None)
    insert_image(str(doc_path), str(img_path), position=0)

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            graphics = doc.getGraphicObjects()
            names = graphics.getElementNames()
            assert names
            graphic = graphics.getByName(names[0])
            anchor = graphic.Anchor
            text = anchor.getText()
            start = text.getStart()
            assert text.compareRegionStarts(anchor, start) == 0
        finally:
            doc.close(True)
