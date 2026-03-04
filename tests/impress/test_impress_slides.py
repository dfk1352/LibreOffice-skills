"""Tests for Impress slide management operations."""

import pytest


def test_add_slide_appends_by_default(tmp_path):
    from libreoffice_skills.impress.core import (
        create_presentation,
        get_slide_count,
    )
    from libreoffice_skills.impress.slides import add_slide

    path = tmp_path / "slides.odp"
    create_presentation(str(path))
    assert get_slide_count(str(path)) == 1

    add_slide(str(path))
    assert get_slide_count(str(path)) == 2


def test_add_slide_with_layout(tmp_path):
    from libreoffice_skills.impress.core import (
        create_presentation,
        get_slide_count,
    )
    from libreoffice_skills.impress.slides import add_slide
    from libreoffice_skills.uno_bridge import uno_context

    path = tmp_path / "layout.odp"
    create_presentation(str(path))

    add_slide(str(path), layout="TITLE_SLIDE")

    # Verify slide was added
    assert get_slide_count(str(path)) == 2

    # Verify the layout type via UNO roundtrip
    # TITLE_SLIDE layout index is 0 in Impress
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            slide = doc.DrawPages.getByIndex(1)
            # Layout 0 = BLANK, Layout 1 = TITLE_CONTENT, etc.
            # Just verify slide has placeholder shapes typical of title slides
            has_title_placeholder = False
            for i in range(slide.Count):
                shape = slide.getByIndex(i)
                if shape.supportsService("com.sun.star.presentation.TitleTextShape"):
                    has_title_placeholder = True
                    break
            assert has_title_placeholder, (
                "Title slide should have a TitleTextShape placeholder"
            )
        finally:
            doc.close(True)


def test_add_slide_rejects_invalid_layout(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import InvalidLayoutError
    from libreoffice_skills.impress.slides import add_slide

    path = tmp_path / "bad_layout.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidLayoutError):
        add_slide(str(path), layout="NONEXISTENT_LAYOUT")


def test_delete_slide(tmp_path):
    from libreoffice_skills.impress.core import (
        create_presentation,
        get_slide_count,
    )
    from libreoffice_skills.impress.slides import add_slide, delete_slide

    path = tmp_path / "delete.odp"
    create_presentation(str(path))
    add_slide(str(path))
    assert get_slide_count(str(path)) == 2

    delete_slide(str(path), 1)
    assert get_slide_count(str(path)) == 1


def test_delete_slide_rejects_invalid_index(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import InvalidSlideIndexError
    from libreoffice_skills.impress.slides import delete_slide

    path = tmp_path / "bad_index.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidSlideIndexError):
        delete_slide(str(path), 99)


def test_move_slide(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.slides import add_slide, move_slide
    from libreoffice_skills.impress.content import add_text_box
    from libreoffice_skills.impress.slides import get_slide_inventory

    path = tmp_path / "move.odp"
    create_presentation(str(path))
    add_slide(str(path))

    # Add identifiable text to slide 0
    add_text_box(str(path), 0, "Slide Zero", 1.0, 1.0, 10.0, 3.0)

    move_slide(str(path), 0, 1)

    # The text should now be on slide 1
    inventory = get_slide_inventory(str(path), 1)
    texts = [s["text"] for s in inventory["shapes"] if s["text"]]
    assert "Slide Zero" in texts


def test_duplicate_slide(tmp_path):
    from libreoffice_skills.impress.core import (
        create_presentation,
        get_slide_count,
    )
    from libreoffice_skills.impress.slides import duplicate_slide

    path = tmp_path / "dup.odp"
    create_presentation(str(path))
    assert get_slide_count(str(path)) == 1

    duplicate_slide(str(path), 0)
    assert get_slide_count(str(path)) == 2


def test_get_slide_inventory_returns_dict(tmp_path):
    from libreoffice_skills.impress.content import add_text_box
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.slides import get_slide_inventory

    path = tmp_path / "inventory.odp"
    create_presentation(str(path))

    # Add a text box so we have a known shape to verify
    add_text_box(str(path), 0, "Inventory test", 1.0, 1.0, 8.0, 3.0)

    result = get_slide_inventory(str(path), 0)

    assert isinstance(result, dict)
    assert "shapes" in result
    assert "slide_index" in result
    assert result["slide_index"] == 0
    # Verify shapes list contains our text box
    assert isinstance(result["shapes"], list)
    assert len(result["shapes"]) >= 1
    texts = [s["text"] for s in result["shapes"]]
    assert "Inventory test" in texts
