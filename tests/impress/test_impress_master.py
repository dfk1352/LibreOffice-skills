"""Tests for Impress master page operations."""

import pytest


def test_list_master_pages(tmp_path):
    from impress.core import create_presentation
    from impress.master import list_master_pages

    path = tmp_path / "master.odp"
    create_presentation(str(path))

    result = list_master_pages(str(path))

    assert isinstance(result, list)
    assert len(result) >= 1
    # Each entry must be a non-empty string (master page name)
    for name in result:
        assert isinstance(name, str)
        assert len(name) > 0


def test_apply_master_page(tmp_path):
    from impress.core import create_presentation
    from impress.master import (
        apply_master_page,
        list_master_pages,
    )
    from impress.slides import add_slide
    from uno_bridge import uno_context

    path = tmp_path / "apply_master.odp"
    create_presentation(str(path))
    add_slide(str(path))
    add_slide(str(path))

    masters = list_master_pages(str(path))
    master_name = masters[0]
    apply_master_page(str(path), master_name)

    # Verify master page was applied to all slides via UNO roundtrip
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(path.resolve().as_uri(), "_blank", 0, ())
        try:
            for i in range(doc.DrawPages.Count):
                slide = doc.DrawPages.getByIndex(i)
                assert slide.MasterPage.Name == master_name, (
                    f"Slide {i}: expected master {master_name!r}, "
                    f"got {slide.MasterPage.Name!r}"
                )
        finally:
            doc.close(True)


def test_apply_master_page_invalid_name_raises(tmp_path):
    from impress.core import create_presentation
    from impress.exceptions import MasterNotFoundError
    from impress.master import apply_master_page

    path = tmp_path / "bad_master.odp"
    create_presentation(str(path))

    with pytest.raises(MasterNotFoundError):
        apply_master_page(str(path), "NonexistentMasterPage")


def test_import_master_from_template(tmp_path):
    from impress.core import create_presentation
    from impress.master import (
        import_master_from_template,
        list_master_pages,
        set_master_background,
    )
    from uno_bridge import uno_context

    target_path = tmp_path / "target.odp"
    template_path = tmp_path / "template.odp"
    create_presentation(str(target_path))
    create_presentation(str(template_path))

    imported_name = import_master_from_template(
        str(target_path),
        str(template_path),
    )

    assert isinstance(imported_name, str)
    masters = list_master_pages(str(target_path))
    assert imported_name in masters

    set_master_background(str(target_path), imported_name, "lightsteelblue")

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            target_path.resolve().as_uri(), "_blank", 0, ()
        )
        try:
            masters = doc.MasterPages
            found = None
            for i in range(masters.Count):
                master = masters.getByIndex(i)
                if master.Name == imported_name:
                    found = master
                    break
            assert found is not None
            assert found.Background.FillColor == 0xB0C4DE
        finally:
            doc.close(True)
