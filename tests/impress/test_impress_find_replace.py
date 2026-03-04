"""Tests for Impress find & replace operations."""


def test_find_replace_in_presentation(tmp_path):
    from libreoffice_skills.impress.content import add_text_box
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.find_replace import find_replace
    from libreoffice_skills.impress.slides import get_slide_inventory

    path = tmp_path / "find_replace.odp"
    create_presentation(str(path))
    add_text_box(str(path), 0, "Hello World", 1.0, 1.0, 10.0, 3.0)

    count = find_replace(str(path), "Hello", "Goodbye")

    assert count >= 1

    inventory = get_slide_inventory(str(path), 0)
    texts = [s["text"] for s in inventory["shapes"]]
    assert any("Goodbye" in t for t in texts if t)


def test_find_replace_returns_zero_for_no_match(tmp_path):
    from libreoffice_skills.impress.content import add_text_box
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.find_replace import find_replace

    path = tmp_path / "no_match.odp"
    create_presentation(str(path))
    add_text_box(str(path), 0, "Hello World", 1.0, 1.0, 10.0, 3.0)

    count = find_replace(str(path), "ZZZZZ", "replacement")

    assert count == 0
