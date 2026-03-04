"""Tests for Impress speaker notes operations."""

import pytest


def test_set_and_get_notes(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.notes import get_notes, set_notes

    path = tmp_path / "notes.odp"
    create_presentation(str(path))

    set_notes(str(path), 0, "Remember to mention the demo")
    result = get_notes(str(path), 0)

    assert result == "Remember to mention the demo"


def test_get_notes_empty_by_default(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.notes import get_notes

    path = tmp_path / "empty_notes.odp"
    create_presentation(str(path))

    result = get_notes(str(path), 0)

    assert result == ""


def test_set_notes_invalid_slide_raises(tmp_path):
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import InvalidSlideIndexError
    from libreoffice_skills.impress.notes import set_notes

    path = tmp_path / "bad_notes.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidSlideIndexError):
        set_notes(str(path), 99, "This should fail")
