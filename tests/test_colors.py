"""Tests for shared color resolution."""

import pytest


def test_resolve_color_accepts_int() -> None:
    from libreoffice_skills.colors import resolve_color

    assert resolve_color(0x112233) == 0x112233


def test_resolve_color_accepts_name_case_insensitive() -> None:
    from libreoffice_skills.colors import resolve_color

    assert resolve_color("Navy") == 0x000080
    assert resolve_color("light steel blue") == 0xB0C4DE


def test_resolve_color_rejects_unknown() -> None:
    from libreoffice_skills.colors import resolve_color

    with pytest.raises(ValueError):
        resolve_color("not-a-color")
