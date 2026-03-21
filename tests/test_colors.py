import pytest


def test_resolve_color_accepts_int() -> None:
    from colors import resolve_color

    assert resolve_color(0x112233) == 0x112233


def test_resolve_color_accepts_name_case_insensitive() -> None:
    from colors import resolve_color

    assert resolve_color("Navy") == 0x000080
    assert resolve_color("light steel blue") == 0xB0C4DE


def test_resolve_color_strips_underscores() -> None:
    from colors import resolve_color

    assert resolve_color("light_steel_blue") == 0xB0C4DE
    assert resolve_color("dark_olive_green") == 0x556B2F


def test_resolve_color_rejects_unknown() -> None:
    from colors import resolve_color

    with pytest.raises(ValueError):
        resolve_color("not-a-color")
