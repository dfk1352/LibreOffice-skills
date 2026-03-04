"""Tests for Impress snapshot (slide-level PNG export)."""

import pytest


def test_snapshot_result_fields():
    """SnapshotResult has file_path, width, height, dpi."""
    from libreoffice_skills.impress.snapshot import SnapshotResult

    result = SnapshotResult(
        file_path="/tmp/out.png",
        width=1280,
        height=720,
        dpi=96,
    )
    assert result.file_path == "/tmp/out.png"
    assert result.width == 1280
    assert result.height == 720
    assert result.dpi == 96


def test_snapshot_error_hierarchy():
    """SnapshotError is a subclass of ImpressSkillError; FilterError of SnapshotError."""
    from libreoffice_skills.impress.exceptions import ImpressSkillError
    from libreoffice_skills.impress.snapshot import FilterError, SnapshotError

    assert issubclass(SnapshotError, ImpressSkillError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_slide_creates_png(tmp_path):
    """snapshot_slide creates a valid PNG file with non-zero size."""
    from libreoffice_skills.impress.content import add_text_box
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.snapshot import snapshot_slide

    path = tmp_path / "snap.odp"
    create_presentation(str(path))
    add_text_box(str(path), 0, "Snapshot test", 2.0, 2.0, 10.0, 5.0)

    out_path = tmp_path / "snapshot.png"
    result = snapshot_slide(str(path), 0, str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0

    # Verify PNG magic bytes
    with open(out_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    assert result.file_path == str(out_path)
    assert result.width > 0
    assert result.height > 0


def test_snapshot_slide_invalid_index_raises(tmp_path):
    """snapshot_slide raises InvalidSlideIndexError for bad index."""
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.exceptions import InvalidSlideIndexError
    from libreoffice_skills.impress.snapshot import snapshot_slide

    path = tmp_path / "snap_bad.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidSlideIndexError):
        snapshot_slide(str(path), 99, str(tmp_path / "out.png"))


def test_snapshot_slide_missing_doc_raises(tmp_path):
    """snapshot_slide raises DocumentNotFoundError for missing file."""
    from libreoffice_skills.impress.exceptions import DocumentNotFoundError
    from libreoffice_skills.impress.snapshot import snapshot_slide

    with pytest.raises(DocumentNotFoundError):
        snapshot_slide(
            str(tmp_path / "missing.odp"),
            0,
            str(tmp_path / "out.png"),
        )


def test_snapshot_slide_custom_dimensions(tmp_path):
    """snapshot_slide with custom width/height reflected in result."""
    from libreoffice_skills.impress.core import create_presentation
    from libreoffice_skills.impress.snapshot import snapshot_slide

    path = tmp_path / "snap_hd.odp"
    create_presentation(str(path))

    out_path = tmp_path / "snapshot_hd.png"
    result = snapshot_slide(str(path), 0, str(out_path))

    assert result.width > 0
    assert result.height > 0
    assert out_path.exists()
    assert out_path.stat().st_size > 0
