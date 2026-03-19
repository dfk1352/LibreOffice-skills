"""Tests for Impress slide snapshot export."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

import pytest


def test_snapshot_result_fields():
    from impress.snapshot import SnapshotResult

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
    from impress.exceptions import ImpressSkillError
    from impress.snapshot import FilterError, SnapshotError

    assert issubclass(SnapshotError, ImpressSkillError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_slide_creates_png(tmp_path):
    from impress import ImpressTarget, ShapePlacement, open_impress_session
    from impress.core import create_presentation
    from impress.snapshot import snapshot_slide

    path = tmp_path / "snap.odp"
    create_presentation(str(path))

    with open_impress_session(str(path)) as session:
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=0),
            "Snapshot test",
            ShapePlacement(2.0, 2.0, 10.0, 5.0),
            name="SnapshotBox",
        )

    out_path = tmp_path / "snapshot.png"
    result = snapshot_slide(str(path), 0, str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0
    with open(out_path, "rb") as handle:
        assert handle.read(8) == b"\x89PNG\r\n\x1a\n"
    assert result.file_path == str(out_path)
    assert result.width > 0
    assert result.height > 0


def test_snapshot_slide_invalid_index_raises(tmp_path):
    from impress.core import create_presentation
    from impress.exceptions import InvalidSlideIndexError
    from impress.snapshot import snapshot_slide

    path = tmp_path / "snap_bad.odp"
    create_presentation(str(path))

    with pytest.raises(InvalidSlideIndexError):
        snapshot_slide(str(path), 99, str(tmp_path / "out.png"))


def test_snapshot_slide_missing_doc_raises(tmp_path):
    from impress.exceptions import DocumentNotFoundError
    from impress.snapshot import snapshot_slide

    with pytest.raises(DocumentNotFoundError):
        snapshot_slide(
            str(tmp_path / "missing.odp"),
            0,
            str(tmp_path / "out.png"),
        )


def test_snapshot_slide_custom_dimensions(tmp_path):
    from impress.core import create_presentation
    from impress.snapshot import snapshot_slide

    path = tmp_path / "snap_hd.odp"
    create_presentation(str(path))

    out_path = tmp_path / "snapshot_hd.png"
    result = snapshot_slide(str(path), 0, str(out_path), width=960, height=540)

    assert result.width > 0
    assert result.height > 0
    assert out_path.exists()
    assert out_path.stat().st_size > 0
