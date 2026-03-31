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
    from impress.exceptions import FilterError, ImpressSkillError, SnapshotError

    assert issubclass(SnapshotError, ImpressSkillError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_slide_creates_png(tmp_path):
    from impress import ImpressSession, ImpressTarget, ShapePlacement
    from impress.core import create_presentation
    from impress.snapshot import snapshot_slide

    path = tmp_path / "snap.odp"
    create_presentation(str(path))

    with ImpressSession(str(path)) as session:
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


def test_convert_to_pngs_raises_snapshot_error_on_timeout(tmp_path, monkeypatch):
    """subprocess.TimeoutExpired propagates as FilterError."""
    import subprocess

    from impress.exceptions import FilterError
    from impress.snapshot import _convert_to_pngs

    def mock_run(*args, **kwargs):
        raise subprocess.TimeoutExpired(cmd="soffice", timeout=120)

    monkeypatch.setattr(subprocess, "run", mock_run)

    with pytest.raises(FilterError, match="timed out"):
        _convert_to_pngs(str(tmp_path / "dummy.odp"), tmp_path)


def test_snapshot_slide_preserves_4_3_aspect_ratio(tmp_path):
    """A 4:3 presentation must produce a non-distorted snapshot (#11).

    Default presentations are 16:9 (25.4 x 14.29 cm).
    A 4:3 presentation (25.4 x 19.05 cm) should produce an image
    whose aspect ratio is closer to 4:3 than 16:9.
    """
    from impress import ImpressSession
    from impress.core import create_presentation
    from impress.snapshot import snapshot_slide

    doc_path = tmp_path / "four_three.odp"
    create_presentation(str(doc_path))

    # Change slide size to 4:3 (25400 x 19050 in 1/100 mm)
    with ImpressSession(str(doc_path)) as session:
        page = session.doc.DrawPages.getByIndex(0)
        page.Width = 25400  # 25.4 cm
        page.Height = 19050  # 19.05 cm

    out_path = tmp_path / "four_three_snap.png"
    result = snapshot_slide(str(doc_path), 0, str(out_path))

    assert out_path.exists()
    assert result.width > 0
    assert result.height > 0

    # 4:3 ratio ≈ 1.333, 16:9 ratio ≈ 1.778
    actual_ratio = result.width / result.height
    four_three_ratio = 4.0 / 3.0
    sixteen_nine_ratio = 16.0 / 9.0

    # The image should be closer to 4:3 than 16:9
    assert abs(actual_ratio - four_three_ratio) < abs(actual_ratio - sixteen_nine_ratio)
