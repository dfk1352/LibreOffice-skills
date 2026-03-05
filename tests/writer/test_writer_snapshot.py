"""Tests for Writer snapshot (page-level PNG export)."""

import pytest


def test_snapshot_error_hierarchy():
    """SnapshotError is a WriterSkillError; subclasses inherit."""
    from writer.snapshot import (
        FilterError,
        InvalidPageError,
        SnapshotError,
    )
    from writer.exceptions import WriterSkillError

    assert issubclass(SnapshotError, WriterSkillError)
    assert issubclass(InvalidPageError, SnapshotError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_result_fields():
    """SnapshotResult has file_path, width, height, dpi."""
    from writer.snapshot import SnapshotResult

    result = SnapshotResult(file_path="/tmp/out.png", width=800, height=600, dpi=150)
    assert result.file_path == "/tmp/out.png"
    assert result.width == 800
    assert result.height == 600
    assert result.dpi == 150


def test_snapshot_page_invalid_page_raises(tmp_path):
    """snapshot_page raises InvalidPageError for out-of-bounds page."""
    from writer.core import create_document
    from writer.snapshot import InvalidPageError, snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidPageError):
        snapshot_page(str(doc_path), str(tmp_path / "out.png"), page=999)


def test_snapshot_page_zero_page_raises(tmp_path):
    """snapshot_page raises InvalidPageError for page=0 (must be 1-indexed)."""
    from writer.core import create_document
    from writer.snapshot import InvalidPageError, snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidPageError):
        snapshot_page(str(doc_path), str(tmp_path / "out.png"), page=0)


def test_snapshot_page_negative_page_raises(tmp_path):
    """snapshot_page raises InvalidPageError for negative page."""
    from writer.core import create_document
    from writer.snapshot import InvalidPageError, snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidPageError):
        snapshot_page(str(doc_path), str(tmp_path / "out.png"), page=-1)


def test_snapshot_page_missing_doc_raises(tmp_path):
    """snapshot_page raises DocumentNotFoundError for missing file."""
    from writer.exceptions import DocumentNotFoundError
    from writer.snapshot import snapshot_page

    with pytest.raises(DocumentNotFoundError):
        snapshot_page(str(tmp_path / "missing.odt"), str(tmp_path / "out.png"))


def test_snapshot_page_creates_png(tmp_path):
    """snapshot_page creates a valid PNG file with non-zero size."""
    from writer.core import create_document
    from writer.text import insert_text
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))
    insert_text(str(doc_path), "Hello Writer Snapshot")

    out_path = tmp_path / "snapshot.png"
    result = snapshot_page(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0

    # Verify it's a valid PNG (magic bytes)
    with open(out_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    # Verify result metadata
    assert result.file_path == str(out_path)
    assert result.width > 0
    assert result.height > 0
    assert result.dpi == 150


def test_snapshot_page_output_path_parent_created(tmp_path):
    """snapshot_page creates parent directories if needed."""
    from writer.core import create_document
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    out_path = tmp_path / "nested" / "dir" / "snapshot.png"
    snapshot_page(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0


def test_snapshot_page_custom_dpi(tmp_path):
    """snapshot_page respects custom dpi parameter."""
    from writer.core import create_document
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    out_path = tmp_path / "snapshot_300dpi.png"
    result = snapshot_page(str(doc_path), str(out_path), dpi=300)

    assert result.dpi == 300
    assert out_path.exists()
