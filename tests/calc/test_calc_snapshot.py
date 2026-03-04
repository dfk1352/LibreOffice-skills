"""Tests for Calc snapshot (area-level PNG export)."""

import pytest


def test_snapshot_error_hierarchy():
    """SnapshotError is a CalcSkillError; subclasses inherit."""
    from libreoffice_skills.calc.snapshot import (
        FilterError,
        InvalidAreaError,
        InvalidSheetError,
        SnapshotError,
    )
    from libreoffice_skills.calc.exceptions import CalcSkillError

    assert issubclass(SnapshotError, CalcSkillError)
    assert issubclass(InvalidSheetError, SnapshotError)
    assert issubclass(InvalidAreaError, SnapshotError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_result_fields():
    """SnapshotResult has file_path, width, height, dpi."""
    from libreoffice_skills.calc.snapshot import SnapshotResult

    result = SnapshotResult(file_path="/tmp/out.png", width=800, height=600, dpi=150)
    assert result.file_path == "/tmp/out.png"
    assert result.width == 800
    assert result.height == 600
    assert result.dpi == 150


def test_snapshot_area_invalid_sheet_raises(tmp_path):
    """snapshot_area raises InvalidSheetError for nonexistent sheet."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import InvalidSheetError, snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidSheetError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), sheet="NonExistent")


def test_snapshot_area_negative_row_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative row."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import InvalidAreaError, snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), row=-1)


def test_snapshot_area_negative_col_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative col."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import InvalidAreaError, snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), col=-1)


def test_snapshot_area_negative_width_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative width."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import InvalidAreaError, snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), width=-100)


def test_snapshot_area_negative_height_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative height."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import InvalidAreaError, snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), height=-100)


def test_snapshot_area_missing_doc_raises(tmp_path):
    """snapshot_area raises DocumentNotFoundError for missing file."""
    from libreoffice_skills.calc.exceptions import CalcSkillError
    from libreoffice_skills.calc.snapshot import snapshot_area

    with pytest.raises(CalcSkillError):
        snapshot_area(str(tmp_path / "missing.ods"), str(tmp_path / "out.png"))


def test_snapshot_area_creates_png(tmp_path):
    """snapshot_area creates a valid PNG file with non-zero size."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.cells import set_cell
    from libreoffice_skills.calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))
    set_cell(str(doc_path), "Sheet1", 0, 0, "Hello", type="text")
    set_cell(str(doc_path), "Sheet1", 1, 0, 42, type="number")

    out_path = tmp_path / "snapshot.png"
    result = snapshot_area(str(doc_path), str(out_path))

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


def test_snapshot_area_output_path_parent_created(tmp_path):
    """snapshot_area creates parent directories if needed."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    out_path = tmp_path / "nested" / "dir" / "snapshot.png"
    snapshot_area(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0


def test_snapshot_area_custom_dpi(tmp_path):
    """snapshot_area respects custom dpi parameter."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    out_path = tmp_path / "snapshot_300dpi.png"
    result = snapshot_area(str(doc_path), str(out_path), dpi=300)

    assert result.dpi == 300
    assert out_path.exists()


def test_snapshot_area_with_cell_anchor(tmp_path):
    """snapshot_area with row/col captures from the specified cell position."""
    from libreoffice_skills.calc.core import create_spreadsheet
    from libreoffice_skills.calc.cells import set_cell
    from libreoffice_skills.calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))
    set_cell(str(doc_path), "Sheet1", 5, 3, "Anchored", type="text")

    out_path = tmp_path / "anchored.png"
    result = snapshot_area(str(doc_path), str(out_path), row=5, col=3)

    assert out_path.exists()
    assert out_path.stat().st_size > 0
    assert result.width > 0
    assert result.height > 0
