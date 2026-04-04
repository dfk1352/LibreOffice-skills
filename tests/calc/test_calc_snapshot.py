# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

import pytest


def test_snapshot_error_hierarchy():
    """SnapshotError is a CalcSkillError; subclasses inherit."""
    from calc.exceptions import (
        CalcSkillError,
        FilterError,
        InvalidAreaError,
        InvalidSheetError,
        SnapshotError,
    )

    assert issubclass(SnapshotError, CalcSkillError)
    assert issubclass(InvalidSheetError, SnapshotError)
    assert issubclass(InvalidAreaError, SnapshotError)
    assert issubclass(FilterError, SnapshotError)


def test_snapshot_result_fields():
    """SnapshotResult has file_path, width, height, dpi."""
    from calc.snapshot import SnapshotResult

    result = SnapshotResult(file_path="/tmp/out.png", width=800, height=600, dpi=150)
    assert result.file_path == "/tmp/out.png"
    assert result.width == 800
    assert result.height == 600
    assert result.dpi == 150


def test_snapshot_area_invalid_sheet_raises(tmp_path):
    """snapshot_area raises InvalidSheetError for nonexistent sheet."""
    from calc.core import create_spreadsheet
    from calc.exceptions import InvalidSheetError
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidSheetError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), sheet="NonExistent")


def test_snapshot_area_negative_row_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative row."""
    from calc.core import create_spreadsheet
    from calc.exceptions import InvalidAreaError
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), row=-1)


def test_snapshot_area_negative_col_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative col."""
    from calc.core import create_spreadsheet
    from calc.exceptions import InvalidAreaError
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), col=-1)


def test_snapshot_area_negative_width_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative width."""
    from calc.core import create_spreadsheet
    from calc.exceptions import InvalidAreaError
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), width=-100)


def test_snapshot_area_negative_height_raises(tmp_path):
    """snapshot_area raises InvalidAreaError for negative height."""
    from calc.core import create_spreadsheet
    from calc.exceptions import InvalidAreaError
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    with pytest.raises(InvalidAreaError):
        snapshot_area(str(doc_path), str(tmp_path / "out.png"), height=-100)


def test_snapshot_area_missing_doc_raises(tmp_path):
    """snapshot_area raises DocumentNotFoundError for missing file."""
    from calc.exceptions import CalcSkillError
    from calc.snapshot import snapshot_area

    with pytest.raises(CalcSkillError):
        snapshot_area(str(tmp_path / "missing.ods"), str(tmp_path / "out.png"))


def test_snapshot_area_creates_png(tmp_path):
    """snapshot_area creates a valid PNG file with non-zero size."""
    from calc import CalcTarget, CalcSession
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))
    with CalcSession(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Hello",
            value_type="text",
        )
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=1, col=0),
            42,
            value_type="number",
        )

    out_path = tmp_path / "snapshot.png"
    result = snapshot_area(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0

    with open(out_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    assert result.file_path == str(out_path)
    assert result.width > 0
    assert result.height > 0
    assert result.dpi == 150


def test_snapshot_area_output_path_parent_created(tmp_path):
    """snapshot_area creates parent directories if needed."""
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    out_path = tmp_path / "nested" / "dir" / "snapshot.png"
    snapshot_area(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0


def test_snapshot_area_custom_dpi(tmp_path):
    """snapshot_area respects custom dpi parameter."""
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))

    out_path = tmp_path / "snapshot_300dpi.png"
    result = snapshot_area(str(doc_path), str(out_path), dpi=300)

    assert result.dpi == 300
    assert out_path.exists()


def test_snapshot_area_with_cell_anchor(tmp_path):
    """snapshot_area with row/col captures from the specified cell position."""
    from calc import CalcTarget, CalcSession
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "test.ods"
    create_spreadsheet(str(doc_path))
    with CalcSession(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=5, col=3),
            "Anchored",
            value_type="text",
        )

    out_path = tmp_path / "anchored.png"
    result = snapshot_area(str(doc_path), str(out_path), row=5, col=3)

    assert out_path.exists()
    assert out_path.stat().st_size > 0
    assert result.width > 0
    assert result.height > 0


def test_snapshot_area_with_nondefault_column_widths(tmp_path):
    """snapshot_area runs without error when column widths differ from defaults (#11)."""
    from calc import CalcTarget, CalcSession
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "wide_cols.ods"
    create_spreadsheet(str(doc_path))

    # Set column A to a very wide width and write data
    with CalcSession(str(doc_path)) as session:
        sheet_obj = session.doc.Sheets.getByName("Sheet1")
        col_obj = sheet_obj.Columns.getByIndex(0)
        col_obj.Width = 10000  # ~10 cm, much wider than default
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=0, col=0),
            "Wide column data",
            value_type="text",
        )

    out_path = tmp_path / "wide_snap.png"
    result = snapshot_area(str(doc_path), str(out_path), width=800, height=600)

    assert out_path.exists()
    assert result.width > 0
    assert result.height > 0


def test_snapshot_area_none_dimensions_include_used_range(tmp_path):
    from calc import CalcTarget, CalcSession
    from calc.core import create_spreadsheet
    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "used_range.ods"
    create_spreadsheet(str(doc_path))

    with CalcSession(str(doc_path)) as session:
        session.write_cell(
            CalcTarget(kind="cell", sheet="Sheet1", row=30, col=15),
            "Far cell",
            value_type="text",
        )

    preview_path = tmp_path / "preview.png"
    full_path = tmp_path / "full.png"
    preview = snapshot_area(
        str(doc_path), str(preview_path), row=0, col=0, width=800, height=600
    )
    full = snapshot_area(
        str(doc_path), str(full_path), row=0, col=0, width=None, height=None
    )

    assert full.width >= preview.width
    assert full.height >= preview.height


def test_snapshot_area_none_dimensions_select_true_used_range(monkeypatch, tmp_path):
    from types import SimpleNamespace

    from calc.snapshot import snapshot_area

    doc_path = tmp_path / "fake.ods"
    doc_path.write_bytes(b"fake")

    class _RangeAddress:
        StartColumn = 0
        StartRow = 0
        EndColumn = 15
        EndRow = 30

    class _Cursor:
        def gotoStartOfUsedArea(self, expand):
            return None

        def gotoEndOfUsedArea(self, expand):
            return None

        def getRangeAddress(self):
            return _RangeAddress()

    class _Sheet:
        def __init__(self):
            self.selected = None
            self.Columns = SimpleNamespace(
                getByIndex=lambda index: SimpleNamespace(Width=1000)
            )
            self.Rows = SimpleNamespace(
                getByIndex=lambda index: SimpleNamespace(Height=500)
            )

        def createCursor(self):
            return _Cursor()

        def getCellRangeByPosition(self, start_col, start_row, end_col, end_row):
            self.selected = (start_col, start_row, end_col, end_row)
            return object()

    class _Sheets:
        Count = 1

        def __init__(self, sheet):
            self._sheet = sheet

        def hasByName(self, name):
            return name == "Sheet1"

        def getByName(self, name):
            return self._sheet

        def getByIndex(self, index):
            return self._sheet

    class _Controller:
        def __init__(self):
            self.selection = None

        def setActiveSheet(self, sheet):
            return None

        def select(self, cell_range):
            self.selection = cell_range

    class _Doc:
        def __init__(self, sheet):
            self.Sheets = _Sheets(sheet)
            self._controller = _Controller()

        def getCurrentController(self):
            return self._controller

        def storeToURL(self, *_args):
            return None

        def close(self, *_args):
            return None

    class _Desktop:
        def __init__(self, doc):
            self._doc = doc

        def loadComponentFromURL(self, *_args):
            return self._doc

    class _UnoContext:
        def __init__(self, desktop):
            self._desktop = desktop

        def __enter__(self):
            return self._desktop

        def __exit__(self, exc_type, exc, tb):
            return False

    sheet = _Sheet()
    doc = _Doc(sheet)
    desktop = _Desktop(doc)

    monkeypatch.setattr("calc.snapshot.uno_context", lambda: _UnoContext(desktop))
    monkeypatch.setitem(
        __import__("sys").modules,
        "uno",
        SimpleNamespace(
            createUnoStruct=lambda _name: SimpleNamespace(Name="", Value=None),
            Any=lambda _type, value: value,
        ),
    )
    monkeypatch.setattr("calc.snapshot._read_png_dimensions", lambda _path: (1600, 900))

    result = snapshot_area(str(doc_path), str(tmp_path / "out.png"), sheet="Sheet1")

    assert sheet.selected == (0, 0, 15, 30)
    assert result.width == 1600
    assert result.height == 900
