import json

import pytest


def test_create_spreadsheet_creates_file(tmp_path) -> None:
    from calc.core import create_spreadsheet

    output_path = tmp_path / "sample.ods"
    create_spreadsheet(str(output_path))
    assert output_path.exists()


def test_calc_package_exports_create_spreadsheet() -> None:
    from calc import create_spreadsheet

    assert callable(create_spreadsheet)
    assert create_spreadsheet.__module__ == "calc.core"


def test_export_spreadsheet_pdf(tmp_path) -> None:
    from calc.core import create_spreadsheet, export_spreadsheet

    path = tmp_path / "export.ods"
    output = tmp_path / "export.pdf"
    create_spreadsheet(str(path))
    export_spreadsheet(str(path), str(output), "pdf")
    assert output.exists()
    assert output.stat().st_size > 0
    with open(output, "rb") as f:
        assert f.read(5) == b"%PDF-"


def test_calc_exceptions_exports_document_not_found_error() -> None:
    from calc.exceptions import CalcSkillError, DocumentNotFoundError

    assert issubclass(DocumentNotFoundError, CalcSkillError)


def test_export_spreadsheet_raises_on_missing_file(tmp_path) -> None:
    from calc.core import export_spreadsheet
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError):
        export_spreadsheet(
            str(tmp_path / "missing.ods"),
            str(tmp_path / "export.pdf"),
            "pdf",
        )


def test_session_export_unknown_format_raises(tmp_path) -> None:
    from calc import CalcSession
    from calc.core import create_spreadsheet
    from calc.exceptions import CalcSessionError

    doc_path = tmp_path / "bad_export.ods"
    create_spreadsheet(str(doc_path))

    with CalcSession(str(doc_path)) as session:
        with pytest.raises(CalcSessionError, match="Unsupported export format"):
            session.export(str(tmp_path / "out.bmp"), "bmp")


# --- JSON / XML import tests ---


def _read_cells(doc_path, sheet_index, rows, cols):
    """Read a rectangular region from an ODS file, returning list of lists."""
    from tests.calc._helpers import open_calc_doc

    with open_calc_doc(doc_path) as doc:
        sheet = doc.getSheets().getByIndex(sheet_index)
        result = []
        for r in range(rows):
            row_vals = []
            for c in range(cols):
                row_vals.append(sheet.getCellByPosition(c, r).getString())
            result.append(row_vals)
        return result


def test_create_spreadsheet_from_json_imports_data(tmp_path) -> None:
    from calc.core import create_spreadsheet

    source = tmp_path / "data.json"
    source.write_text(
        json.dumps(
            [
                {"Name": "Alice", "Age": 30},
                {"Name": "Bob", "Age": 25},
            ]
        )
    )
    out = tmp_path / "imported.ods"
    create_spreadsheet(str(out), source=str(source))

    assert out.exists()
    cells = _read_cells(out, 0, 3, 2)
    assert cells[0] == ["Name", "Age"]
    assert cells[1] == ["Alice", "30"]
    assert cells[2] == ["Bob", "25"]


def test_create_spreadsheet_from_xml_imports_data(tmp_path) -> None:
    from calc.core import create_spreadsheet

    source = tmp_path / "data.xml"
    source.write_text(
        '<?xml version="1.0"?>\n'
        "<items><item><name>X</name><value>1</value></item>"
        "<item><name>Y</name><value>2</value></item></items>"
    )
    out = tmp_path / "imported_xml.ods"
    create_spreadsheet(str(out), source=str(source))

    assert out.exists()
    cells = _read_cells(out, 0, 3, 2)
    assert cells[0] == ["name", "value"]
    assert cells[1] == ["X", "1"]
    assert cells[2] == ["Y", "2"]


def test_create_spreadsheet_from_missing_source_raises(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.exceptions import DocumentNotFoundError

    with pytest.raises(DocumentNotFoundError, match="Source file not found"):
        create_spreadsheet(
            str(tmp_path / "out.ods"),
            source=str(tmp_path / "missing.json"),
        )


def test_create_spreadsheet_from_unsupported_format_raises(tmp_path) -> None:
    from calc.core import create_spreadsheet
    from calc.exceptions import CalcSkillError

    source = tmp_path / "data.csv"
    source.write_text("a,b\n1,2\n")
    with pytest.raises(CalcSkillError, match="Unsupported import format"):
        create_spreadsheet(str(tmp_path / "out.ods"), source=str(source))


def test_create_spreadsheet_no_source_still_creates_blank(tmp_path) -> None:
    """Existing behaviour: no source parameter creates an empty spreadsheet."""
    from calc.core import create_spreadsheet

    out = tmp_path / "blank.ods"
    create_spreadsheet(str(out))
    assert out.exists()
