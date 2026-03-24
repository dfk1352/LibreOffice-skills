"""Core document lifecycle operations for Calc."""

from pathlib import Path

from calc.exceptions import CalcSkillError, DocumentNotFoundError
from uno_bridge import uno_context


EXPORT_FILTERS = {
    "pdf": "calc_pdf_Export",
    "xlsx": "Calc MS Excel 2007 XML",
    "csv": "Text - txt - csv (StarCalc)",
}

_IMPORT_FORMATS = {".json", ".xml"}


def create_spreadsheet(path: str, source: str | None = None) -> None:
    """Create a new Calc spreadsheet at the specified path.

    When *source* is given it is treated as an input file to convert.
    Supported source formats: ``.json``, ``.xml``.  LibreOffice
    auto-detects the structure and maps tabular data to sheets.

    When *source* is ``None`` an empty spreadsheet is created.

    Args:
        path: Output ``.ods`` path for the new spreadsheet.
        source: Optional path to a JSON or XML file to import.

    Raises:
        DocumentNotFoundError: If *source* does not exist.
        CalcSkillError: If the source format is unsupported or
            import fails.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if source is not None:
        _import_source(source, output)
    else:
        _create_blank(output)


def export_spreadsheet(path: str, output_path: str, export_format: str) -> None:
    """Export a spreadsheet to another format.

    Args:
        path: Path to the spreadsheet file.
        output_path: Destination file path.
        export_format: Export format key.

    Raises:
        CalcSkillError: If the export format is unsupported.
    """
    if export_format not in EXPORT_FILTERS:
        raise CalcSkillError(f"Unsupported export format: {export_format}")
    file_path = Path(path)
    if not file_path.exists():
        raise DocumentNotFoundError(f"Document not found: {path}")
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            file_path.resolve().as_uri(),
            "_blank",
            0,
            (),
        )
        try:
            import uno

            filter_name = EXPORT_FILTERS[export_format]
            filter_prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
            filter_prop.Name = "FilterName"
            filter_prop.Value = filter_name
            doc.storeToURL(output.resolve().as_uri(), (filter_prop,))
        finally:
            doc.close(True)


def _create_blank(output: Path) -> None:
    """Create an empty Calc spreadsheet."""
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
        try:
            doc.storeAsURL(output.resolve().as_uri(), ())
        finally:
            doc.close(True)


def _import_source(source: str, output: Path) -> None:
    """Import a JSON or XML file via UNO auto-detection and save as ODS."""
    source_path = Path(source)
    if not source_path.exists():
        raise DocumentNotFoundError(f"Source file not found: {source}")

    suffix = source_path.suffix.lower()
    if suffix not in _IMPORT_FORMATS:
        raise CalcSkillError(f"Unsupported import format: {suffix}")

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(
            source_path.resolve().as_uri(), "_blank", 0, ()
        )
        if doc is None:
            raise CalcSkillError(f"Failed to import {source}: LO returned no document")
        try:
            doc.storeAsURL(output.resolve().as_uri(), ())
        finally:
            doc.close(True)
