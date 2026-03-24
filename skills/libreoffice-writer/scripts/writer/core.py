"""Core document lifecycle operations for Writer."""

from pathlib import Path

from uno_bridge import uno_context
from writer.exceptions import DocumentNotFoundError, WriterSkillError


EXPORT_FILTERS = {
    "pdf": "writer_pdf_Export",
    "docx": "MS Word 2007 XML",
    "md": "Markdown",
}

_IMPORT_FILTERS = {
    ".md": "Markdown",
}


def create_document(path: str, source: str | None = None) -> None:
    """Create a new Writer document at the specified path.

    When *source* is given it is treated as an input file to convert.
    The source format is detected from its extension (currently ``.md``
    for Markdown).  When *source* is ``None`` an empty document is
    created.

    Args:
        path: Output ``.odt`` path for the new document.
        source: Optional path to a source file to import.

    Raises:
        DocumentNotFoundError: If *source* does not exist.
        WriterSkillError: If the source format is unsupported or
            document creation fails.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if source is not None:
        _import_source(source, output)
    else:
        _create_blank(output)


def export_document(path: str, output_path: str, export_format: str) -> None:
    """Export a Writer document to another format.

    Args:
        path: Path to the source document.
        output_path: Destination file path.
        export_format: Export format key (``"pdf"``, ``"docx"``, ``"md"``).

    Raises:
        DocumentNotFoundError: If the source document does not exist.
        WriterSkillError: If the export format is unsupported.
    """
    file_path = Path(path)
    if not file_path.exists():
        raise DocumentNotFoundError(f"Document not found: {path}")
    if export_format not in EXPORT_FILTERS:
        raise WriterSkillError(f"Unsupported export format: {export_format}")

    from writer.session import WriterSession

    with WriterSession(str(file_path)) as session:
        session.export(output_path, export_format)


def _create_blank(output: Path) -> None:
    """Create an empty Writer document."""
    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        try:
            doc.storeAsURL(output.resolve().as_uri(), ())
        finally:
            doc.close(True)


def _import_source(source: str, output: Path) -> None:
    """Import a file via UNO filter and save as ODT."""
    source_path = Path(source)
    if not source_path.exists():
        raise DocumentNotFoundError(f"Source file not found: {source}")

    suffix = source_path.suffix.lower()
    if suffix not in _IMPORT_FILTERS:
        raise WriterSkillError(f"Unsupported import format: {suffix}")

    with uno_context() as desktop:
        import uno  # type: ignore[import-not-found]

        filter_prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        filter_prop.Name = "FilterName"
        filter_prop.Value = _IMPORT_FILTERS[suffix]

        doc = desktop.loadComponentFromURL(
            source_path.resolve().as_uri(), "_blank", 0, (filter_prop,)
        )
        try:
            doc.storeAsURL(output.resolve().as_uri(), ())
        finally:
            doc.close(True)
