"""Core document lifecycle operations for Writer."""

from pathlib import Path

from uno_bridge import uno_context
from writer.exceptions import DocumentNotFoundError, WriterSkillError


EXPORT_FILTERS = {
    "pdf": "writer_pdf_Export",
    "docx": "MS Word 2007 XML",
}


def create_document(path: str) -> None:
    """Create a new Writer document at the specified path.

    Args:
        path: Output path for the new document.

    Raises:
        WriterSkillError: If document creation fails.
    """
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

        try:
            file_url = output.resolve().as_uri()
            doc.storeAsURL(file_url, ())
        finally:
            doc.close(True)


def export_document(path: str, output_path: str, export_format: str) -> None:
    """Export a Writer document to another format.

    Args:
        path: Path to the source document.
        output_path: Destination file path.
        export_format: Export format key.

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
