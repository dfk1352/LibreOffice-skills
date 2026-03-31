# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

import pytest


def test_snapshot_error_hierarchy():
    """SnapshotError is a WriterSkillError; subclasses inherit."""
    from writer.exceptions import (
        FilterError,
        InvalidPageError,
        SnapshotError,
        WriterSkillError,
    )

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
    from writer.exceptions import InvalidPageError
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidPageError):
        snapshot_page(str(doc_path), str(tmp_path / "out.png"), page=999)


def test_snapshot_page_zero_page_raises(tmp_path):
    """snapshot_page raises InvalidPageError for page=0 (must be 1-indexed)."""
    from writer.core import create_document
    from writer.exceptions import InvalidPageError
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))

    with pytest.raises(InvalidPageError):
        snapshot_page(str(doc_path), str(tmp_path / "out.png"), page=0)


def test_snapshot_page_negative_page_raises(tmp_path):
    """snapshot_page raises InvalidPageError for negative page."""
    from writer.core import create_document
    from writer.exceptions import InvalidPageError
    from writer.snapshot import snapshot_page

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
    from writer import WriterSession
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "test.odt"
    create_document(str(doc_path))
    with WriterSession(str(doc_path)) as session:
        session.insert_text("Hello Writer Snapshot")

    out_path = tmp_path / "snapshot.png"
    result = snapshot_page(str(doc_path), str(out_path))

    assert out_path.exists()
    assert out_path.stat().st_size > 0

    with open(out_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

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


def test_snapshot_page_respects_document_page_geometry(tmp_path):
    """A Letter-sized document should produce an image matching Letter aspect ratio (#11).

    Letter is 8.5 x 11 inches. A4 is 8.27 x 11.69 inches.
    The exported image dimensions should reflect the actual document page size,
    not hardcoded A4 values.
    """
    from writer import WriterSession
    from writer.core import create_document
    from writer.snapshot import snapshot_page

    doc_path = tmp_path / "letter.odt"
    create_document(str(doc_path))

    # Set page size to Letter (8.5" x 11" = 21590 x 27940 in 1/100 mm)
    # Use UNO context directly to avoid session/contextmanager interaction
    # with the LO 26.2 UNO struct __setattr__ bug on exception propagation.
    from uno_bridge import uno_context

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            styles = doc.StyleFamilies.getByName("PageStyles")
            page_style = styles.getByName("Standard")
            page_style.setPropertyValue("IsLandscape", False)
            page_style.setPropertyValue("Width", 21590)  # 8.5 inches
            page_style.setPropertyValue("Height", 27940)  # 11.0 inches
            doc.store()
        finally:
            doc.close(True)

    out_path = tmp_path / "letter_snap.png"
    result = snapshot_page(str(doc_path), str(out_path), dpi=150)

    assert out_path.exists()
    assert result.width > 0
    assert result.height > 0

    # Letter aspect ratio: 8.5/11 ≈ 0.7727
    # A4 aspect ratio: 8.27/11.69 ≈ 0.7075
    actual_ratio = result.width / result.height
    letter_ratio = 8.5 / 11.0
    a4_ratio = 8.27 / 11.69

    # The image should be closer to Letter ratio than A4 ratio
    assert abs(actual_ratio - letter_ratio) < abs(actual_ratio - a4_ratio)
