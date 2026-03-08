"""Integration workflow tests for Writer skill."""

# pyright: reportMissingImports=false

from pathlib import Path

from PIL import Image


def run_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build a Writer document exercising every tool, then snapshot before/after.

    This is the single function that produces all inspectable output files.
    It exercises: create, insert_text, apply_formatting, add_table,
    insert_image, set/get_metadata, read_document_text, and snapshot_page.

    Args:
        output_dir: Directory where all output files are written.

    Returns:
        Dict mapping logical names to output file paths:
            "document"        -> report.odt
            "snapshot_before" -> writer_snapshot_before.png
            "snapshot_after"  -> writer_snapshot_after.png
    """
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.formatting import apply_formatting
    from writer.images import insert_image
    from writer.metadata import (
        get_metadata,
        set_metadata,
    )
    from writer.snapshot import snapshot_page
    from writer.tables import add_table
    from writer.text import insert_text

    output_dir.mkdir(parents=True, exist_ok=True)

    # --- Build the document ---
    doc_path = output_dir / "report.odt"
    create_document(str(doc_path))

    metadata = {
        "title": "Integration Test Report",
        "author": "Test Suite",
        "subject": "Testing",
    }
    set_metadata(str(doc_path), metadata)

    insert_text(str(doc_path), "Integration Test Report", position=None)
    apply_formatting(
        str(doc_path),
        {"bold": True, "font_size": 18, "align": "center"},
        selection="last_paragraph",
    )
    insert_text(str(doc_path), "\n\n", position=None)

    insert_text(str(doc_path), "This is a test document.", position=None)
    apply_formatting(
        str(doc_path),
        {"align": "left", "bold": False, "font_size": 12},
        selection="last_paragraph",
    )
    insert_text(str(doc_path), "\n\n", position=None)

    table_data = [
        ["Feature", "Status"],
        ["Text", "Working"],
        ["Tables", "Working"],
    ]
    add_table(str(doc_path), 3, 2, table_data)

    img_path = output_dir / "test_image.png"
    img = Image.new("RGB", (50, 50), color=0)
    img.save(img_path)
    insert_image(str(doc_path), str(img_path), position=None)

    # Verify metadata round-trip
    retrieved_meta = get_metadata(str(doc_path))
    assert retrieved_meta["title"] == "Integration Test Report"
    assert retrieved_meta["author"] == "Test Suite"

    # Verify text content
    content = read_document_text(str(doc_path))
    assert "Integration Test Report" in content
    assert "This is a test document" in content

    # --- Snapshot BEFORE formatting changes ---
    before_path = output_dir / "writer_snapshot_before.png"
    snapshot_page(str(doc_path), str(before_path), page=1)

    # --- Apply visible formatting changes ---
    apply_formatting(
        str(doc_path),
        {"italic": True, "color": 0x003399, "font_size": 24},
        selection="all",
    )
    insert_text(
        str(doc_path),
        "\n\nThis paragraph was added after the initial snapshot.",
        position=None,
    )
    apply_formatting(
        str(doc_path),
        {"bold": True, "color": 0x990000, "font_size": 14},
        selection="last_paragraph",
    )

    # --- Snapshot AFTER formatting changes ---
    after_path = output_dir / "writer_snapshot_after.png"
    snapshot_page(str(doc_path), str(after_path), page=1)

    return {
        "document": doc_path,
        "snapshot_before": before_path,
        "snapshot_after": after_path,
    }


# ---------------------------------------------------------------------------
# Deterministic assertion tests
# ---------------------------------------------------------------------------


def test_multiple_operations_on_same_document(tmp_path):
    """Test that multiple operations can be performed sequentially."""
    from writer.core import (
        create_document,
        read_document_text,
    )
    from writer.text import (
        insert_text,
        replace_text,
    )

    doc_path = tmp_path / "multi_op.odt"
    create_document(str(doc_path))

    # Perform multiple operations
    insert_text(str(doc_path), "First line\n")
    insert_text(str(doc_path), "Second line\n", position=None)
    insert_text(str(doc_path), "Third line\n", position=None)

    # Replace text
    replace_text(str(doc_path), "Second", "SECOND")

    # Verify all operations
    content = read_document_text(str(doc_path))
    assert "First line" in content
    assert "SECOND line" in content
    assert "Third line" in content


def test_title_and_body_are_separate_paragraphs(tmp_path):
    """Assert title paragraph is centered and body paragraph is not."""
    from uno_bridge import uno_context

    outputs = run_end_to_end_workflow(tmp_path)
    doc_path = outputs["document"]

    with uno_context() as desktop:
        doc = desktop.loadComponentFromURL(doc_path.resolve().as_uri(), "_blank", 0, ())
        try:
            paragraphs = []
            enum = doc.Text.createEnumeration()
            while enum.hasMoreElements():
                item = enum.nextElement()
                if hasattr(item, "getString"):
                    text = item.getString().strip()
                    if text:
                        paragraphs.append(item)

            assert len(paragraphs) >= 2
            assert paragraphs[0].ParaAdjust == 2
            assert paragraphs[1].ParaAdjust != 2

            cursor = doc.Text.createTextCursorByRange(paragraphs[1].getStart())
            cursor.gotoEndOfParagraph(True)
            assert cursor.CharWeight == 100
            assert cursor.CharHeight >= 10
        finally:
            doc.close(True)


def test_writer_snapshot_in_workflow(tmp_path):
    """Assert snapshot_page produces a valid PNG for a Writer document."""
    from writer.snapshot import snapshot_page

    outputs = run_end_to_end_workflow(tmp_path)
    # The workflow already produces snapshots; verify the before snapshot
    snapshot_path = outputs["snapshot_before"]

    assert snapshot_path.exists()
    assert snapshot_path.stat().st_size > 0

    # Verify PNG magic bytes
    with open(snapshot_path, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"

    # Also verify we can take an independent snapshot
    independent_path = tmp_path / "independent_snapshot.png"
    result = snapshot_page(str(outputs["document"]), str(independent_path), page=1)
    assert independent_path.exists()
    assert result.width > 0
    assert result.height > 0
    assert result.dpi == 150


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable output files in test-output/writer/.

    Calls run_end_to_end_workflow which builds a Writer document,
    snapshots it before and after formatting changes. Assertions
    verify that all output files exist and the snapshots differ.

    Output files:
        test-output/writer/report.odt                 - the document
        test-output/writer/writer_snapshot_before.png  - before formatting changes
        test-output/writer/writer_snapshot_after.png   - after formatting changes
    """
    output_dir = Path("test-output/writer")

    # Clean up previous runs
    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_end_to_end_workflow(output_dir)

    # Assert all output files exist and are non-empty
    for key in ("document", "snapshot_before", "snapshot_after"):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    # Assert before and after snapshots differ (formatting changes are visible)
    before_bytes = outputs["snapshot_before"].read_bytes()
    after_bytes = outputs["snapshot_after"].read_bytes()
    assert before_bytes != after_bytes, "Snapshots should differ after formatting"

    # Assert PNG magic bytes on both snapshots
    for key in ("snapshot_before", "snapshot_after"):
        with open(outputs[key], "rb") as f:
            assert f.read(8) == b"\x89PNG\r\n\x1a\n"
