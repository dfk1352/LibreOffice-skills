"""Integration workflow tests for Writer tools and session workflows."""

# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from pathlib import Path

from tests.writer._helpers import (
    create_test_image,
    get_graphic_names,
    get_table_names,
)


def run_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build Writer documents that exercise every public Writer tool."""
    from writer import open_writer_session, patch
    from writer.core import create_document, export_document
    from writer.formatting import apply_formatting
    from writer.metadata import get_metadata, set_metadata
    from writer.snapshot import snapshot_page

    output_dir.mkdir(parents=True, exist_ok=True)

    session_doc = output_dir / "session_workflow.odt"
    create_document(str(session_doc))
    set_metadata(
        str(session_doc),
        {
            "title": "Writer Session Workflow",
            "author": "Test Suite",
            "subject": "Session and patch coverage",
        },
    )
    metadata = get_metadata(str(session_doc))
    assert metadata["title"] == "Writer Session Workflow"
    assert metadata["author"] == "Test Suite"

    logo_v1 = create_test_image(output_dir / "logo_v1.png", color="blue")
    logo_v2 = create_test_image(output_dir / "logo_v2.png", color="green")
    logo_patch = create_test_image(output_dir / "logo_patch.png", color="black")

    with open_writer_session(str(session_doc)) as session:
        session.insert_text("Introduction\n\nBody paragraph.\n\nRemove this sentence.")
        assert "Introduction" in session.read_text()
        assert "Body paragraph." in session.read_text(
            selector='contains:"Body paragraph."'
        )

        session.insert_text(
            "\n\nInserted after intro.", selector='after:"Introduction"'
        )
        session.replace_text('contains:"Body paragraph."', "Updated body paragraph.")
        session.delete_text('contains:"Remove this sentence."')

        session.insert_table(
            2,
            2,
            [["Quarter", "Revenue"], ["Q1", "10"]],
            "Quarterly Results",
        )
        session.insert_table(1, 1, [["discard"]], "Temporary Table")
        session.update_table(
            'name:"Quarterly Results"',
            [["Quarter", "Revenue"], ["Q1", "25"]],
        )
        session.delete_table('name:"Temporary Table"')

        session.insert_image(
            str(logo_v1),
            width=2000,
            height=2000,
            name="Logo",
        )
        session.insert_image(
            str(logo_patch), width=1800, height=1800, name="Disposable"
        )
        session.update_image(
            'name:"Logo"',
            image_path=str(logo_v2),
            width=3500,
            height=2500,
        )
        session.delete_image('name:"Disposable"')
        session.patch(
            "[operation]\n"
            "type = insert_text\n"
            'selector = after:"Updated body paragraph."\n'
            "text = Patched paragraph.\n"
            "[operation]\n"
            "type = insert_image\n"
            "image_path = "
            f"{logo_patch}\n"
            "width = 2200\n"
            "height = 2200\n"
            "name = Patch Artifact\n"
            "[operation]\n"
            "type = insert_text\n"
            "text = Closing note added by patch.\n",
            mode="atomic",
        )

    snapshot_before = output_dir / "writer_snapshot_before.png"
    snapshot_page(str(session_doc), str(snapshot_before), page=1)

    apply_formatting(
        str(session_doc),
        {"italic": True, "color": 0x003366, "font_size": 14},
        selection="all",
    )
    apply_formatting(
        str(session_doc),
        {"bold": True, "align": "center"},
        selection="last_paragraph",
    )

    snapshot_after = output_dir / "writer_snapshot_after.png"
    snapshot_page(str(session_doc), str(snapshot_after), page=1)

    session_export = output_dir / "session_workflow.pdf"
    with open_writer_session(str(session_doc)) as session:
        session.export(str(session_export), "pdf")

    atomic_doc = output_dir / "patch_atomic.odt"
    create_document(str(atomic_doc))
    with open_writer_session(str(atomic_doc)) as session:
        session.insert_text("Executive Summary\n\nLegacy sentence.")

    patch(
        str(atomic_doc),
        "[operation]\n"
        "type = insert_text\n"
        'selector = after:"Executive Summary"\n'
        "text = Atomic addition.\n"
        "[operation]\n"
        "type = replace_text\n"
        'selector = contains:"Legacy sentence."\n'
        "new_text = Modern sentence.\n",
        mode="atomic",
    )

    atomic_export = output_dir / "patch_atomic.docx"
    export_document(str(atomic_doc), str(atomic_export), "docx")

    best_effort_doc = output_dir / "patch_best_effort.odt"
    create_document(str(best_effort_doc))
    with open_writer_session(str(best_effort_doc)) as session:
        session.insert_text("Status Report\n\nKeep this sentence.\n\nTail line.")

    patch(
        str(best_effort_doc),
        "[operation]\n"
        "type = insert_text\n"
        'selector = after:"Status Report"\n'
        "text = Best effort addition.\n"
        "[operation]\n"
        "type = delete_text\n"
        'selector = contains:"missing sentence"\n'
        "[operation]\n"
        "type = replace_text\n"
        'selector = contains:"Tail line."\n'
        "new_text = Tail line updated.\n",
        mode="best_effort",
    )

    return {
        "session_workflow": session_doc,
        "snapshot_before": snapshot_before,
        "snapshot_after": snapshot_after,
        "session_export": session_export,
        "patch_atomic": atomic_doc,
        "patch_atomic_export": atomic_export,
        "patch_best_effort": best_effort_doc,
    }


def test_session_workflow_document_state(tmp_path):
    """Session workflow leaves visible text, tables, images, and metadata."""
    from writer import open_writer_session
    from writer.metadata import get_metadata

    outputs = run_end_to_end_workflow(tmp_path)

    with open_writer_session(str(outputs["session_workflow"])) as session:
        text = session.read_text()

    metadata = get_metadata(str(outputs["session_workflow"]))

    assert "Inserted after intro." in text
    assert "Updated body paragraph." in text
    assert "Patched paragraph." in text
    assert "Closing note added by patch." in text
    assert "Remove this sentence." not in text
    assert get_table_names(outputs["session_workflow"]) == ["Quarterly_Results"]
    assert sorted(get_graphic_names(outputs["session_workflow"])) == [
        "Logo",
        "Patch Artifact",
    ]
    assert metadata["title"] == "Writer Session Workflow"
    assert metadata["author"] == "Test Suite"


def test_patch_workflow_documents_capture_atomic_and_best_effort_results(tmp_path):
    """Standalone patch workflow preserves the intended document outcomes."""
    from writer import open_writer_session

    outputs = run_end_to_end_workflow(tmp_path)

    with open_writer_session(str(outputs["patch_atomic"])) as session:
        atomic_text = session.read_text()

    with open_writer_session(str(outputs["patch_best_effort"])) as session:
        best_effort_text = session.read_text()

    assert "Atomic addition." in atomic_text
    assert "Executive Summary\nAtomic addition." in atomic_text
    assert "Modern sentence." in atomic_text
    assert "Legacy sentence." not in atomic_text

    assert "Best effort addition." in best_effort_text
    assert "Status Report\nBest effort addition." in best_effort_text
    assert "Tail line updated." in best_effort_text
    assert "Keep this sentence." in best_effort_text


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable workflow output files in test-output/writer/."""
    output_dir = Path("test-output/writer")

    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_end_to_end_workflow(output_dir)

    for key in (
        "session_workflow",
        "snapshot_before",
        "snapshot_after",
        "session_export",
        "patch_atomic",
        "patch_atomic_export",
        "patch_best_effort",
    ):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    before_bytes = outputs["snapshot_before"].read_bytes()
    after_bytes = outputs["snapshot_after"].read_bytes()
    assert before_bytes != after_bytes, "Snapshots should differ after formatting"

    for key in ("snapshot_before", "snapshot_after"):
        with open(outputs[key], "rb") as handle:
            assert handle.read(8) == b"\x89PNG\r\n\x1a\n"
