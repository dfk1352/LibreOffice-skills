# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from pathlib import Path

from tests.writer._helpers import (
    ARABIC_NUMBERING_TYPE,
    BULLET_NUMBERING_TYPE,
    assert_list_items,
    assert_text_formatting,
    create_test_image,
    get_graphic_names,
    get_table_names,
)


def run_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build Writer documents that exercise every public Writer editing tool."""
    from writer import (
        ListItem,
        TextFormatting,
        WriterTarget,
        WriterSession,
        patch,
        snapshot_page,
    )
    from writer.core import create_document

    output_dir.mkdir(parents=True, exist_ok=True)

    session_doc = output_dir / "session_workflow.odt"
    create_document(str(session_doc))

    logo_v1 = create_test_image(output_dir / "logo_v1.png", color="blue")
    logo_v2 = create_test_image(output_dir / "logo_v2.png", color="green")

    with WriterSession(str(session_doc)) as session:
        session.insert_text(
            "Executive Summary\n\n"
            "Financial Summary\n\n"
            "Quarterly revenue grew 18%.\n\n"
            "Action Items\n\n"
            "Remove this sentence.\n\n"
            "Risks\n\n"
            "Appendix"
        )
        assert "Executive Summary" in session.read_text()
        assert "Quarterly revenue grew 18%." in session.read_text(
            target=WriterTarget(kind="text", text="Quarterly revenue grew 18%.")
        )

        session.insert_text(
            "Inserted after summary.",
            target=WriterTarget(kind="insertion", after="Executive Summary"),
        )
        session.replace_text(
            WriterTarget(kind="text", text="Quarterly revenue grew 18%."),
            "Quarterly revenue grew 21%.",
        )
        session.delete_text(WriterTarget(kind="text", text="Remove this sentence."))
        session.format_text(
            WriterTarget(
                kind="text",
                text="Quarterly revenue grew 21%.",
                after="Financial Summary",
                before="Action Items",
            ),
            TextFormatting(bold=True, italic=True, align="center"),
        )

        session.insert_table(
            2,
            2,
            [["Quarter", "Revenue"], ["Q1", "10"]],
            "Quarterly Results",
        )
        session.insert_table(1, 1, [["discard"]], "Temporary Table")
        session.update_table(
            WriterTarget(kind="table", name="Quarterly Results"),
            [["Quarter", "Revenue"], ["Q1", "25"]],
        )
        session.delete_table(WriterTarget(kind="table", name="Temporary Table"))

        session.insert_image(
            str(logo_v1),
            width=2000,
            height=2000,
            name="Logo",
        )
        session.insert_image(str(logo_v2), width=1800, height=1800, name="Disposable")
        session.update_image(
            WriterTarget(kind="image", name="Logo"),
            image_path=str(logo_v2),
            width=3500,
            height=2500,
        )
        session.delete_image(WriterTarget(kind="image", name="Disposable"))

        session.insert_list(
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review output", level=0),
                ListItem(text="Update packaging", level=1),
            ],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Action Items"),
        )
        session.replace_list(
            WriterTarget(
                kind="list", text="Confirm scope", after="Action Items", before="Risks"
            ),
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review output", level=0),
                ListItem(text="Publish package", level=1),
            ],
            ordered=True,
        )
        session.delete_list(WriterTarget(kind="list", text="Publish package"))
        session.insert_list(
            [
                ListItem(text="Escalate blocker", level=0),
                ListItem(text="Collect approvals", level=1),
            ],
            ordered=False,
            target=WriterTarget(kind="insertion", after="Risks"),
        )

        session.patch(
            "[operation]\n"
            "type = insert_text\n"
            "target.kind = insertion\n"
            "target.after = Risks\n"
            "text = Patched paragraph.\n"
            "[operation]\n"
            "type = format_text\n"
            "target.kind = text\n"
            "target.text = Patched paragraph.\n"
            "format.underline = true\n"
            "[operation]\n"
            "type = insert_list\n"
            "target.kind = insertion\n"
            "target.after = Patched paragraph.\n"
            "list.ordered = false\n"
            'items = [{"text": "Patched checklist", "level": 0}]\n',
            mode="atomic",
        )

    atomic_baseline_text = (
        "Executive Summary\n\n"
        "Financial Summary\n\n"
        "Quarterly revenue grew 18%.\n\n"
        "Action Items"
    )

    atomic_success_doc = output_dir / "patch_atomic_success.odt"
    create_document(str(atomic_success_doc))
    with WriterSession(str(atomic_success_doc)) as session:
        session.insert_text(atomic_baseline_text)

    patch(
        str(atomic_success_doc),
        "[operation]\n"
        "type = insert_text\n"
        "target.kind = insertion\n"
        "target.after = Executive Summary\n"
        "text = Atomic addition.\n"
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Quarterly revenue grew 18%.\n"
        "format.bold = true\n"
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Action Items\n"
        "list.ordered = true\n"
        'items = [{"text": "Confirm scope", "level": 0}, {"text": "Review output", "level": 1}]\n',
        mode="atomic",
    )

    atomic_fail_doc = output_dir / "patch_atomic_fail.odt"
    create_document(str(atomic_fail_doc))
    with WriterSession(str(atomic_fail_doc)) as session:
        session.insert_text(atomic_baseline_text)

    patch(
        str(atomic_fail_doc),
        "[operation]\n"
        "type = insert_text\n"
        "target.kind = insertion\n"
        "target.after = Executive Summary\n"
        "text = Atomic addition.\n"
        "[operation]\n"
        "type = delete_table\n"
        "target.kind = table\n"
        "target.name = MissingTable\n"
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Quarterly revenue grew 18%.\n"
        "format.bold = true\n",
        mode="atomic",
    )

    best_effort_doc = output_dir / "patch_best_effort.odt"
    create_document(str(best_effort_doc))
    with WriterSession(str(best_effort_doc)) as session:
        session.insert_text(
            "Status Report\n\n"
            "Financial Summary\n\n"
            "Quarterly revenue grew 18%.\n\n"
            "Action Items\n\n"
            "Tail line."
        )

    patch(
        str(best_effort_doc),
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.after = Action Items\n"
        "list.ordered = false\n"
        'items = [{"text": "Best effort addition", "level": 0}]\n'
        "[operation]\n"
        "type = replace_list\n"
        "target.kind = list\n"
        "target.text = missing sentence\n"
        'items = [{"text": "No change", "level": 0}]\n'
        "[operation]\n"
        "type = format_text\n"
        "target.kind = text\n"
        "target.text = Tail line.\n"
        "format.underline = true\n",
        mode="best_effort",
    )

    session_snapshot = output_dir / "session_workflow_page1.png"
    snapshot_page(str(session_doc), str(session_snapshot), page=1)

    md_source = output_dir / "import_source.md"
    md_source.write_text(
        "# Imported Report\n\n"
        "This document was created from a Markdown file.\n\n"
        "## Details\n\n"
        "- First point\n"
        "- Second point\n"
    )
    md_imported = output_dir / "md_imported.odt"
    create_document(str(md_imported), source=str(md_source))

    md_exported = output_dir / "session_workflow.md"
    with WriterSession(str(session_doc)) as session:
        session.export(str(md_exported), "md")

    return {
        "session_workflow": session_doc,
        "patch_atomic_success": atomic_success_doc,
        "patch_atomic_fail": atomic_fail_doc,
        "patch_best_effort": best_effort_doc,
        "session_snapshot": session_snapshot,
        "md_imported": md_imported,
        "md_exported": md_exported,
    }


def test_session_workflow_document_state(tmp_path):
    """Session workflow leaves visible text, tables, images, formatting, and lists."""
    from writer import WriterSession

    outputs = run_end_to_end_workflow(tmp_path)

    with WriterSession(str(outputs["session_workflow"])) as session:
        text = session.read_text()

    assert "Executive Summary\nInserted after summary.\n\nFinancial Summary" in text
    assert "Quarterly revenue grew 21%." in text
    assert "Patched paragraph." in text
    assert "Remove this sentence." not in text
    assert get_table_names(outputs["session_workflow"]) == ["Quarterly_Results"]
    assert get_graphic_names(outputs["session_workflow"]) == ["Logo"]
    assert_text_formatting(
        outputs["session_workflow"],
        "Quarterly revenue grew 21%.",
        char_weight=150.0,
        align=3,
    )
    assert_list_items(
        outputs["session_workflow"],
        ["Escalate blocker", "Collect approvals", "Patched checklist"],
        expected_levels=[0, 1, 0],
        expected_numbering_type=BULLET_NUMBERING_TYPE,
    )


def test_patch_workflow_documents_capture_atomic_and_best_effort_results(tmp_path):
    """Standalone patch workflow preserves the intended document outcomes."""
    from writer import WriterSession

    outputs = run_end_to_end_workflow(tmp_path)

    with WriterSession(str(outputs["patch_atomic_success"])) as session:
        atomic_success_text = session.read_text()

    with WriterSession(str(outputs["patch_atomic_fail"])) as session:
        atomic_fail_text = session.read_text()

    with WriterSession(str(outputs["patch_best_effort"])) as session:
        best_effort_text = session.read_text()

    # Atomic success: all ops applied
    assert "Atomic addition." in atomic_success_text
    assert_text_formatting(
        outputs["patch_atomic_success"],
        "Quarterly revenue grew 18%.",
        char_weight=150.0,
    )
    assert_list_items(
        outputs["patch_atomic_success"],
        ["Confirm scope", "Review output"],
        expected_levels=[0, 1],
        expected_numbering_type=ARABIC_NUMBERING_TYPE,
    )

    # Atomic fail: rolled back, document unchanged from baseline
    assert "Atomic addition." not in atomic_fail_text
    assert "Executive Summary" in atomic_fail_text
    assert "Quarterly revenue grew 18%." in atomic_fail_text

    # Best effort: valid ops applied, invalid ops skipped
    assert "Best effort addition" in best_effort_text
    assert "Tail line." in best_effort_text
    assert_list_items(
        outputs["patch_best_effort"],
        ["Best effort addition"],
        expected_levels=[0],
        expected_numbering_type=BULLET_NUMBERING_TYPE,
    )
    assert_text_formatting(outputs["patch_best_effort"], "Tail line.", char_underline=1)


def test_markdown_workflow_document_state(tmp_path):
    """Markdown import creates a readable ODT; markdown export captures text."""
    from writer import WriterSession

    outputs = run_end_to_end_workflow(tmp_path)

    with WriterSession(str(outputs["md_imported"])) as session:
        imported_text = session.read_text()

    assert "Imported Report" in imported_text
    assert "created from a Markdown file" in imported_text

    md_content = outputs["md_exported"].read_text()
    assert "Executive Summary" in md_content
    assert "Quarterly revenue grew 21%." in md_content


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable workflow output files in test-output/writer/."""
    output_dir = Path("test-output/writer")

    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_end_to_end_workflow(output_dir)

    for key in (
        "session_workflow",
        "patch_atomic_success",
        "patch_atomic_fail",
        "patch_best_effort",
        "session_snapshot",
        "md_imported",
        "md_exported",
    ):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    with open(outputs["session_snapshot"], "rb") as handle:
        assert handle.read(8) == b"\x89PNG\r\n\x1a\n"
