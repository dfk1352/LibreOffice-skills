# pyright: reportMissingImports=false, reportAttributeAccessIssue=false

from __future__ import annotations

from pathlib import Path

from tests.impress._helpers import (
    BLANK_LAYOUT,
    TITLE_ONLY_LAYOUT,
    TITLE_AND_CONTENT_LAYOUT,
    append_slide,
    create_test_audio,
    create_test_image,
    create_test_video,
    get_chart_details,
    get_list_paragraphs,
    get_master_background,
    get_media_url,
    get_notes_text,
    get_shape_text,
    get_slide_master_name,
    get_slide_shapes,
    get_table_matrix,
    get_text_properties,
    open_impress_doc,
)


def run_end_to_end_workflow(output_dir: Path) -> dict[str, Path]:
    """Build Impress workflow artifacts that exercise every public tool."""
    from impress import (
        ImpressSession,
        ImpressTarget,
        ListItem,
        ShapePlacement,
        TextFormatting,
        patch,
        snapshot_slide,
    )
    from impress.core import create_presentation

    output_dir.mkdir(parents=True, exist_ok=True)

    session_doc = output_dir / "session_workflow.odp"
    create_presentation(str(session_doc))

    logo_v1 = create_test_image(output_dir / "logo_v1.png", color="blue")
    logo_v2 = create_test_image(output_dir / "logo_v2.png", color="green")
    audio_path = create_test_audio(output_dir / "sample.wav")
    video_path = create_test_video(output_dir / "sample.mp4")
    template_path = output_dir / "template.odp"
    create_presentation(str(template_path))
    with open_impress_doc(template_path) as template_doc:
        append_slide(template_doc, TITLE_AND_CONTENT_LAYOUT)
        template_doc.store()

    snapshot_before = output_dir / "snapshot_before.png"
    snapshot_after = output_dir / "snapshot_after.png"
    workflow_pdf = output_dir / "presentation.pdf"
    workflow_pptx = output_dir / "presentation.pptx"

    with ImpressSession(str(session_doc)) as session:
        assert session.get_slide_count() == 1
        session.add_slide(layout="TITLE_AND_CONTENT")
        session.add_slide(layout="BLANK")
        session.add_slide(layout="TITLE_ONLY")
        session.add_slide(index=2, layout="BLANK")

        session.insert_text(
            "Executive Summary",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="title"),
        )
        session.insert_text(
            "Quarterly revenue rose 18%.",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="body"),
        )
        assert (
            session.read_text(
                ImpressTarget(kind="text", slide_index=1, placeholder="body")
            )
            == "Quarterly revenue rose 18%."
        )
        session.replace_text(
            ImpressTarget(kind="text", slide_index=1, placeholder="body"),
            "Quarterly revenue rose 21%.",
        )

        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=2),
            "Action Items\nRemove this line",
            ShapePlacement(1.0, 1.0, 7.5, 4.0),
            name="Agenda Box",
        )
        session.delete_item(
            ImpressTarget(
                kind="text",
                slide_index=2,
                shape_name="Agenda Box",
                text="Remove this line",
            )
        )
        session.insert_text(
            "\nStatus update",
            ImpressTarget(
                kind="insertion",
                slide_index=2,
                shape_name="Agenda Box",
                after="Action Items",
            ),
        )
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=2),
            "Priority Update",
            ShapePlacement(9.0, 1.0, 4.0, 1.8),
            name="Callout Box",
        )
        session.format_text(
            ImpressTarget(kind="text", slide_index=2, shape_name="Callout Box"),
            TextFormatting(
                bold=True,
                italic=True,
                font_name="Liberation Sans Narrow",
                font_size=18,
                color="red",
                align="center",
            ),
        )

        session.insert_list(
            [
                ListItem(text="Confirm scope", level=0),
                ListItem(text="Review outputs", level=0),
            ],
            ordered=False,
            target=ImpressTarget(
                kind="insertion",
                slide_index=2,
                shape_name="Agenda Box",
                after="Status update",
            ),
        )
        session.replace_list(
            ImpressTarget(
                kind="list",
                slide_index=2,
                shape_name="Agenda Box",
                text="Confirm scope",
            ),
            [
                ListItem(text="Publish package", level=0),
                ListItem(text="Collect approvals", level=1),
            ],
            ordered=True,
        )
        session.delete_item(
            ImpressTarget(
                kind="list",
                slide_index=2,
                shape_name="Agenda Box",
                text="Publish package",
            )
        )
        session.insert_list(
            [
                ListItem(text="Escalate blocker", level=0),
                ListItem(text="Collect approvals", level=1),
            ],
            ordered=False,
            target=ImpressTarget(
                kind="insertion",
                slide_index=2,
                shape_name="Agenda Box",
                after="Status update",
            ),
        )

        session.insert_shape(
            ImpressTarget(kind="slide", slide_index=2),
            "rectangle",
            ShapePlacement(13.0, 1.0, 4.0, 2.0),
            fill_color="navy",
            line_color="black",
            name="Accent Shape",
        )
        session.delete_item(
            ImpressTarget(kind="shape", slide_index=2, shape_name="Accent Shape")
        )
        session.insert_shape(
            ImpressTarget(kind="slide", slide_index=2),
            "ellipse",
            ShapePlacement(14.0, 1.0, 3.2, 1.8),
            fill_color="gold",
            line_color="black",
            name="Badge Shape",
        )

        session.insert_image(
            ImpressTarget(kind="slide", slide_index=2),
            str(logo_v1),
            ShapePlacement(14.0, 3.2, 3.5, 3.5),
            name="Logo",
        )
        session.insert_image(
            ImpressTarget(kind="slide", slide_index=2),
            str(logo_v1),
            ShapePlacement(17.8, 3.2, 1.8, 1.8),
            name="Disposable Logo",
        )
        session.replace_image(
            ImpressTarget(kind="image", slide_index=2, shape_name="Logo"),
            image_path=str(logo_v2),
            placement=ShapePlacement(14.0, 3.3, 4.0, 4.0),
        )
        session.delete_item(
            ImpressTarget(kind="image", slide_index=2, shape_name="Disposable Logo")
        )

        session.insert_table(
            ImpressTarget(kind="slide", slide_index=2),
            3,
            2,
            ShapePlacement(1.0, 6.2, 11.0, 3.8),
            data=[["Metric", "Value"], ["Revenue", "100"], ["Cost", "80"]],
            name="Metrics Table",
        )
        session.update_table(
            ImpressTarget(kind="table", slide_index=2, shape_name="Metrics Table"),
            [["Metric", "Value"], ["Revenue", "120"], ["Cost", "90"]],
        )
        session.insert_table(
            ImpressTarget(kind="slide", slide_index=2),
            1,
            1,
            ShapePlacement(11.5, 10.0, 2.5, 1.5),
            data=[["X"]],
            name="Disposable Table",
        )
        session.delete_item(
            ImpressTarget(kind="table", slide_index=2, shape_name="Disposable Table")
        )

        session.insert_chart(
            ImpressTarget(kind="slide", slide_index=4),
            "bar",
            [["Category", "Value"], ["Revenue", 100], ["Cost", 80]],
            ShapePlacement(1.0, 1.0, 10.0, 6.0),
            title="Revenue Trend",
            name="Revenue Chart",
        )
        session.insert_chart(
            ImpressTarget(kind="slide", slide_index=4),
            "line",
            [["Category", "Value"], ["Temp", 1]],
            ShapePlacement(12.0, 1.0, 6.0, 4.0),
            title="Disposable Chart",
            name="Disposable Chart",
        )
        session.update_chart(
            ImpressTarget(kind="chart", slide_index=4, shape_name="Revenue Chart"),
            chart_type="line",
            data=[["Category", "Value"], ["Revenue", 110], ["Cost", 75]],
            placement=ShapePlacement(2.0, 1.5, 12.0, 7.0),
            title="Revenue Trend Updated",
        )
        session.delete_item(
            ImpressTarget(kind="chart", slide_index=4, shape_name="Disposable Chart")
        )

        session.insert_media(
            ImpressTarget(kind="slide", slide_index=4),
            str(video_path),
            ShapePlacement(1.0, 9.0, 5.0, 4.0),
            name="Demo Media",
        )
        session.insert_media(
            ImpressTarget(kind="slide", slide_index=4),
            str(video_path),
            ShapePlacement(6.5, 9.0, 4.0, 3.0),
            name="Disposable Media",
        )
        session.replace_media(
            ImpressTarget(kind="media", slide_index=4, shape_name="Demo Media"),
            media_path=str(audio_path),
            placement=ShapePlacement(1.5, 9.5, 5.5, 3.5),
        )
        session.delete_item(
            ImpressTarget(kind="media", slide_index=4, shape_name="Disposable Media")
        )

        imported_master = session.import_master_page(str(template_path))
        imported_master_name = imported_master.master_name
        session.set_master_background(
            ImpressTarget(kind="master_page", master_name=imported_master_name),
            "lightsteelblue",
        )
        session.apply_master_page(
            ImpressTarget(kind="slide", slide_index=2),
            ImpressTarget(kind="master_page", master_name=imported_master_name),
        )

        inventory = session.get_slide_inventory(
            ImpressTarget(kind="slide", slide_index=2)
        )
        assert inventory["shape_count"] >= 3

        session.duplicate_slide(ImpressTarget(kind="slide", slide_index=0))
        session.move_slide(ImpressTarget(kind="slide", slide_index=1), 3)
        session.delete_slide(ImpressTarget(kind="slide", slide_index=4))
        assert session.get_slide_count() == 5

        session.set_notes(
            ImpressTarget(kind="notes", slide_index=4),
            "Remember to greet the audience.",
        )
        assert (
            "greet"
            in session.get_notes(ImpressTarget(kind="notes", slide_index=4)).lower()
        )

    snapshot_slide(str(session_doc), 2, str(snapshot_before))

    with ImpressSession(str(session_doc)) as session:
        patch_result = session.patch(
            "[operation]\n"
            "type = insert_text\n"
            "target.kind = insertion\n"
            "target.slide_index = 2\n"
            "target.shape_name = Agenda Box\n"
            "target.after = Status update\n"
            "text <<EOF\n"
            "\nPatched paragraph\n"
            "EOF\n"
            "[operation]\n"
            "type = format_text\n"
            "target.kind = text\n"
            "target.slide_index = 2\n"
            "target.shape_name = Callout Box\n"
            "target.text = Priority Update\n"
            "format.underline = true\n"
            "[operation]\n"
            "type = format_text\n"
            "target.kind = text\n"
            "target.slide_index = 2\n"
            "target.shape_name = Agenda Box\n"
            "target.text = Patched paragraph\n"
            "format.underline = true\n"
            "[operation]\n"
            "type = insert_list\n"
            "target.kind = insertion\n"
            "target.slide_index = 2\n"
            "target.shape_name = Agenda Box\n"
            "target.after = Patched paragraph\n"
            "list.ordered = false\n"
            'items = [{"text": "Patched checklist", "level": 0}]\n',
            mode="atomic",
        )
        assert patch_result.overall_status == "ok", patch_result.operations

        session.export(str(workflow_pdf), "pdf")
        session.export(str(workflow_pptx), "pptx")

    snapshot_slide(str(session_doc), 2, str(snapshot_after))

    atomic_success_doc = output_dir / "patch_atomic_success.odp"
    create_presentation(str(atomic_success_doc))
    with ImpressSession(str(atomic_success_doc)) as session:
        session.add_slide(layout="TITLE_AND_CONTENT")
        session.insert_text(
            "Atomic body",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="body"),
        )

    patch(
        str(atomic_success_doc),
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.slide_index = 1\n"
        "target.placeholder = body\n"
        "new_text = Atomic success text\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Atomic success notes\n",
        mode="atomic",
    )

    atomic_fail_doc = output_dir / "patch_atomic_fail.odp"
    create_presentation(str(atomic_fail_doc))
    with ImpressSession(str(atomic_fail_doc)) as session:
        session.add_slide(layout="TITLE_AND_CONTENT")
        session.insert_text(
            "Atomic body",
            ImpressTarget(kind="insertion", slide_index=1, placeholder="body"),
        )

    patch(
        str(atomic_fail_doc),
        "[operation]\n"
        "type = replace_text\n"
        "target.kind = text\n"
        "target.slide_index = 1\n"
        "target.placeholder = body\n"
        "new_text = rolled back\n"
        "[operation]\n"
        "type = delete_item\n"
        "target.kind = chart\n"
        "target.slide_index = 1\n"
        "target.shape_name = Missing Chart\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = skipped later\n",
        mode="atomic",
    )

    best_effort_doc = output_dir / "patch_best_effort.odp"
    create_presentation(str(best_effort_doc))
    with ImpressSession(str(best_effort_doc)) as session:
        session.add_slide(layout="TITLE_AND_CONTENT")
        session.insert_text_box(
            ImpressTarget(kind="slide", slide_index=1),
            "Action Items",
            ShapePlacement(1.0, 1.0, 10.0, 4.0),
            name="Agenda Box",
        )

    patch(
        str(best_effort_doc),
        "[operation]\n"
        "type = insert_list\n"
        "target.kind = insertion\n"
        "target.slide_index = 1\n"
        "target.shape_name = Agenda Box\n"
        "target.after = Action Items\n"
        "list.ordered = false\n"
        'items = [{"text": "Best effort addition", "level": 0}]\n'
        "[operation]\n"
        "type = replace_media\n"
        "target.kind = media\n"
        "target.slide_index = 1\n"
        "target.shape_name = Missing Media\n"
        f"media_path = {audio_path}\n"
        "[operation]\n"
        "type = set_notes\n"
        "target.kind = notes\n"
        "target.slide_index = 1\n"
        "text = Best effort notes\n",
        mode="best_effort",
    )

    return {
        "session_workflow": session_doc,
        "patch_atomic_success": atomic_success_doc,
        "patch_atomic_fail": atomic_fail_doc,
        "patch_best_effort": best_effort_doc,
        "workflow_pdf": workflow_pdf,
        "workflow_pptx": workflow_pptx,
        "snapshot_before": snapshot_before,
        "snapshot_after": snapshot_after,
    }


def test_session_workflow_document_state(tmp_path):
    """Session workflow leaves visible traces for every major editing domain."""
    outputs = run_end_to_end_workflow(tmp_path)

    assert get_shape_text(outputs["session_workflow"], 1, placeholder="title") == (
        "Executive Summary"
    )
    assert get_shape_text(outputs["session_workflow"], 1, placeholder="body") == (
        "Quarterly revenue rose 21%."
    )

    agenda_text = get_shape_text(outputs["session_workflow"], 2, name="Agenda Box")
    assert "Status update" in agenda_text
    assert "Patched paragraph" in agenda_text
    assert "Remove this line" not in agenda_text
    assert [
        item["text"]
        for item in get_list_paragraphs(
            outputs["session_workflow"], 2, name="Agenda Box"
        )
    ] == ["Patched checklist", "Escalate blocker", "Collect approvals"]

    properties = get_text_properties(outputs["session_workflow"], 2, name="Callout Box")
    assert properties["char_weight"] == 150
    assert properties["char_posture"] == 2
    assert properties["align"] in {2, 3}

    names = [
        shape["name"] for shape in get_slide_shapes(outputs["session_workflow"], 2)
    ]
    assert "Badge_Shape" in names
    assert "Callout_Box" in names
    assert "Logo" in names
    assert "Metrics_Table" in names

    assert "logo_v2" in get_media_url(outputs["session_workflow"], 2, name="Logo")
    assert get_table_matrix(outputs["session_workflow"], 2, name="Metrics Table") == [
        ["Metric", "Value"],
        ["Revenue", "120"],
        ["Cost", "90"],
    ]

    chart = get_chart_details(outputs["session_workflow"], 4, name="Revenue Chart")
    assert chart["title"] == "Revenue Trend Updated"
    assert chart["x"] == 2000
    assert get_media_url(outputs["session_workflow"], 4, name="Demo Media").endswith(
        "sample.wav"
    )
    assert (
        get_notes_text(outputs["session_workflow"], 4)
        == "Remember to greet the audience."
    )

    master_name = get_slide_master_name(outputs["session_workflow"], 2)
    assert master_name
    assert get_master_background(outputs["session_workflow"], master_name) == 0xB0C4DE

    with open_impress_doc(outputs["session_workflow"]) as doc:
        layouts = [
            int(doc.DrawPages.getByIndex(index).Layout)
            for index in range(doc.DrawPages.Count)
        ]
    assert layouts[0] == layouts[3], "Slides 0 and 3 should be identical duplicates"
    assert layouts.count(TITLE_AND_CONTENT_LAYOUT) >= 1
    assert layouts.count(TITLE_ONLY_LAYOUT) == 1
    assert layouts.count(BLANK_LAYOUT) == 1

    for key in (
        "session_workflow",
        "workflow_pdf",
        "workflow_pptx",
        "snapshot_before",
        "snapshot_after",
    ):
        path = outputs[key]
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"


def test_patch_workflow_documents_capture_atomic_and_best_effort_results(tmp_path):
    """Standalone patch workflow preserves atomic rollback and best-effort saves."""
    outputs = run_end_to_end_workflow(tmp_path)

    # Atomic success: all ops applied
    assert (
        get_shape_text(outputs["patch_atomic_success"], 1, placeholder="body")
        == "Atomic success text"
    )
    assert get_notes_text(outputs["patch_atomic_success"], 1) == "Atomic success notes"

    # Atomic fail: rolled back, document unchanged from baseline
    assert (
        get_shape_text(outputs["patch_atomic_fail"], 1, placeholder="body")
        == "Atomic body"
    )
    assert get_notes_text(outputs["patch_atomic_fail"], 1) == ""

    # Best effort: valid ops applied, invalid ops skipped
    assert [
        item["text"]
        for item in get_list_paragraphs(
            outputs["patch_best_effort"], 1, name="Agenda Box"
        )
    ] == ["Best effort addition"]
    assert get_notes_text(outputs["patch_best_effort"], 1) == "Best effort notes"


def test_workflow_outputs_to_test_output_dir():
    """Produce inspectable workflow artifacts in test-output/impress/."""
    output_dir = Path("test-output/impress")

    if output_dir.exists():
        import shutil

        shutil.rmtree(output_dir)

    outputs = run_end_to_end_workflow(output_dir)

    for key, path in outputs.items():
        assert path.exists(), f"{key} not found at {path}"
        assert path.stat().st_size > 0, f"{key} is empty at {path}"

    with open(outputs["snapshot_before"], "rb") as handle:
        assert handle.read(8) == b"\x89PNG\r\n\x1a\n"
    with open(outputs["snapshot_after"], "rb") as handle:
        assert handle.read(8) == b"\x89PNG\r\n\x1a\n"
    with open(outputs["workflow_pdf"], "rb") as handle:
        assert handle.read(5) == b"%PDF-"
